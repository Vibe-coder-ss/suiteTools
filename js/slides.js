/**
 * Slides Editor Module
 * Handles slides/presentation creation, viewing, and export
 */

const SlidesEditor = {
    // XML namespace resolver for OOXML
    nsResolver(prefix) {
        const ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        };
        return ns[prefix] || null;
    },

    // Create empty presentation
    async createEmpty() {
        const state = AppState;
        
        if (state.hasUnsavedChanges || state.docHasUnsavedChanges || state.slidesHasUnsavedChanges) {
            const confirmed = await Utils.confirm('You have unsaved changes. Creating a new presentation will clear them. Continue?', {
                title: 'Unsaved Changes',
                icon: '⚠️',
                okText: 'Continue',
                cancelText: 'Cancel',
                danger: true
            });
            if (!confirmed) return;
        }
        
        Spreadsheet.clearSilent();
        
        state.currentMode = 'slides';
        state.slidesFile = { name: 'new_presentation.pptx', size: 0 };
        state.slidesData = [{
            id: 1,
            title: 'Welcome',
            subtitle: 'Click to add subtitle',
            content: '',
            layout: 'title',
            background: '#667eea',
            backgroundImage: null,
            images: [],
            notes: ''
        }];
        state.currentSlideIndex = 0;
        state.slidesHasUnsavedChanges = true;
        state.slidesFileHandle = null;
        
        this.display(state.slidesFile);
        
        Utils.showToast('New presentation created! Use the sidebar to navigate slides.', 'success');
    },

    // Parse PPTX or PPT file
    async parse(arrayBuffer, file) {
        const state = AppState;
        const ext = file.name.split('.').pop().toLowerCase();
        
        // Handle legacy .ppt format
        if (ext === 'ppt') {
            return this.parseLegacyPpt(arrayBuffer, file);
        }
        
        try {
            state.slidesFile = file;
            state.currentMode = 'slides';
            state.slidesHasUnsavedChanges = false;
            
            // Use JSZip to extract PPTX contents
            const zip = await JSZip.loadAsync(arrayBuffer);
            
            // Parse slides from the XML files
            const slides = await this.extractSlidesFromZip(zip);
            
            if (slides.length === 0) {
                // Fall back to a placeholder if no slides found
                slides.push({
                    id: 1,
                    title: 'Imported Presentation',
                    subtitle: file.name,
                    content: 'Slide content could not be fully parsed.',
                    layout: 'title',
                    background: '#ffffff',
                    notes: ''
                });
            }
            
            state.slidesData = slides;
            state.currentSlideIndex = 0;
            
            this.display(file);
            
        } catch (error) {
            console.error('Error parsing presentation:', error);
            Utils.alert('Error reading presentation file. Some content may not be displayable.', {
                title: 'Import Error',
                icon: '⚠️'
            });
            
            // Still try to show something
            const state = AppState;
            state.slidesData = [{
                id: 1,
                title: file.name,
                subtitle: 'Import Error',
                content: 'The presentation could not be fully parsed. Try exporting to create a new compatible file.',
                layout: 'title',
                background: '#ffffff',
                backgroundImage: null,
                images: [],
                notes: ''
            }];
            state.currentSlideIndex = 0;
            state.slidesFile = file;
            state.currentMode = 'slides';
            this.display(file);
        }
    },

    // Parse legacy .ppt format (limited support)
    async parseLegacyPpt(arrayBuffer, file) {
        const state = AppState;
        state.slidesFile = file;
        state.currentMode = 'slides';
        state.slidesHasUnsavedChanges = false;
        
        // Try to extract text from PPT using a basic approach
        // PPT files are OLE compound documents with embedded text
        try {
            const bytes = new Uint8Array(arrayBuffer);
            const textContent = this.extractTextFromPpt(bytes);
            
            // Create slides from extracted text
            const slides = [];
            
            if (textContent && textContent.length > 0) {
                // Split text into potential slides (by double newlines or significant breaks)
                const chunks = textContent.split(/\n\s*\n/).filter(c => c.trim());
                
                chunks.forEach((chunk, index) => {
                    const lines = chunk.split('\n').map(l => l.trim()).filter(l => l);
                    slides.push({
                        id: index + 1,
                        title: lines[0] || `Slide ${index + 1}`,
                        subtitle: lines[1] || '',
                        content: lines.slice(2).join('\n'),
                        layout: lines.length <= 2 ? 'title' : 'content',
                        background: '#ffffff',
                        backgroundImage: null,
                        images: [],
                        notes: ''
                    });
                });
            }
            
            if (slides.length === 0) {
                slides.push({
                    id: 1,
                    title: file.name.replace('.ppt', ''),
                    subtitle: 'Legacy PPT Format',
                    content: 'This is a legacy .ppt file. Text content has been extracted where possible.\n\nFor full visual fidelity, please convert to .pptx format using Microsoft PowerPoint or LibreOffice.',
                    layout: 'title',
                    background: '#4a5568',
                    backgroundImage: null,
                    images: [],
                    notes: ''
                });
            }
            
            state.slidesData = slides;
            state.currentSlideIndex = 0;
            
            this.display(file);
            Utils.showToast('Legacy .ppt file loaded. Some formatting may not be preserved. Consider saving as .pptx for full support.', 'warning');
            
        } catch (error) {
            console.error('Error parsing PPT:', error);
            
            state.slidesData = [{
                id: 1,
                title: file.name.replace('.ppt', ''),
                subtitle: 'Legacy Format - Limited Support',
                content: 'This legacy .ppt file could not be fully parsed.\n\nPlease convert to .pptx format using:\n• Microsoft PowerPoint\n• LibreOffice Impress\n• Google Slides',
                layout: 'title',
                background: '#4a5568',
                backgroundImage: null,
                images: [],
                notes: ''
            }];
            state.currentSlideIndex = 0;
            
            this.display(file);
            Utils.showToast('Legacy .ppt format has limited support. Please convert to .pptx for best results.', 'warning');
        }
    },

    // Extract text content from PPT binary
    extractTextFromPpt(bytes) {
        const texts = [];
        
        // Look for text in the binary - PPT stores text in various records
        // This is a simplified extraction that looks for readable ASCII/Unicode text
        let currentText = '';
        let inText = false;
        
        for (let i = 0; i < bytes.length - 1; i++) {
            const byte = bytes[i];
            const nextByte = bytes[i + 1];
            
            // Check for UTF-16LE text (common in PPT)
            if (byte >= 32 && byte < 127 && nextByte === 0) {
                currentText += String.fromCharCode(byte);
                inText = true;
                i++; // Skip null byte
            } 
            // Check for ASCII text
            else if (byte >= 32 && byte < 127) {
                currentText += String.fromCharCode(byte);
                inText = true;
            }
            // Check for newline
            else if ((byte === 13 || byte === 10) && inText) {
                if (currentText.length >= 3) {
                    currentText += '\n';
                }
            }
            // End of text block
            else if (inText && currentText.length >= 3) {
                // Filter out binary garbage and common non-content strings
                const cleaned = currentText.trim();
                if (cleaned.length >= 3 && 
                    !cleaned.match(/^[A-Z]{2,}$/) && // Skip all-caps short codes
                    !cleaned.match(/^\d+$/) && // Skip pure numbers
                    !cleaned.includes('Microsoft') &&
                    !cleaned.includes('PowerPoint') &&
                    !cleaned.includes('www.') &&
                    !cleaned.includes('http') &&
                    !cleaned.match(/^[_\-\.]+$/) && // Skip separator strings
                    cleaned.match(/[a-zA-Z]/) // Must contain at least one letter
                ) {
                    texts.push(cleaned);
                }
                currentText = '';
                inText = false;
            } else {
                currentText = '';
                inText = false;
            }
        }
        
        // Add final text if any
        if (currentText.trim().length >= 3) {
            texts.push(currentText.trim());
        }
        
        // Remove duplicates and join
        const uniqueTexts = [...new Set(texts)];
        return uniqueTexts.join('\n\n');
    },

    // Extract slides from ZIP structure
    async extractSlidesFromZip(zip) {
        const slides = [];
        const slideFiles = [];
        
        // Find all slide XML files
        zip.forEach((relativePath, zipEntry) => {
            if (relativePath.match(/ppt\/slides\/slide\d+\.xml$/)) {
                slideFiles.push({ path: relativePath, entry: zipEntry });
            }
        });
        
        // Sort by slide number
        slideFiles.sort((a, b) => {
            const numA = parseInt(a.path.match(/slide(\d+)\.xml$/)[1]);
            const numB = parseInt(b.path.match(/slide(\d+)\.xml$/)[1]);
            return numA - numB;
        });
        
        // Try to get theme colors from theme1.xml
        let themeColors = await this.extractThemeColors(zip);
        
        // Try to get master slide background
        let masterBackground = await this.extractMasterBackground(zip);
        
        // Extract media files (images) for backgrounds
        const mediaFiles = await this.extractMediaFiles(zip);
        
        // Also extract relationships from slide masters and layouts
        const masterRels = await this.extractMasterRelationships(zip);
        // Merge master rels into mediaFiles if they reference media
        for (const [rId, info] of Object.entries(masterRels)) {
            if (info.mediaUrl) {
                mediaFiles[`master_${rId}`] = info.mediaUrl;
            }
        }
        
        // Get slide layout backgrounds
        const layoutBackgrounds = await this.extractLayoutBackgrounds(zip);
        
        // Parse each slide
        for (let i = 0; i < slideFiles.length; i++) {
            try {
                const content = await slideFiles[i].entry.async('string');
                const slideNum = parseInt(slideFiles[i].path.match(/slide(\d+)\.xml$/)[1]);
                
                // Get slide relationships for images
                const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
                let slideRels = {};
                let layoutRef = null;
                if (zip.file(relsPath)) {
                    const relsContent = await zip.file(relsPath).async('string');
                    slideRels = this.parseRelationships(relsContent);
                    
                    // Find layout reference
                    const layoutMatch = relsContent.match(/Target="\.\.\/slideLayouts\/slideLayout(\d+)\.xml"/i);
                    if (layoutMatch) {
                        layoutRef = parseInt(layoutMatch[1]);
                    }
                }
                
                // Get layout-specific background if available
                const layoutBg = layoutRef && layoutBackgrounds[layoutRef] ? layoutBackgrounds[layoutRef] : null;
                const effectiveMasterBg = layoutBg || masterBackground;
                
                const slide = await this.parseSlideXML(content, i + 1, themeColors, effectiveMasterBg, slideRels, mediaFiles);
                slides.push(slide);
            } catch (err) {
                console.warn(`Failed to parse slide ${i + 1}:`, err);
                let fallbackBg = '#ffffff';
                if (masterBackground && masterBackground.color) {
                    fallbackBg = typeof masterBackground.color === 'string' ? masterBackground.color : '#ffffff';
                }
                slides.push({
                    id: i + 1,
                    title: `Slide ${i + 1}`,
                    subtitle: '',
                    content: 'Could not parse slide content',
                    layout: 'content',
                    background: fallbackBg,
                    backgroundImage: null,
                    images: [],
                    notes: ''
                });
            }
        }
        
        return slides;
    },

    // Extract theme colors from theme XML
    async extractThemeColors(zip) {
        const colors = {
            dk1: '#000000',
            lt1: '#ffffff',
            dk2: '#1f497d',
            lt2: '#eeece1',
            accent1: '#4f81bd',
            accent2: '#c0504d',
            accent3: '#9bbb59',
            accent4: '#8064a2',
            accent5: '#4bacc6',
            accent6: '#f79646',
            bg1: '#ffffff',
            bg2: '#eeece1',
            tx1: '#000000',
            tx2: '#1f497d'
        };
        
        try {
            const themeFile = zip.file('ppt/theme/theme1.xml');
            if (themeFile) {
                const themeXml = await themeFile.async('string');
                
                // Use regex to extract colors more reliably
                for (const colorName of Object.keys(colors)) {
                    // Look for <a:colorName> with srgbClr or sysClr
                    const srgbMatch = themeXml.match(new RegExp(`<a:${colorName}[^>]*>\\s*<a:srgbClr\\s+val="([A-Fa-f0-9]{6})"`, 'i'));
                    if (srgbMatch) {
                        colors[colorName] = '#' + srgbMatch[1];
                        continue;
                    }
                    
                    const sysClrMatch = themeXml.match(new RegExp(`<a:${colorName}[^>]*>\\s*<a:sysClr[^>]+lastClr="([A-Fa-f0-9]{6})"`, 'i'));
                    if (sysClrMatch) {
                        colors[colorName] = '#' + sysClrMatch[1];
                    }
                }
            }
        } catch (err) {
            console.warn('Could not extract theme colors:', err);
        }
        
        return colors;
    },

    // Extract master slide background
    async extractMasterBackground(zip) {
        try {
            // Try slide master first
            const masterFile = zip.file('ppt/slideMasters/slideMaster1.xml');
            if (masterFile) {
                const masterXml = await masterFile.async('string');
                const bgInfo = this.extractBackgroundFromXml(masterXml);
                if (bgInfo.color || bgInfo.imageRId || bgInfo.gradient) return bgInfo;
            }
            
            // Try slide layout 1 as fallback
            const layoutFile = zip.file('ppt/slideLayouts/slideLayout1.xml');
            if (layoutFile) {
                const layoutXml = await layoutFile.async('string');
                const bgInfo = this.extractBackgroundFromXml(layoutXml);
                if (bgInfo.color || bgInfo.imageRId || bgInfo.gradient) return bgInfo;
            }
        } catch (err) {
            console.warn('Could not extract master background:', err);
        }
        return { color: null, imageRId: null, gradient: null };
    },

    // Extract backgrounds from all slide layouts
    async extractLayoutBackgrounds(zip) {
        const layouts = {};
        
        try {
            // Find all layout files
            const layoutFiles = [];
            zip.forEach((path, entry) => {
                const match = path.match(/ppt\/slideLayouts\/slideLayout(\d+)\.xml$/);
                if (match) {
                    layoutFiles.push({ num: parseInt(match[1]), entry });
                }
            });
            
            for (const { num, entry } of layoutFiles) {
                try {
                    const content = await entry.async('string');
                    const bgInfo = this.extractBackgroundFromXml(content);
                    if (bgInfo.color || bgInfo.imageRId || bgInfo.gradient) {
                        layouts[num] = bgInfo;
                    }
                } catch (err) {
                    // Skip this layout
                }
            }
        } catch (err) {
            console.warn('Could not extract layout backgrounds:', err);
        }
        
        return layouts;
    },

    // Extract background color from XML content using regex (more reliable than DOM for namespaced XML)
    extractBackgroundFromXml(xmlString) {
        let color = null;
        let gradient = null;
        let imageRId = null;
        
        // Method 1: Look for p:bg > p:bgPr > a:solidFill > a:srgbClr
        const solidSrgbMatch = xmlString.match(/<p:bg\b[^>]*>[\s\S]*?<a:srgbClr\s+val="([A-Fa-f0-9]{6})"/i);
        if (solidSrgbMatch) {
            color = '#' + solidSrgbMatch[1];
        }
        
        // Method 2: Look for any solidFill with srgbClr inside bgPr
        if (!color) {
            const bgPrSrgbMatch = xmlString.match(/<p:bgPr\b[^>]*>[\s\S]*?<a:srgbClr\s+val="([A-Fa-f0-9]{6})"/i);
            if (bgPrSrgbMatch) {
                color = '#' + bgPrSrgbMatch[1];
            }
        }
        
        // Method 3: Look for scheme color in background
        if (!color) {
            const schemeMatch = xmlString.match(/<p:bg\b[^>]*>[\s\S]*?<a:schemeClr\s+val="(\w+)"/i);
            if (schemeMatch) {
                color = { scheme: schemeMatch[1] };
            }
        }
        
        // Method 4: Look for bgRef (background reference with scheme color)
        if (!color) {
            const bgRefMatch = xmlString.match(/<p:bgRef\b[^>]*idx="(\d+)"[^>]*>[\s\S]*?<a:schemeClr\s+val="(\w+)"/i);
            if (bgRefMatch) {
                color = { scheme: bgRefMatch[2] };
            }
        }
        
        // Method 5: Look for bgRef without nested schemeClr but with idx
        if (!color) {
            const bgRefIdxMatch = xmlString.match(/<p:bgRef\b[^>]*idx="([1-9]\d*)"/i);
            if (bgRefIdxMatch) {
                // idx > 0 means a themed background is used
                const idx = parseInt(bgRefIdxMatch[1]);
                // Map common indices to theme colors
                const idxToScheme = { 1: 'dk1', 2: 'lt1', 3: 'dk2', 4: 'lt2', 1001: 'accent1', 1002: 'accent2' };
                color = { scheme: idxToScheme[idx] || 'bg1' };
            }
        }
        
        // Method 6: Look for cSld > bg pattern (common in many PPTX files)
        if (!color) {
            const cSldBgMatch = xmlString.match(/<p:cSld\b[^>]*>[\s\S]*?<p:bg\b[^>]*>[\s\S]*?<a:srgbClr\s+val="([A-Fa-f0-9]{6})"/i);
            if (cSldBgMatch) {
                color = '#' + cSldBgMatch[1];
            }
        }
        
        // Look for gradient fill anywhere in background elements
        const gradientMatch = xmlString.match(/<p:bg\b[^>]*>[\s\S]*?<a:gradFill\b[^>]*>([\s\S]*?)<\/a:gradFill>/i);
        if (gradientMatch) {
            const gradContent = gradientMatch[1];
            const gradColors = [];
            const gsMatches = gradContent.matchAll(/<a:srgbClr\s+val="([A-Fa-f0-9]{6})"/gi);
            for (const m of gsMatches) {
                gradColors.push('#' + m[1]);
            }
            if (gradColors.length >= 2) {
                gradient = gradColors;
                color = null; // Prefer gradient over solid color
            } else if (gradColors.length === 1 && !color) {
                color = gradColors[0];
            }
        }
        
        // Look for background image (blipFill) - multiple patterns
        const blipPatterns = [
            /<p:bg\b[^>]*>[\s\S]*?<a:blip\s[^>]*r:embed="(rId\d+)"/i,
            /<p:bgPr\b[^>]*>[\s\S]*?<a:blip\s[^>]*r:embed="(rId\d+)"/i,
            /<p:bg\b[^>]*>[\s\S]*?r:embed="(rId\d+)"[\s\S]*?<\/p:bg>/i
        ];
        
        for (const pattern of blipPatterns) {
            const blipMatch = xmlString.match(pattern);
            if (blipMatch) {
                imageRId = blipMatch[1];
                break;
            }
        }
        
        return { color, gradient, imageRId };
    },

    // Extract media files from the PPTX
    async extractMediaFiles(zip) {
        const mediaFiles = {};
        
        try {
            zip.forEach((relativePath, zipEntry) => {
                if (relativePath.startsWith('ppt/media/')) {
                    const fileName = relativePath.split('/').pop();
                    mediaFiles[fileName] = zipEntry;
                }
            });
            
            // Convert to base64 data URLs
            for (const [fileName, entry] of Object.entries(mediaFiles)) {
                try {
                    const data = await entry.async('base64');
                    const ext = fileName.split('.').pop().toLowerCase();
                    const mimeType = this.getMimeType(ext);
                    mediaFiles[fileName] = `data:${mimeType};base64,${data}`;
                } catch (err) {
                    console.warn(`Could not extract media file ${fileName}:`, err);
                }
            }
        } catch (err) {
            console.warn('Could not extract media files:', err);
        }
        
        return mediaFiles;
    },

    // Get MIME type for file extension
    getMimeType(ext) {
        const mimeTypes = {
            'png': 'image/png',
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'svg': 'image/svg+xml',
            'webp': 'image/webp',
            'tiff': 'image/tiff',
            'tif': 'image/tiff'
        };
        return mimeTypes[ext] || 'image/png';
    },

    // Parse relationship file
    parseRelationships(relsXml) {
        const rels = {};
        
        // Use regex for reliability
        const relMatches = relsXml.matchAll(/<Relationship\s+[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"[^>]*(?:Type="([^"]*)")?[^>]*\/?>/gi);
        for (const match of relMatches) {
            const id = match[1];
            const target = match[2];
            const type = match[3] || '';
            if (id && target) {
                rels[id] = { target, type };
            }
        }
        
        // Also try reverse order (Target before Id)
        const relMatches2 = relsXml.matchAll(/<Relationship\s+[^>]*Target="([^"]+)"[^>]*Id="(rId\d+)"[^>]*(?:Type="([^"]*)")?[^>]*\/?>/gi);
        for (const match of relMatches2) {
            const target = match[1];
            const id = match[2];
            const type = match[3] || '';
            if (id && target && !rels[id]) {
                rels[id] = { target, type };
            }
        }
        
        return rels;
    },

    // Extract relationships from slide masters and layouts for background images
    async extractMasterRelationships(zip) {
        const allRels = {};
        
        try {
            // Master rels
            const masterRelsFile = zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels');
            if (masterRelsFile) {
                const content = await masterRelsFile.async('string');
                const rels = this.parseRelationships(content);
                for (const [id, info] of Object.entries(rels)) {
                    if (info.target.includes('media/')) {
                        const fileName = info.target.split('/').pop();
                        allRels[id] = { target: info.target, type: info.type };
                    }
                }
            }
            
            // Layout rels
            const layoutRelsFiles = [];
            zip.forEach((path, entry) => {
                if (path.match(/ppt\/slideLayouts\/_rels\/slideLayout\d+\.xml\.rels$/)) {
                    layoutRelsFiles.push(entry);
                }
            });
            
            for (const entry of layoutRelsFiles) {
                try {
                    const content = await entry.async('string');
                    const rels = this.parseRelationships(content);
                    for (const [id, info] of Object.entries(rels)) {
                        if (info.target.includes('media/')) {
                            allRels[id] = { target: info.target, type: info.type };
                        }
                    }
                } catch (err) {
                    // Skip
                }
            }
        } catch (err) {
            console.warn('Could not extract master relationships:', err);
        }
        
        return allRels;
    },

    // Parse individual slide XML
    async parseSlideXML(xmlString, slideNum, themeColors, masterBackground, slideRels, mediaFiles) {
        // Extract background from this slide using regex
        const bgInfo = this.extractBackgroundFromXml(xmlString);
        let background = null;
        let backgroundImage = null;
        
        // Handle background image
        if (bgInfo.imageRId && slideRels[bgInfo.imageRId]) {
            const target = slideRels[bgInfo.imageRId].target;
            const mediaFileName = target.split('/').pop();
            if (mediaFiles[mediaFileName]) {
                backgroundImage = mediaFiles[mediaFileName];
            }
        }
        
        // Handle gradient
        if (bgInfo.gradient && bgInfo.gradient.length >= 2) {
            background = `linear-gradient(180deg, ${bgInfo.gradient.join(', ')})`;
        }
        // Handle solid color
        else if (bgInfo.color) {
            if (typeof bgInfo.color === 'object' && bgInfo.color.scheme) {
                // Resolve scheme color
                const schemeName = bgInfo.color.scheme;
                background = themeColors[schemeName] || themeColors['bg1'] || themeColors['lt1'] || '#ffffff';
            } else {
                background = bgInfo.color;
            }
        }
        
        // Fall back to master background
        if (!background && !backgroundImage && masterBackground) {
            if (masterBackground.gradient && masterBackground.gradient.length >= 2) {
                background = `linear-gradient(180deg, ${masterBackground.gradient.join(', ')})`;
            } else if (masterBackground.color) {
                if (typeof masterBackground.color === 'object' && masterBackground.color.scheme) {
                    const schemeName = masterBackground.color.scheme;
                    background = themeColors[schemeName] || themeColors['bg1'] || '#ffffff';
                } else {
                    background = masterBackground.color;
                }
            }
            if (masterBackground.imageRId && slideRels[masterBackground.imageRId]) {
                const target = slideRels[masterBackground.imageRId].target;
                const mediaFileName = target.split('/').pop();
                if (mediaFiles[mediaFileName]) {
                    backgroundImage = mediaFiles[mediaFileName];
                }
            }
        }
        
        // Final fallback
        if (!background && !backgroundImage) {
            background = '#ffffff';
        }
        
        // Extract text content using regex (more reliable for namespaced XML)
        const textBlocks = [];
        
        // Find all shape text frames
        const spMatches = xmlString.matchAll(/<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/gi);
        for (const spMatch of spMatches) {
            const spContent = spMatch[1];
            const texts = [];
            
            // Find all text runs within this shape
            const tMatches = spContent.matchAll(/<a:t>([^<]*)<\/a:t>/gi);
            for (const tMatch of tMatches) {
                const text = tMatch[1].trim();
                if (text) {
                    texts.push(text);
                }
            }
            
            if (texts.length > 0) {
                textBlocks.push(texts.join(' '));
            }
        }
        
        // Also look for images in the slide (not background)
        const images = [];
        const picMatches = xmlString.matchAll(/<p:pic\b[^>]*>[\s\S]*?<a:blip\s+r:embed="(rId\d+)"[\s\S]*?<\/p:pic>/gi);
        for (const picMatch of picMatches) {
            const embedId = picMatch[1];
            if (embedId && slideRels[embedId]) {
                const target = slideRels[embedId].target;
                const mediaFileName = target.split('/').pop();
                if (mediaFiles[mediaFileName] && !backgroundImage) {
                    // Don't add if it's the same as background
                    if (mediaFiles[mediaFileName] !== backgroundImage) {
                        images.push(mediaFiles[mediaFileName]);
                    }
                }
            }
        }
        
        // Try to identify title and content
        let title = '';
        let subtitle = '';
        let content = '';
        
        if (textBlocks.length > 0) {
            title = textBlocks[0];
            if (textBlocks.length > 1) {
                subtitle = textBlocks[1];
            }
            if (textBlocks.length > 2) {
                content = textBlocks.slice(2).join('\n');
            }
        } else {
            title = `Slide ${slideNum}`;
        }
        
        return {
            id: slideNum,
            title: title,
            subtitle: subtitle,
            content: content,
            layout: textBlocks.length <= 2 ? 'title' : 'content',
            background: background,
            backgroundImage: backgroundImage,
            images: images,
            notes: ''
        };
    },

    // Display presentation UI
    display(file) {
        const state = AppState;
        const canSaveBack = state.slidesFileHandle !== null;
        
        DOM.slidesFileName.textContent = '📽️ ' + file.name + (canSaveBack ? ' ✓' : '');
        DOM.slidesFileSize.textContent = Utils.formatFileSize(file.size);
        
        DOM.slidesInfo.style.display = 'flex';
        DOM.fileInfo.style.display = 'none';
        DOM.docInfo.style.display = 'none';
        DOM.sqlSection.style.display = 'none';
        DOM.tableContainer.style.display = 'none';
        DOM.documentContainer.style.display = 'none';
        DOM.slidesContainer.style.display = 'flex';
        DOM.uploadSection.style.display = 'none';
        DOM.clearBtn.style.display = 'inline-block';
        DOM.sheetTabs.style.display = 'none';
        
        DOM.slidesEditIndicator.style.display = state.slidesHasUnsavedChanges ? 'flex' : 'none';
        
        this.renderThumbnails();
        this.showSlide(state.currentSlideIndex);
    },

    // Render slide thumbnails
    renderThumbnails() {
        const state = AppState;
        
        let html = '';
        state.slidesData.forEach((slide, index) => {
            const isActive = index === state.currentSlideIndex ? 'active' : '';
            
            // Build thumbnail background style
            let thumbBgStyle = '';
            if (slide.backgroundImage) {
                thumbBgStyle = `background-image: url('${slide.backgroundImage}'); background-size: cover; background-position: center;`;
            } else if (slide.background) {
                if (slide.background.startsWith('linear-gradient') || slide.background.startsWith('radial-gradient')) {
                    thumbBgStyle = `background: ${slide.background};`;
                } else {
                    thumbBgStyle = `background-color: ${slide.background};`;
                }
            }
            
            html += `
                <div class="slide-thumbnail ${isActive}" data-index="${index}">
                    <div class="thumbnail-preview" style="${thumbBgStyle}">
                        <span class="thumbnail-number">${index + 1}</span>
                    </div>
                    <div class="thumbnail-content">
                        <div class="thumbnail-title">${Utils.escapeHtml(slide.title || 'Untitled')}</div>
                    </div>
                </div>
            `;
        });
        
        DOM.slidesThumbnails.innerHTML = html;
        
        // Add click handlers
        DOM.slidesThumbnails.querySelectorAll('.slide-thumbnail').forEach(thumb => {
            thumb.addEventListener('click', (e) => {
                const index = parseInt(thumb.dataset.index);
                this.showSlide(index);
            });
        });
    },

    // Show specific slide
    showSlide(index) {
        const state = AppState;
        
        if (index < 0 || index >= state.slidesData.length) return;
        
        state.currentSlideIndex = index;
        const slide = state.slidesData[index];
        
        // Update counter
        DOM.slideCounter.textContent = `${index + 1} / ${state.slidesData.length}`;
        
        // Update navigation buttons
        DOM.prevSlideBtn.disabled = index === 0;
        DOM.nextSlideBtn.disabled = index === state.slidesData.length - 1;
        
        // Render the slide
        this.renderSlide(slide);
        
        // Update thumbnail selection
        DOM.slidesThumbnails.querySelectorAll('.slide-thumbnail').forEach((thumb, i) => {
            thumb.classList.toggle('active', i === index);
        });
        
        // Scroll thumbnail into view
        const activeThumb = DOM.slidesThumbnails.querySelector('.slide-thumbnail.active');
        if (activeThumb) {
            activeThumb.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
    },

    // Render slide content
    renderSlide(slide) {
        const state = AppState;
        const layoutClass = slide.layout === 'title' ? 'slide-layout-title' : 'slide-layout-content';
        
        // Build background style
        let bgStyle = '';
        if (slide.backgroundImage) {
            bgStyle = `background-image: url('${slide.backgroundImage}'); background-size: cover; background-position: center;`;
        } else if (slide.background) {
            // Check if it's a gradient (starts with 'linear-gradient' or 'radial-gradient')
            if (slide.background.startsWith('linear-gradient') || slide.background.startsWith('radial-gradient')) {
                bgStyle = `background: ${slide.background};`;
            } else {
                bgStyle = `background-color: ${slide.background};`;
            }
        } else {
            bgStyle = 'background-color: #ffffff;';
        }
        
        // Determine text color - use custom if set, otherwise auto-detect
        let textColorClass = '';
        let customTextStyle = '';
        if (slide.textColor) {
            customTextStyle = `color: ${slide.textColor};`;
            // Still add a class for editing styles
            textColorClass = this.isLightColor(slide.textColor) ? 'light-text' : 'dark-text';
        } else {
            textColorClass = this.isLightBackground(slide.background, slide.backgroundImage) ? 'dark-text' : 'light-text';
        }
        
        let contentHtml = '';
        if (slide.layout === 'title') {
            contentHtml = `
                <div class="slide-title-main slide-editable" contenteditable="true" data-field="title" placeholder="Click to add title">${Utils.escapeHtml(slide.title)}</div>
                <div class="slide-subtitle slide-editable" contenteditable="true" data-field="subtitle" placeholder="Click to add subtitle">${Utils.escapeHtml(slide.subtitle)}</div>
            `;
        } else {
            contentHtml = `
                <div class="slide-heading slide-editable" contenteditable="true" data-field="title" placeholder="Click to add title">${Utils.escapeHtml(slide.title)}</div>
                <div class="slide-body slide-editable" contenteditable="true" data-field="content" placeholder="Click to add content">${Utils.escapeHtml(slide.content).replace(/\n/g, '<br>')}</div>
            `;
        }
        
        // Add embedded images if any
        let imagesHtml = '';
        if (slide.images && slide.images.length > 0) {
            imagesHtml = '<div class="slide-images">';
            slide.images.forEach((imgSrc, idx) => {
                imagesHtml += `<img src="${imgSrc}" class="slide-embedded-image" alt="Slide image" data-image-idx="${idx}">`;
            });
            imagesHtml += '</div>';
        }
        
        // Background controls
        const bgControlsHtml = `
            <div class="slide-bg-controls">
                <button class="slide-bg-btn" id="changeBgColorBtn" title="Slide style (background & text color)">🎨</button>
                <button class="slide-bg-btn" id="toggleLayoutBtn" title="Toggle layout">${slide.layout === 'title' ? '📑' : '📋'}</button>
            </div>
        `;
        
        DOM.slidesViewer.innerHTML = `
            <div class="slide-canvas ${layoutClass} ${textColorClass}" style="${bgStyle} ${customTextStyle}" data-slide-index="${state.currentSlideIndex}">
                ${bgControlsHtml}
                ${contentHtml}
                ${imagesHtml}
            </div>
        `;
        
        // Attach edit event listeners
        this.attachEditListeners();
    },

    // Attach event listeners for editing
    attachEditListeners() {
        const state = AppState;
        const slideIndex = state.currentSlideIndex;
        
        // Editable fields
        const editableElements = DOM.slidesViewer.querySelectorAll('.slide-editable');
        editableElements.forEach(el => {
            el.addEventListener('input', () => {
                this.handleFieldEdit(el, slideIndex);
            });
            
            el.addEventListener('blur', () => {
                this.saveFieldEdit(el, slideIndex);
            });
            
            el.addEventListener('keydown', (e) => {
                // Prevent arrow key navigation while editing
                if (['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) {
                    e.stopPropagation();
                }
                // Allow Shift+Enter for line breaks in content
                if (e.key === 'Enter' && !e.shiftKey && el.dataset.field !== 'content') {
                    e.preventDefault();
                    el.blur();
                }
            });
        });
        
        // Background color button
        const bgColorBtn = document.getElementById('changeBgColorBtn');
        if (bgColorBtn) {
            bgColorBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                this.showBackgroundPicker(slideIndex);
            });
        }
        
        // Toggle layout button
        const toggleLayoutBtn = document.getElementById('toggleLayoutBtn');
        if (toggleLayoutBtn) {
            toggleLayoutBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                this.toggleSlideLayout(slideIndex);
            });
        }
    },

    // Handle field editing (live update)
    handleFieldEdit(element, slideIndex) {
        // Just mark as editing - actual save happens on blur
        element.classList.add('editing');
    },

    // Save field edit
    saveFieldEdit(element, slideIndex) {
        const state = AppState;
        const field = element.dataset.field;
        let newValue = '';
        
        if (field === 'content') {
            // For content, preserve line breaks
            newValue = element.innerHTML
                .replace(/<br\s*\/?>/gi, '\n')
                .replace(/<div>/gi, '\n')
                .replace(/<\/div>/gi, '')
                .replace(/<[^>]+>/g, '')
                .trim();
        } else {
            newValue = element.textContent.trim();
        }
        
        const slide = state.slidesData[slideIndex];
        if (slide && slide[field] !== newValue) {
            slide[field] = newValue;
            this.markAsEdited();
            
            // Update thumbnail
            this.updateThumbnail(slideIndex);
        }
        
        element.classList.remove('editing');
    },

    // Update a single thumbnail
    updateThumbnail(slideIndex) {
        const state = AppState;
        const slide = state.slidesData[slideIndex];
        const thumbnail = DOM.slidesThumbnails.querySelector(`.slide-thumbnail[data-index="${slideIndex}"]`);
        
        if (thumbnail && slide) {
            const titleEl = thumbnail.querySelector('.thumbnail-title');
            if (titleEl) {
                titleEl.textContent = slide.title || 'Untitled';
            }
        }
    },

    // Show background color picker panel
    showBackgroundPicker(slideIndex) {
        const state = AppState;
        const slide = state.slidesData[slideIndex];
        
        const panel = document.getElementById('colorPickerPanel');
        const colorInput = document.getElementById('slideBgColorInput');
        const colorValue = document.getElementById('colorPickerValue');
        const textColorInput = document.getElementById('slideTextColorInput');
        const textColorValue = document.getElementById('textColorPickerValue');
        const autoTextColorBtn = document.getElementById('autoTextColorBtn');
        const closeBtn = document.getElementById('colorPickerClose');
        const header = document.getElementById('colorPickerHeader');
        const bgImageInput = document.getElementById('bgImageInput');
        const bgImagePreview = document.getElementById('bgImagePreview');
        const bgImageRemove = document.getElementById('bgImageRemove');
        
        if (!panel) return;
        
        // Store current slide index for updates
        this.colorPickerSlideIndex = slideIndex;
        
        // Set current background color
        const currentBgColor = this.getHexColor(slide.background) || '#ffffff';
        colorInput.value = currentBgColor;
        colorValue.textContent = currentBgColor;
        
        // Set current text color
        const isAutoTextColor = !slide.textColor;
        const currentTextColor = slide.textColor || (this.isLightBackground(slide.background, slide.backgroundImage) ? '#1a1a2e' : '#ffffff');
        textColorInput.value = currentTextColor;
        textColorValue.textContent = currentTextColor;
        autoTextColorBtn.classList.toggle('active', isAutoTextColor);
        
        // Set current background image preview
        this.updateBgImagePreview(slide.backgroundImage);
        
        // Show panel
        panel.classList.add('active');
        
        // Setup background color input handler
        const handleBgColorInput = (e) => {
            const color = e.target.value;
            colorValue.textContent = color;
            this.applyBackgroundColor(color);
        };
        
        // Setup text color input handler
        const handleTextColorInput = (e) => {
            const color = e.target.value;
            textColorValue.textContent = color;
            autoTextColorBtn.classList.remove('active');
            this.applyTextColor(color);
        };
        
        // Setup auto text color handler
        const handleAutoTextColor = () => {
            autoTextColorBtn.classList.add('active');
            this.applyAutoTextColor();
            // Update the text color input to show current auto color
            const autoColor = this.isLightBackground(slide.background, slide.backgroundImage) ? '#1a1a2e' : '#ffffff';
            textColorInput.value = autoColor;
            textColorValue.textContent = autoColor + ' (auto)';
        };
        
        // Setup background image upload handler
        const handleImageUpload = (e) => {
            const file = e.target.files[0];
            if (file) {
                this.uploadBackgroundImage(file);
            }
            e.target.value = '';
        };
        
        // Setup background image remove handler
        const handleImageRemove = () => {
            this.removeBackgroundImage();
        };
        
        // Setup close handler
        const handleClose = () => {
            panel.classList.remove('active');
            colorInput.removeEventListener('input', handleBgColorInput);
            textColorInput.removeEventListener('input', handleTextColorInput);
            autoTextColorBtn.removeEventListener('click', handleAutoTextColor);
            bgImageInput.removeEventListener('change', handleImageUpload);
            bgImageRemove.removeEventListener('click', handleImageRemove);
            closeBtn.removeEventListener('click', handleClose);
            this.removeDragListeners();
        };
        
        // Setup background preset color handlers
        const bgPresets = panel.querySelectorAll('.color-preset.bg-preset');
        bgPresets.forEach(preset => {
            preset.onclick = () => {
                const color = preset.dataset.color;
                colorInput.value = color;
                colorValue.textContent = color;
                this.applyBackgroundColor(color);
                bgPresets.forEach(p => p.classList.remove('active'));
                panel.querySelectorAll('.gradient-preset').forEach(g => g.classList.remove('active'));
                preset.classList.add('active');
            };
        });
        
        // Setup text preset color handlers
        const textPresets = panel.querySelectorAll('.color-preset.text-preset');
        textPresets.forEach(preset => {
            preset.onclick = () => {
                const color = preset.dataset.color;
                textColorInput.value = color;
                textColorValue.textContent = color;
                autoTextColorBtn.classList.remove('active');
                this.applyTextColor(color);
                textPresets.forEach(p => p.classList.remove('active'));
                preset.classList.add('active');
            };
        });
        
        // Setup gradient preset handlers
        const gradients = panel.querySelectorAll('.gradient-preset');
        gradients.forEach(gradient => {
            gradient.onclick = () => {
                const gradientValue = gradient.dataset.gradient;
                this.applyBackgroundGradient(gradientValue);
                bgPresets.forEach(p => p.classList.remove('active'));
                gradients.forEach(g => g.classList.remove('active'));
                gradient.classList.add('active');
            };
        });
        
        colorInput.addEventListener('input', handleBgColorInput);
        textColorInput.addEventListener('input', handleTextColorInput);
        autoTextColorBtn.addEventListener('click', handleAutoTextColor);
        bgImageInput.addEventListener('change', handleImageUpload);
        bgImageRemove.addEventListener('click', handleImageRemove);
        closeBtn.addEventListener('click', handleClose);
        
        // Setup drag functionality
        this.setupDragPanel(panel, header);
        
        // Close when clicking outside
        const handleClickOutside = (e) => {
            if (!panel.contains(e.target) && !e.target.closest('.slide-bg-btn')) {
                handleClose();
                document.removeEventListener('mousedown', handleClickOutside);
            }
        };
        setTimeout(() => {
            document.addEventListener('mousedown', handleClickOutside);
        }, 100);
    },
    
    // Update background image preview
    updateBgImagePreview(imageUrl) {
        const preview = document.getElementById('bgImagePreview');
        if (!preview) return;
        
        if (imageUrl) {
            preview.style.backgroundImage = `url('${imageUrl}')`;
            preview.classList.add('has-image');
        } else {
            preview.style.backgroundImage = '';
            preview.classList.remove('has-image');
        }
    },
    
    // Upload and apply background image
    uploadBackgroundImage(file) {
        const state = AppState;
        const slide = state.slidesData[this.colorPickerSlideIndex];
        if (!slide) return;
        
        // Validate file type
        if (!file.type.startsWith('image/')) {
            Utils.showToast('Please select an image file', 'error');
            return;
        }
        
        // Validate file size (max 5MB)
        if (file.size > 5 * 1024 * 1024) {
            Utils.showToast('Image too large. Maximum size is 5MB.', 'error');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = (e) => {
            const imageData = e.target.result;
            slide.backgroundImage = imageData;
            this.markAsEdited();
            this.renderSlide(slide);
            this.renderThumbnails();
            this.updateBgImagePreview(imageData);
            Utils.showToast('Background image applied', 'success');
        };
        reader.onerror = () => {
            Utils.showToast('Error reading image file', 'error');
        };
        reader.readAsDataURL(file);
    },
    
    // Remove background image
    removeBackgroundImage() {
        const state = AppState;
        const slide = state.slidesData[this.colorPickerSlideIndex];
        if (!slide) return;
        
        if (slide.backgroundImage) {
            slide.backgroundImage = null;
            this.markAsEdited();
            this.renderSlide(slide);
            this.renderThumbnails();
            this.updateBgImagePreview(null);
            Utils.showToast('Background image removed', 'info');
        }
    },
    
    // Apply background color to current slide
    applyBackgroundColor(color) {
        const state = AppState;
        const slide = state.slidesData[this.colorPickerSlideIndex];
        if (slide) {
            slide.background = color;
            slide.backgroundImage = null;
            this.markAsEdited();
            this.renderSlide(slide);
            this.renderThumbnails();
        }
    },
    
    // Apply gradient background to current slide
    applyBackgroundGradient(gradient) {
        const state = AppState;
        const slide = state.slidesData[this.colorPickerSlideIndex];
        if (slide) {
            slide.background = gradient;
            slide.backgroundImage = null;
            this.markAsEdited();
            this.renderSlide(slide);
            this.renderThumbnails();
        }
    },
    
    // Apply text color to current slide
    applyTextColor(color) {
        const state = AppState;
        const slide = state.slidesData[this.colorPickerSlideIndex];
        if (slide) {
            slide.textColor = color;
            this.markAsEdited();
            this.renderSlide(slide);
            this.renderThumbnails();
        }
    },
    
    // Apply auto text color (remove custom text color)
    applyAutoTextColor() {
        const state = AppState;
        const slide = state.slidesData[this.colorPickerSlideIndex];
        if (slide) {
            slide.textColor = null; // null means auto-detect
            this.markAsEdited();
            this.renderSlide(slide);
            this.renderThumbnails();
            Utils.showToast('Text color set to auto', 'info');
        }
    },
    
    // Setup drag functionality for panel
    setupDragPanel(panel, header) {
        let isDragging = false;
        let startX, startY, startLeft, startTop;
        
        const getEventCoords = (e) => {
            if (e.touches && e.touches.length > 0) {
                return { x: e.touches[0].clientX, y: e.touches[0].clientY };
            }
            return { x: e.clientX, y: e.clientY };
        };
        
        const handleDragStart = (e) => {
            if (e.target.closest('.color-picker-close')) return;
            
            isDragging = true;
            panel.classList.add('dragging');
            
            const rect = panel.getBoundingClientRect();
            const coords = getEventCoords(e);
            startX = coords.x;
            startY = coords.y;
            startLeft = rect.left;
            startTop = rect.top;
            
            // Switch to fixed positioning with current position
            panel.style.left = startLeft + 'px';
            panel.style.top = startTop + 'px';
            panel.style.right = 'auto';
            
            document.addEventListener('mousemove', handleDragMove);
            document.addEventListener('mouseup', handleDragEnd);
            document.addEventListener('touchmove', handleDragMove, { passive: false });
            document.addEventListener('touchend', handleDragEnd);
            e.preventDefault();
        };
        
        const handleDragMove = (e) => {
            if (!isDragging) return;
            
            const coords = getEventCoords(e);
            const dx = coords.x - startX;
            const dy = coords.y - startY;
            
            let newLeft = startLeft + dx;
            let newTop = startTop + dy;
            
            // Keep panel within viewport
            const panelRect = panel.getBoundingClientRect();
            newLeft = Math.max(0, Math.min(newLeft, window.innerWidth - panelRect.width));
            newTop = Math.max(0, Math.min(newTop, window.innerHeight - panelRect.height));
            
            panel.style.left = newLeft + 'px';
            panel.style.top = newTop + 'px';
            
            if (e.cancelable) e.preventDefault();
        };
        
        const handleDragEnd = () => {
            isDragging = false;
            panel.classList.remove('dragging');
            document.removeEventListener('mousemove', handleDragMove);
            document.removeEventListener('mouseup', handleDragEnd);
            document.removeEventListener('touchmove', handleDragMove);
            document.removeEventListener('touchend', handleDragEnd);
        };
        
        header.addEventListener('mousedown', handleDragStart);
        header.addEventListener('touchstart', handleDragStart, { passive: false });
        
        // Store cleanup function
        this.removeDragListeners = () => {
            header.removeEventListener('mousedown', handleDragStart);
            header.removeEventListener('touchstart', handleDragStart);
            document.removeEventListener('mousemove', handleDragMove);
            document.removeEventListener('mouseup', handleDragEnd);
            document.removeEventListener('touchmove', handleDragMove);
            document.removeEventListener('touchend', handleDragEnd);
        };
    },

    // Get hex color from background value
    getHexColor(bg) {
        if (!bg) return '#ffffff';
        if (bg.startsWith('#')) return bg;
        if (bg.startsWith('linear-gradient') || bg.startsWith('radial-gradient')) {
            // Extract first color from gradient
            const match = bg.match(/#[A-Fa-f0-9]{6}/);
            return match ? match[0] : '#ffffff';
        }
        return '#ffffff';
    },

    // Toggle slide layout between title and content
    toggleSlideLayout(slideIndex) {
        const state = AppState;
        const slide = state.slidesData[slideIndex];
        
        if (slide.layout === 'title') {
            slide.layout = 'content';
            // Move subtitle to content if switching to content layout
            if (slide.subtitle && !slide.content) {
                slide.content = slide.subtitle;
            }
        } else {
            slide.layout = 'title';
            // Move content to subtitle if it's short enough
            if (slide.content && !slide.subtitle && slide.content.length < 100) {
                slide.subtitle = slide.content;
                slide.content = '';
            }
        }
        
        this.markAsEdited();
        this.renderSlide(slide);
        Utils.showToast(`Changed to ${slide.layout} layout`, 'info');
    },

    // Check if background is light (to determine text color)
    isLightBackground(bgColor, bgImage) {
        // If there's a background image, assume dark text is safer (most images work with dark text)
        if (bgImage) {
            return true;
        }
        
        if (!bgColor || bgColor.startsWith('linear-gradient') || bgColor.startsWith('radial-gradient')) {
            return true; // Default to dark text for gradients
        }
        
        return this.isLightColor(bgColor);
    },
    
    // Check if a color is light
    isLightColor(color) {
        if (!color) return true;
        
        // Parse hex color
        let hex = color.replace('#', '');
        if (hex.length === 3) {
            hex = hex.split('').map(c => c + c).join('');
        }
        
        if (!/^[0-9A-Fa-f]{6}$/.test(hex)) {
            return true; // Default to light
        }
        
        const r = parseInt(hex.substr(0, 2), 16);
        const g = parseInt(hex.substr(2, 2), 16);
        const b = parseInt(hex.substr(4, 2), 16);
        
        // Calculate luminance
        const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
        
        return luminance > 0.5;
    },
    
    // Get text color for a slide (custom or auto-detected)
    getSlideTextColor(slide) {
        if (slide.textColor) {
            return slide.textColor;
        }
        return this.isLightBackground(slide.background, slide.backgroundImage) ? '#1a1a2e' : '#ffffff';
    },

    // Navigate to previous slide
    prevSlide() {
        const state = AppState;
        if (state.currentSlideIndex > 0) {
            this.showSlide(state.currentSlideIndex - 1);
        }
    },

    // Navigate to next slide
    nextSlide() {
        const state = AppState;
        if (state.currentSlideIndex < state.slidesData.length - 1) {
            this.showSlide(state.currentSlideIndex + 1);
        }
    },

    // Add new slide
    addSlide() {
        const state = AppState;
        
        // Get background from current slide or use default
        const currentSlide = state.slidesData[state.currentSlideIndex];
        const bgColor = currentSlide ? currentSlide.background : '#ffffff';
        
        const newSlide = {
            id: state.slidesData.length + 1,
            title: 'New Slide',
            subtitle: '',
            content: '',
            layout: 'content',
            background: bgColor,
            backgroundImage: null,
            images: [],
            notes: ''
        };
        
        // Insert after current slide
        const insertIndex = state.currentSlideIndex + 1;
        state.slidesData.splice(insertIndex, 0, newSlide);
        
        // Renumber slides
        this.renumberSlides();
        
        this.markAsEdited();
        this.renderThumbnails();
        this.showSlide(insertIndex);
        
        Utils.showToast('New slide added', 'success');
    },

    // Duplicate current slide
    duplicateSlide() {
        const state = AppState;
        const currentSlide = state.slidesData[state.currentSlideIndex];
        
        if (!currentSlide) return;
        
        const duplicatedSlide = {
            ...currentSlide,
            id: state.slidesData.length + 1,
            title: currentSlide.title,
            subtitle: currentSlide.subtitle,
            content: currentSlide.content,
            images: currentSlide.images ? [...currentSlide.images] : []
        };
        
        // Insert after current slide
        const insertIndex = state.currentSlideIndex + 1;
        state.slidesData.splice(insertIndex, 0, duplicatedSlide);
        
        // Renumber slides
        this.renumberSlides();
        
        this.markAsEdited();
        this.renderThumbnails();
        this.showSlide(insertIndex);
        
        Utils.showToast('Slide duplicated', 'success');
    },

    // Renumber slides
    renumberSlides() {
        const state = AppState;
        state.slidesData.forEach((slide, i) => {
            slide.id = i + 1;
        });
    },

    // Delete current slide
    async deleteSlide() {
        const state = AppState;
        
        if (state.slidesData.length <= 1) {
            Utils.showToast('Cannot delete the only slide', 'warning');
            return;
        }
        
        const confirmed = await Utils.confirm('Are you sure you want to delete this slide?', {
            title: 'Delete Slide',
            icon: '🗑️',
            okText: 'Delete',
            cancelText: 'Cancel',
            danger: true
        });
        if (!confirmed) return;
        
        state.slidesData.splice(state.currentSlideIndex, 1);
        
        // Renumber slides
        this.renumberSlides();
        
        // Adjust current index
        if (state.currentSlideIndex >= state.slidesData.length) {
            state.currentSlideIndex = state.slidesData.length - 1;
        }
        
        this.markAsEdited();
        this.renderThumbnails();
        this.showSlide(state.currentSlideIndex);
        
        Utils.showToast('Slide deleted', 'info');
    },

    // Move slide up (earlier in presentation)
    moveSlideUp() {
        const state = AppState;
        if (state.currentSlideIndex <= 0) return;
        
        const slide = state.slidesData.splice(state.currentSlideIndex, 1)[0];
        state.slidesData.splice(state.currentSlideIndex - 1, 0, slide);
        
        this.renumberSlides();
        state.currentSlideIndex--;
        
        this.markAsEdited();
        this.renderThumbnails();
        this.showSlide(state.currentSlideIndex);
    },

    // Move slide down (later in presentation)
    moveSlideDown() {
        const state = AppState;
        if (state.currentSlideIndex >= state.slidesData.length - 1) return;
        
        const slide = state.slidesData.splice(state.currentSlideIndex, 1)[0];
        state.slidesData.splice(state.currentSlideIndex + 1, 0, slide);
        
        this.renumberSlides();
        state.currentSlideIndex++;
        
        this.markAsEdited();
        this.renderThumbnails();
        this.showSlide(state.currentSlideIndex);
    },

    // Mark as edited
    markAsEdited() {
        const state = AppState;
        if (!state.slidesHasUnsavedChanges) {
            state.slidesHasUnsavedChanges = true;
            DOM.slidesEditIndicator.style.display = 'flex';
        }
    },

    // Save presentation
    async save() {
        const state = AppState;
        
        if (state.currentMode !== 'slides') {
            return;
        }
        
        try {
            if ('showSaveFilePicker' in window) {
                await this.saveWithPicker();
            } else {
                await this.exportAsPptx();
            }
            
            state.slidesHasUnsavedChanges = false;
            DOM.slidesEditIndicator.style.display = 'none';
            Utils.showToast('Presentation saved!', 'success');
            
        } catch (error) {
            if (error.name === 'AbortError') {
                return;
            }
            console.error('Save error:', error);
            Utils.showToast('Failed to save presentation: ' + error.message, 'error');
        }
    },

    // Save with file picker
    async saveWithPicker() {
        const state = AppState;
        const fileName = state.slidesFile?.name?.replace(/\.[^/.]+$/, '') || 'presentation';
        
        const handle = await window.showSaveFilePicker({
            suggestedName: `${fileName}.pptx`,
            types: [{
                description: 'PowerPoint Presentation',
                accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] }
            }]
        });
        
        state.slidesFileHandle = handle;
        state.slidesFile = { name: handle.name, size: 0 };
        
        const blob = await this.createPptxBlob();
        
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        
        DOM.slidesFileName.textContent = '📽️ ' + handle.name + ' ✓';
        DOM.slidesFileSize.textContent = 'Saved';
    },

    // Export presentation - show modal with format options
    export() {
        const state = AppState;
        
        if (state.slidesData.length === 0) {
            Utils.showToast('No presentation to export', 'error');
            return;
        }
        
        // Show export modal
        const modal = document.getElementById('slidesExportModal');
        const closeBtn = document.getElementById('slidesExportModalClose');
        const pptxBtn = document.getElementById('exportPptxBtn');
        const pdfBtn = document.getElementById('exportPdfBtn');
        
        if (!modal) {
            // Fallback to PPTX if modal not found
            this.exportAsPptx();
            return;
        }
        
        modal.classList.add('show');
        
        const closeModal = () => {
            modal.classList.remove('show');
        };
        
        // Setup handlers
        const handlePptx = () => {
            closeModal();
            this.exportAsPptx();
        };
        
        const handlePdf = () => {
            closeModal();
            this.exportAsPdf();
        };
        
        const handleClose = () => {
            closeModal();
        };
        
        const handleOverlayClick = (e) => {
            if (e.target === modal) {
                closeModal();
            }
        };
        
        // Attach one-time listeners
        pptxBtn.onclick = handlePptx;
        pdfBtn.onclick = handlePdf;
        closeBtn.onclick = handleClose;
        modal.onclick = handleOverlayClick;
    },

    // Export as PPTX
    async exportAsPptx() {
        const state = AppState;
        const fileName = state.slidesFile?.name?.replace(/\.[^/.]+$/, '') || 'presentation';
        
        Utils.showToast('Generating PowerPoint...', 'info');
        const blob = await this.createPptxBlob();
        Utils.downloadBlobFile(blob, `${fileName}.pptx`);
        Utils.showToast('PowerPoint exported successfully!', 'success');
    },
    
    // Export as PDF using html2canvas + jsPDF
    async exportAsPdf() {
        const state = AppState;
        const fileName = state.slidesFile?.name?.replace(/\.[^/.]+$/, '') || 'presentation';
        
        // Check if libraries are available
        if (typeof html2canvas === 'undefined' || typeof jspdf === 'undefined') {
            Utils.showToast('PDF libraries not loaded. Please refresh the page.', 'error');
            return;
        }
        
        Utils.showToast('Generating PDF... This may take a moment.', 'info');
        
        try {
            // Create a hidden container for rendering slides
            const renderContainer = document.createElement('div');
            renderContainer.id = 'pdfRenderContainer';
            renderContainer.style.cssText = `
                position: fixed;
                left: -9999px;
                top: 0;
                width: 1280px;
                height: 720px;
                z-index: -1;
                overflow: hidden;
            `;
            document.body.appendChild(renderContainer);
            
            // Create PDF (landscape, 16:9 aspect ratio)
            const { jsPDF } = jspdf;
            const pdf = new jsPDF({
                orientation: 'landscape',
                unit: 'px',
                format: [1280, 720],
                hotfixes: ['px_scaling']
            });
            
            // Process each slide
            for (let i = 0; i < state.slidesData.length; i++) {
                const slide = state.slidesData[i];
                
                // Update progress
                Utils.showToast(`Rendering slide ${i + 1} of ${state.slidesData.length}...`, 'info');
                
                // Render slide to container
                renderContainer.innerHTML = this.createPdfSlideHtml(slide, i);
                
                // Wait for images to load
                await this.waitForImages(renderContainer);
                
                // Capture as canvas
                const canvas = await html2canvas(renderContainer, {
                    scale: 2,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: null,
                    logging: false,
                    width: 1280,
                    height: 720
                });
                
                // Add page (except for first slide)
                if (i > 0) {
                    pdf.addPage([1280, 720], 'landscape');
                }
                
                // Add canvas image to PDF
                const imgData = canvas.toDataURL('image/jpeg', 0.95);
                pdf.addImage(imgData, 'JPEG', 0, 0, 1280, 720);
            }
            
            // Clean up
            document.body.removeChild(renderContainer);
            
            // Save PDF
            pdf.save(`${fileName}.pdf`);
            
            Utils.showToast('PDF exported successfully!', 'success');
            
        } catch (error) {
            console.error('PDF export error:', error);
            Utils.showToast('Error generating PDF: ' + error.message, 'error');
        }
    },
    
    // Create HTML for a single slide for PDF rendering
    createPdfSlideHtml(slide, index) {
        const bgStyle = this.getBackgroundStyle(slide);
        const textColor = this.getSlideTextColor(slide);
        const textColorClass = this.isLightColor(textColor) ? 'light-text' : 'dark-text';
        
        let contentHtml = '';
        if (slide.layout === 'title') {
            contentHtml = `
                <div class="pdf-title">${Utils.escapeHtml(slide.title || '')}</div>
                <div class="pdf-subtitle">${Utils.escapeHtml(slide.subtitle || '')}</div>
            `;
        } else {
            contentHtml = `
                <div class="pdf-heading">${Utils.escapeHtml(slide.title || '')}</div>
                <div class="pdf-body">${Utils.escapeHtml(slide.content || '').replace(/\n/g, '<br>')}</div>
            `;
        }
        
        // Add images if present
        let imagesHtml = '';
        if (slide.images && slide.images.length > 0) {
            imagesHtml = '<div class="pdf-images">';
            slide.images.forEach(imgSrc => {
                imagesHtml += `<img src="${imgSrc}" class="pdf-image" crossorigin="anonymous">`;
            });
            imagesHtml += '</div>';
        }
        
        return `
            <style>
                .pdf-slide-container {
                    width: 1280px;
                    height: 720px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    position: relative;
                    overflow: hidden;
                    ${bgStyle}
                }
                
                .pdf-slide-container.dark-text { color: #1a1a2e; }
                .pdf-slide-container.light-text { color: #ffffff; }
                
                .pdf-content {
                    width: 90%;
                    max-width: 1100px;
                    text-align: center;
                    padding: 40px;
                }
                
                .pdf-content.content-layout {
                    text-align: left;
                }
                
                .pdf-title {
                    font-size: 64px;
                    font-weight: 700;
                    margin-bottom: 24px;
                    line-height: 1.2;
                }
                
                .pdf-subtitle {
                    font-size: 32px;
                    opacity: 0.85;
                    line-height: 1.4;
                }
                
                .pdf-heading {
                    font-size: 48px;
                    font-weight: 600;
                    margin-bottom: 30px;
                    padding-bottom: 15px;
                    border-bottom: 4px solid currentColor;
                    opacity: 0.95;
                }
                
                .pdf-body {
                    font-size: 28px;
                    line-height: 1.6;
                }
                
                .pdf-images {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 20px;
                    justify-content: center;
                    margin-top: 30px;
                }
                
                .pdf-image {
                    max-width: 45%;
                    max-height: 280px;
                    object-fit: contain;
                    border-radius: 8px;
                }
                
                .pdf-slide-number {
                    position: absolute;
                    bottom: 20px;
                    right: 30px;
                    font-size: 16px;
                    opacity: 0.5;
                }
            </style>
            <div class="pdf-slide-container ${textColorClass}" style="color: ${textColor};">
                <div class="pdf-content ${slide.layout === 'title' ? 'title-layout' : 'content-layout'}">
                    ${contentHtml}
                    ${imagesHtml}
                </div>
                <div class="pdf-slide-number">${index + 1}</div>
            </div>
        `;
    },
    
    // Wait for all images in container to load
    waitForImages(container) {
        const images = container.querySelectorAll('img');
        const promises = Array.from(images).map(img => {
            if (img.complete) return Promise.resolve();
            return new Promise((resolve) => {
                img.onload = resolve;
                img.onerror = resolve; // Continue even if image fails
            });
        });
        return Promise.all(promises);
    },
    
    // Get background style string for a slide
    getBackgroundStyle(slide) {
        if (slide.backgroundImage) {
            return `background-image: url('${slide.backgroundImage}'); background-size: cover; background-position: center;`;
        } else if (slide.background) {
            if (slide.background.startsWith('linear-gradient') || slide.background.startsWith('radial-gradient')) {
                return `background: ${slide.background};`;
            } else {
                return `background-color: ${slide.background};`;
            }
        }
        return 'background-color: #ffffff;';
    },

    // Create PPTX blob using PptxGenJS
    async createPptxBlob() {
        const state = AppState;
        
        // Create new presentation
        const pptx = new PptxGenJS();
        pptx.author = 'SuiteTools';
        pptx.title = state.slidesFile?.name?.replace(/\.[^/.]+$/, '') || 'Presentation';
        pptx.subject = 'Created with SuiteTools';
        
        // Add each slide
        state.slidesData.forEach((slideData, index) => {
            const slide = pptx.addSlide();
            
            // Set background
            if (slideData.background && slideData.background !== '#ffffff') {
                slide.background = { color: slideData.background.replace('#', '') };
            }
            
            if (slideData.layout === 'title') {
                // Title slide layout
                slide.addText(slideData.title || '', {
                    x: 0.5,
                    y: 2.5,
                    w: 9,
                    h: 1.5,
                    fontSize: 44,
                    bold: true,
                    color: '363636',
                    align: 'center',
                    valign: 'middle'
                });
                
                slide.addText(slideData.subtitle || '', {
                    x: 0.5,
                    y: 4,
                    w: 9,
                    h: 0.8,
                    fontSize: 24,
                    color: '666666',
                    align: 'center',
                    valign: 'middle'
                });
            } else {
                // Content slide layout
                slide.addText(slideData.title || '', {
                    x: 0.5,
                    y: 0.3,
                    w: 9,
                    h: 0.8,
                    fontSize: 32,
                    bold: true,
                    color: '363636'
                });
                
                slide.addText(slideData.content || '', {
                    x: 0.5,
                    y: 1.3,
                    w: 9,
                    h: 4.5,
                    fontSize: 18,
                    color: '444444',
                    valign: 'top'
                });
            }
        });
        
        // Generate blob
        const blob = await pptx.write({ outputType: 'blob' });
        return blob;
    },

    // Clear slides
    clear() {
        const state = AppState;
        state.slidesData = [];
        state.currentSlideIndex = 0;
        state.slidesFile = null;
        state.slidesFileHandle = null;
        state.slidesHasUnsavedChanges = false;
        
        DOM.slidesContainer.style.display = 'none';
        DOM.slidesInfo.style.display = 'none';
        DOM.slidesEditIndicator.style.display = 'none';
    },

    // ==========================================
    // SLIDESHOW MODE
    // ==========================================

    slideshowAutoplayTimer: null,
    slideshowAutoplayInterval: 5000, // 5 seconds per slide

    // Start slideshow mode
    startSlideshow() {
        const state = AppState;
        if (state.slidesData.length === 0) {
            Utils.showToast('No slides to present', 'warning');
            return;
        }

        const overlay = document.getElementById('slideshowOverlay');
        overlay.classList.add('active');
        
        // Request fullscreen
        this.requestFullscreen(overlay);
        
        // Show first slide or current slide
        this.slideshowCurrentIndex = state.currentSlideIndex;
        this.renderSlideshowSlide();
        
        // Setup event listeners
        this.setupSlideshowListeners();
        
        // Show controls initially
        this.showSlideshowControls();
        
        Utils.showToast('Slideshow started. Press Esc to exit.', 'info');
    },

    // Exit slideshow mode
    exitSlideshow() {
        const overlay = document.getElementById('slideshowOverlay');
        overlay.classList.remove('active');
        overlay.classList.remove('show-controls');
        
        // Stop autoplay
        this.stopAutoplay();
        
        // Exit fullscreen
        this.exitFullscreen();
        
        // Remove event listeners
        this.removeSlideshowListeners();
        
        // Sync back to editor
        AppState.currentSlideIndex = this.slideshowCurrentIndex;
        this.showSlide(this.slideshowCurrentIndex);
    },

    // Request fullscreen
    requestFullscreen(element) {
        if (element.requestFullscreen) {
            element.requestFullscreen().catch(err => {
                console.log('Fullscreen not available:', err);
            });
        } else if (element.webkitRequestFullscreen) {
            element.webkitRequestFullscreen();
        } else if (element.msRequestFullscreen) {
            element.msRequestFullscreen();
        }
    },

    // Exit fullscreen
    exitFullscreen() {
        if (document.exitFullscreen) {
            document.exitFullscreen().catch(() => {});
        } else if (document.webkitExitFullscreen) {
            document.webkitExitFullscreen();
        } else if (document.msExitFullscreen) {
            document.msExitFullscreen();
        }
    },

    // Render current slideshow slide
    renderSlideshowSlide() {
        const state = AppState;
        const slide = state.slidesData[this.slideshowCurrentIndex];
        if (!slide) return;

        const container = document.getElementById('slideshowContainer');
        const layoutClass = slide.layout === 'title' ? 'slide-layout-title' : 'slide-layout-content';
        
        // Build background style
        let bgStyle = '';
        if (slide.backgroundImage) {
            bgStyle = `background-image: url('${slide.backgroundImage}'); background-size: cover; background-position: center;`;
        } else if (slide.background) {
            if (slide.background.startsWith('linear-gradient') || slide.background.startsWith('radial-gradient')) {
                bgStyle = `background: ${slide.background};`;
            } else {
                bgStyle = `background-color: ${slide.background};`;
            }
        } else {
            bgStyle = 'background-color: #ffffff;';
        }
        
        // Get text color
        const textColor = this.getSlideTextColor(slide);
        const textColorClass = this.isLightColor(textColor) ? 'light-text' : 'dark-text';
        const customTextStyle = `color: ${textColor};`;
        
        let contentHtml = '';
        if (slide.layout === 'title') {
            contentHtml = `
                <div class="slide-title-main">${Utils.escapeHtml(slide.title)}</div>
                <div class="slide-subtitle">${Utils.escapeHtml(slide.subtitle)}</div>
            `;
        } else {
            contentHtml = `
                <div class="slide-heading">${Utils.escapeHtml(slide.title)}</div>
                <div class="slide-body">${Utils.escapeHtml(slide.content).replace(/\n/g, '<br>')}</div>
            `;
        }
        
        // Add embedded images
        let imagesHtml = '';
        if (slide.images && slide.images.length > 0) {
            imagesHtml = '<div class="slide-images">';
            slide.images.forEach(imgSrc => {
                imagesHtml += `<img src="${imgSrc}" class="slide-embedded-image" alt="">`;
            });
            imagesHtml += '</div>';
        }
        
        container.innerHTML = `
            <div class="slide-canvas ${layoutClass} ${textColorClass}" style="${bgStyle} ${customTextStyle}">
                ${contentHtml}
                ${imagesHtml}
            </div>
            <div class="slideshow-nav-area prev" id="slideshowNavPrev">
                <div class="slideshow-nav-icon">◀</div>
            </div>
            <div class="slideshow-nav-area next" id="slideshowNavNext">
                <div class="slideshow-nav-icon">▶</div>
            </div>
        `;
        
        // Update counter
        document.getElementById('slideshowCounter').textContent = 
            `${this.slideshowCurrentIndex + 1} / ${state.slidesData.length}`;
        
        // Update progress bar
        const progress = ((this.slideshowCurrentIndex + 1) / state.slidesData.length) * 100;
        document.getElementById('slideshowProgressBar').style.width = `${progress}%`;
        
        // Update navigation button states
        document.getElementById('slideshowPrev').disabled = this.slideshowCurrentIndex === 0;
        document.getElementById('slideshowNext').disabled = this.slideshowCurrentIndex === state.slidesData.length - 1;
        
        // Add click handlers for nav areas
        document.getElementById('slideshowNavPrev').onclick = () => this.slideshowPrev();
        document.getElementById('slideshowNavNext').onclick = () => this.slideshowNext();
    },

    // Go to next slide in slideshow
    slideshowNext() {
        const state = AppState;
        if (this.slideshowCurrentIndex < state.slidesData.length - 1) {
            this.slideshowCurrentIndex++;
            this.renderSlideshowSlide();
            this.resetAutoplayTimer();
        } else if (this.isAutoplayActive()) {
            // Loop back to first slide in autoplay mode
            this.slideshowCurrentIndex = 0;
            this.renderSlideshowSlide();
        }
    },

    // Go to previous slide in slideshow
    slideshowPrev() {
        if (this.slideshowCurrentIndex > 0) {
            this.slideshowCurrentIndex--;
            this.renderSlideshowSlide();
            this.resetAutoplayTimer();
        }
    },

    // Go to specific slide in slideshow
    slideshowGoTo(index) {
        const state = AppState;
        if (index >= 0 && index < state.slidesData.length) {
            this.slideshowCurrentIndex = index;
            this.renderSlideshowSlide();
            this.resetAutoplayTimer();
        }
    },

    // Toggle autoplay
    toggleAutoplay() {
        if (this.isAutoplayActive()) {
            this.stopAutoplay();
        } else {
            this.startAutoplay();
        }
    },

    // Start autoplay
    startAutoplay() {
        const btn = document.getElementById('slideshowAutoplay');
        btn.classList.add('active', 'autoplay-active');
        btn.textContent = '⏸';
        btn.title = 'Pause auto-play';
        
        this.slideshowAutoplayTimer = setInterval(() => {
            this.slideshowNext();
        }, this.slideshowAutoplayInterval);
    },

    // Stop autoplay
    stopAutoplay() {
        if (this.slideshowAutoplayTimer) {
            clearInterval(this.slideshowAutoplayTimer);
            this.slideshowAutoplayTimer = null;
        }
        
        const btn = document.getElementById('slideshowAutoplay');
        if (btn) {
            btn.classList.remove('active', 'autoplay-active');
            btn.textContent = '⏵';
            btn.title = 'Auto-play';
        }
    },

    // Reset autoplay timer (when manually navigating)
    resetAutoplayTimer() {
        if (this.isAutoplayActive()) {
            this.stopAutoplay();
            this.startAutoplay();
        }
    },

    // Check if autoplay is active
    isAutoplayActive() {
        return this.slideshowAutoplayTimer !== null;
    },

    // Show controls (on mouse move)
    showSlideshowControls() {
        const overlay = document.getElementById('slideshowOverlay');
        overlay.classList.add('show-controls');
        
        // Hide after 3 seconds of no movement
        if (this.controlsHideTimer) {
            clearTimeout(this.controlsHideTimer);
        }
        this.controlsHideTimer = setTimeout(() => {
            overlay.classList.remove('show-controls');
        }, 3000);
    },

    // Setup slideshow event listeners
    setupSlideshowListeners() {
        // Keyboard navigation
        this.slideshowKeyHandler = (e) => {
            switch (e.key) {
                case 'Escape':
                    this.exitSlideshow();
                    break;
                case 'ArrowRight':
                case 'ArrowDown':
                case ' ':
                case 'PageDown':
                    e.preventDefault();
                    this.slideshowNext();
                    break;
                case 'ArrowLeft':
                case 'ArrowUp':
                case 'PageUp':
                    e.preventDefault();
                    this.slideshowPrev();
                    break;
                case 'Home':
                    e.preventDefault();
                    this.slideshowGoTo(0);
                    break;
                case 'End':
                    e.preventDefault();
                    this.slideshowGoTo(AppState.slidesData.length - 1);
                    break;
                case 'p':
                case 'P':
                    this.toggleAutoplay();
                    break;
            }
            this.showSlideshowControls();
        };
        document.addEventListener('keydown', this.slideshowKeyHandler);
        
        // Mouse movement to show controls
        this.slideshowMouseHandler = () => {
            this.showSlideshowControls();
        };
        document.getElementById('slideshowOverlay').addEventListener('mousemove', this.slideshowMouseHandler);
        
        // Control button handlers
        document.getElementById('slideshowPrev').onclick = () => this.slideshowPrev();
        document.getElementById('slideshowNext').onclick = () => this.slideshowNext();
        document.getElementById('slideshowExit').onclick = () => this.exitSlideshow();
        document.getElementById('slideshowAutoplay').onclick = () => this.toggleAutoplay();
        
        // Fullscreen change handler
        this.fullscreenChangeHandler = () => {
            if (!document.fullscreenElement && !document.webkitFullscreenElement) {
                // Exited fullscreen, exit slideshow
                const overlay = document.getElementById('slideshowOverlay');
                if (overlay.classList.contains('active')) {
                    this.exitSlideshow();
                }
            }
        };
        document.addEventListener('fullscreenchange', this.fullscreenChangeHandler);
        document.addEventListener('webkitfullscreenchange', this.fullscreenChangeHandler);
    },

    // Remove slideshow event listeners
    removeSlideshowListeners() {
        if (this.slideshowKeyHandler) {
            document.removeEventListener('keydown', this.slideshowKeyHandler);
        }
        if (this.slideshowMouseHandler) {
            document.getElementById('slideshowOverlay')?.removeEventListener('mousemove', this.slideshowMouseHandler);
        }
        if (this.fullscreenChangeHandler) {
            document.removeEventListener('fullscreenchange', this.fullscreenChangeHandler);
            document.removeEventListener('webkitfullscreenchange', this.fullscreenChangeHandler);
        }
        if (this.controlsHideTimer) {
            clearTimeout(this.controlsHideTimer);
        }
    }
};

// Export for use in other modules
window.SlidesEditor = SlidesEditor;