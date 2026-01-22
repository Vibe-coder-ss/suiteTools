/**
 * Document Editor Module
 * Handles document creation, editing, and export
 */

const DocumentEditor = {
    // Create empty document
    async createEmpty() {
        const state = AppState;
        
        if (state.hasUnsavedChanges || state.docHasUnsavedChanges || state.slidesHasUnsavedChanges) {
            const confirmed = await Utils.confirm('You have unsaved changes. Creating a new document will clear them. Continue?', {
                title: 'Unsaved Changes',
                icon: '⚠️',
                okText: 'Continue',
                cancelText: 'Cancel',
                danger: true
            });
            if (!confirmed) return;
        }
        
        Spreadsheet.clearSilent();
        
        state.currentMode = 'document';
        state.currentFile = { name: 'new_document.docx', size: 0 };
        state.originalDocContent = '';
        state.docHasUnsavedChanges = true;
        state.docFileHandle = null;
        
        DOM.docEditor.innerHTML = '<p></p>';
        
        DOM.docFileName.textContent = '📄 new_document.docx (New)';
        DOM.docFileSize.textContent = 'New Document';
        
        DOM.docInfo.style.display = 'flex';
        DOM.fileInfo.style.display = 'none';
        DOM.sqlSection.style.display = 'none';
        DOM.tableContainer.style.display = 'none';
        DOM.documentContainer.style.display = 'flex';
        DOM.uploadSection.style.display = 'none';
        DOM.clearBtn.style.display = 'inline-block';
        DOM.sheetTabs.style.display = 'none';
        
        DOM.docEditIndicator.style.display = 'flex';
        
        DOM.docEditor.focus();
        this.updateStats();
        
        Utils.showToast('New document created! Start typing to add content.', 'success');
    },

    // Parse document file (DOCX)
    async parse(arrayBuffer, file) {
        try {
            const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
            const html = result.value;
            
            const state = AppState;
            state.currentFile = file;
            state.currentMode = 'document';
            state.originalDocContent = html;
            state.docHasUnsavedChanges = false;
            
            DOM.docEditor.innerHTML = html || '<p>Start typing your document here...</p>';
            
            this.display(file);
            this.updateStats();
            
            if (result.messages.length > 0) {
                console.log('Document conversion messages:', result.messages);
            }
            
        } catch (error) {
            console.error('Error parsing document:', error);
            Utils.alert('Error reading document file. Please try again.', {
                title: 'Import Error',
                icon: '⚠️'
            });
            Spreadsheet.clear();
        }
    },

    // Display document UI
    display(file) {
        const state = AppState;
        const canSaveBack = state.docFileHandle !== null;
        
        DOM.docFileName.textContent = '📄 ' + file.name + (canSaveBack ? ' ✓' : '');
        DOM.docFileSize.textContent = Utils.formatFileSize(file.size);
        
        DOM.docInfo.style.display = 'flex';
        DOM.fileInfo.style.display = 'none';
        DOM.sqlSection.style.display = 'none';
        DOM.tableContainer.style.display = 'none';
        DOM.documentContainer.style.display = 'flex';
        DOM.uploadSection.style.display = 'none';
        DOM.clearBtn.style.display = 'inline-block';
        DOM.sheetTabs.style.display = 'none';
        
        DOM.docEditIndicator.style.display = 'none';
        
        DOM.docEditor.focus();
    },

    // Mark document as edited
    markAsEdited() {
        const state = AppState;
        if (!state.docHasUnsavedChanges) {
            state.docHasUnsavedChanges = true;
            DOM.docEditIndicator.style.display = 'flex';
        }
    },

    // Update document statistics
    updateStats() {
        const text = DOM.docEditor.innerText || '';
        const words = text.trim() ? text.trim().split(/\s+/).length : 0;
        const chars = text.length;
        
        DOM.docWordCount.textContent = `${words} word${words !== 1 ? 's' : ''}`;
        DOM.docCharCount.textContent = `${chars} character${chars !== 1 ? 's' : ''}`;
    },

    // Update toolbar button states
    updateToolbarState() {
        const commands = ['bold', 'italic', 'underline', 'strikeThrough', 
                          'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull',
                          'insertUnorderedList', 'insertOrderedList'];
        
        commands.forEach(command => {
            const btn = document.querySelector(`.doc-tool-btn[data-command="${command}"]`);
            if (btn) {
                if (document.queryCommandState(command)) {
                    btn.classList.add('active');
                } else {
                    btn.classList.remove('active');
                }
            }
        });
    },

    // Save document
    async save() {
        const state = AppState;
        
        if (state.currentMode !== 'document') {
            return;
        }
        
        try {
            if (state.docFileHandle !== null && state.docFileHandle !== undefined) {
                try {
                    await this.saveToHandle();
                    this.onSaveSuccess();
                    return;
                } catch (handleError) {
                    console.warn('Could not save to existing handle, will prompt for new location:', handleError);
                    state.docFileHandle = null;
                }
            }
            
            if ('showSaveFilePicker' in window) {
                await this.saveWithPicker();
                this.onSaveSuccess();
            } else {
                await this.exportAsDocx();
                this.onSaveSuccess();
            }
            
        } catch (error) {
            if (error.name === 'AbortError') {
                return;
            }
            console.error('Save error:', error);
            Utils.showToast('Failed to save document: ' + error.message, 'error');
        }
    },

    // Called when document save is successful
    onSaveSuccess() {
        const state = AppState;
        state.docHasUnsavedChanges = false;
        DOM.docEditIndicator.style.display = 'none';
        Utils.showToast('Document saved!', 'success');
        
        if (state.currentFile) {
            DOM.docFileName.textContent = '📄 ' + state.currentFile.name + (state.docFileHandle ? ' ✓' : '');
        }
    },

    // Save with file picker
    async saveWithPicker() {
        const state = AppState;
        const fileName = state.currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
        
        const handle = await window.showSaveFilePicker({
            suggestedName: `${fileName}.docx`,
            types: [{
                description: 'Word Document',
                accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] }
            }]
        });
        
        state.docFileHandle = handle;
        state.currentFile = { name: handle.name, size: 0 };
        
        await this.saveToHandle();
        
        DOM.docFileName.textContent = '📄 ' + handle.name + ' ✓';
        DOM.docFileSize.textContent = 'Saved';
    },

    // Save to existing file handle
    async saveToHandle() {
        const state = AppState;
        
        if (!state.docFileHandle) {
            throw new Error('No file handle available');
        }
        
        // Verify we still have write permission
        const options = { mode: 'readwrite' };
        if ((await state.docFileHandle.queryPermission(options)) !== 'granted') {
            if ((await state.docFileHandle.requestPermission(options)) !== 'granted') {
                throw new Error('Write permission denied');
            }
        }
        
        const { Document, Packer, Paragraph, TextRun, HeadingLevel } = docx;
        
        const htmlContent = DOM.docEditor.innerHTML;
        const paragraphs = this.htmlToDocxParagraphs(htmlContent);
        
        const doc = new Document({
            sections: [{
                properties: {},
                children: paragraphs
            }]
        });
        
        const blob = await Packer.toBlob(doc);
        
        const writable = await state.docFileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
    },

    // Convert HTML to DOCX paragraphs
    htmlToDocxParagraphs(html) {
        const { Paragraph, TextRun, HeadingLevel } = docx;
        const paragraphs = [];
        
        const temp = document.createElement('div');
        temp.innerHTML = html;
        
        const processNode = (node) => {
            if (node.nodeType === Node.TEXT_NODE) {
                const text = node.textContent;
                if (text.trim()) {
                    paragraphs.push(new Paragraph({
                        children: [new TextRun(text)]
                    }));
                }
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                const tagName = node.tagName.toLowerCase();
                const text = node.innerText || '';
                
                switch (tagName) {
                    case 'h1':
                        paragraphs.push(new Paragraph({
                            text: text,
                            heading: HeadingLevel.HEADING_1
                        }));
                        break;
                    case 'h2':
                        paragraphs.push(new Paragraph({
                            text: text,
                            heading: HeadingLevel.HEADING_2
                        }));
                        break;
                    case 'h3':
                        paragraphs.push(new Paragraph({
                            text: text,
                            heading: HeadingLevel.HEADING_3
                        }));
                        break;
                    case 'p':
                    case 'div':
                        const runs = [];
                        this.parseInlineElements(node, runs);
                        if (runs.length > 0) {
                            paragraphs.push(new Paragraph({ children: runs }));
                        }
                        break;
                    case 'ul':
                    case 'ol':
                        node.querySelectorAll('li').forEach(li => {
                            paragraphs.push(new Paragraph({
                                children: [new TextRun('• ' + li.innerText)]
                            }));
                        });
                        break;
                    case 'br':
                        paragraphs.push(new Paragraph({ children: [] }));
                        break;
                    default:
                        node.childNodes.forEach(child => processNode(child));
                }
            }
        };
        
        temp.childNodes.forEach(child => processNode(child));
        
        if (paragraphs.length === 0) {
            paragraphs.push(new Paragraph({ children: [new TextRun('')] }));
        }
        
        return paragraphs;
    },

    // Parse inline elements for text runs
    parseInlineElements(node, runs) {
        const { TextRun } = docx;
        
        node.childNodes.forEach(child => {
            if (child.nodeType === Node.TEXT_NODE) {
                const text = child.textContent;
                if (text) {
                    runs.push(new TextRun(text));
                }
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                const tagName = child.tagName.toLowerCase();
                const text = child.innerText || '';
                
                let options = { text: text };
                
                if (tagName === 'b' || tagName === 'strong') {
                    options.bold = true;
                } else if (tagName === 'i' || tagName === 'em') {
                    options.italics = true;
                } else if (tagName === 'u') {
                    options.underline = {};
                } else if (tagName === 's' || tagName === 'strike') {
                    options.strike = true;
                }
                
                if (text) {
                    runs.push(new TextRun(options));
                }
            }
        });
    },

    // Export document
    async export() {
        const state = AppState;
        
        if (!state.currentFile) {
            Utils.showToast('No document to export', 'error');
            return;
        }
        
        const format = await Utils.prompt('Enter format: 1 = DOCX, 2 = HTML, 3 = TXT', '1', {
            title: 'Export Document',
            icon: '⬇️',
            okText: 'Export',
            placeholder: 'Enter 1, 2, or 3'
        });
        
        switch (format) {
            case '1':
                this.exportAsDocx();
                break;
            case '2':
                this.exportAsHtml();
                break;
            case '3':
                this.exportAsText();
                break;
            default:
                if (format !== null) {
                    Utils.showToast('Invalid format selected', 'warning');
                }
        }
    },

    // Export as DOCX
    async exportAsDocx() {
        const state = AppState;
        const fileName = state.currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
        
        const { Document, Packer, Paragraph, TextRun, HeadingLevel } = docx;
        
        const htmlContent = DOM.docEditor.innerHTML;
        const paragraphs = this.htmlToDocxParagraphs(htmlContent);
        
        const doc = new Document({
            sections: [{
                properties: {},
                children: paragraphs
            }]
        });
        
        const blob = await Packer.toBlob(doc);
        Utils.downloadBlobFile(blob, `${fileName}.docx`);
    },

    // Export as HTML
    exportAsHtml() {
        const state = AppState;
        const fileName = state.currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
        const htmlContent = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>${fileName}</title>
    <style>
        body {
            font-family: 'Times New Roman', serif;
            font-size: 12pt;
            line-height: 1.5;
            max-width: 800px;
            margin: 40px auto;
            padding: 20px;
        }
    </style>
</head>
<body>
${DOM.docEditor.innerHTML}
</body>
</html>`;
        
        Utils.downloadBlob(htmlContent, `${fileName}.html`, 'text/html;charset=utf-8;');
    },

    // Export as plain text
    exportAsText() {
        const state = AppState;
        const fileName = state.currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
        const textContent = DOM.docEditor.innerText || '';
        
        Utils.downloadBlob(textContent, `${fileName}.txt`, 'text/plain;charset=utf-8;');
    }
};

// Export for use in other modules
window.DocumentEditor = DocumentEditor;