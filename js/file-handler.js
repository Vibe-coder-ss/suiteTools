/**
 * File Handler Module
 * Handles file import, export, and processing
 */

const FileHandler = {
    // Process uploaded file
    process(file) {
        const spreadsheetTypes = ['.csv', '.xlsx', '.xls'];
        const documentTypes = ['.doc', '.docx'];
        const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
        
        const allValidTypes = [...spreadsheetTypes, ...documentTypes];
        
        if (!allValidTypes.includes(fileExtension)) {
            alert('Please upload a supported file (.csv, .xlsx, .xls, .doc, .docx)');
            return;
        }

        DOM.uploadSection.style.display = 'none';

        const reader = new FileReader();

        if (spreadsheetTypes.includes(fileExtension)) {
            DOM.tableContainer.style.display = 'block';
            DOM.tableContainer.innerHTML = '<div class="loading">Processing file</div>';
            DOM.documentContainer.style.display = 'none';
            
            if (fileExtension === '.csv') {
                reader.onload = (e) => {
                    const text = e.target.result;
                    this.parseCSV(text, file);
                };
                reader.readAsText(file);
            } else {
                reader.onload = (e) => {
                    const data = new Uint8Array(e.target.result);
                    this.parseExcel(data, file);
                };
                reader.readAsArrayBuffer(file);
            }
        } else if (documentTypes.includes(fileExtension)) {
            DOM.documentContainer.style.display = 'flex';
            DOM.tableContainer.style.display = 'none';
            
            reader.onload = (e) => {
                const arrayBuffer = e.target.result;
                DocumentEditor.parse(arrayBuffer, file);
            };
            reader.readAsArrayBuffer(file);
        }
    },

    // Parse CSV file
    parseCSV(text, file) {
        const state = AppState;
        const lines = text.split(/\r\n|\n/).filter(line => line.trim());
        
        if (lines.length === 0) {
            alert('The file appears to be empty');
            Spreadsheet.clear();
            return;
        }

        const delimiter = Utils.detectDelimiter(lines[0]);
        
        state.headers = Utils.parseCSVLine(lines[0], delimiter);
        state.currentData = lines.slice(1).map(line => Utils.parseCSVLine(line, delimiter));

        state.allSheets = [{ name: 'Sheet1', headers: state.headers, data: state.currentData }];
        state.currentFile = file;
        state.currentSheetIndex = 0;
        
        DOM.sheetTabs.style.display = 'none';

        Spreadsheet.displayData(file);
    },

    // Parse Excel file
    parseExcel(data, file) {
        const state = AppState;
        
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            
            state.allSheets = workbook.SheetNames.map(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                
                if (jsonData.length === 0) {
                    return { name: sheetName, headers: [], data: [] };
                }
                
                return {
                    name: sheetName,
                    headers: jsonData[0].map(h => h?.toString() || ''),
                    data: jsonData.slice(1).map(row => 
                        Array.isArray(row) ? row.map(cell => cell?.toString() || '') : []
                    )
                };
            });

            if (state.allSheets.every(sheet => sheet.data.length === 0 && sheet.headers.length === 0)) {
                alert('The file appears to be empty');
                Spreadsheet.clear();
                return;
            }

            state.currentFile = file;
            state.currentSheetIndex = 0;
            Spreadsheet.loadSheet(0);
            Spreadsheet.displayData(file);
            
            if (state.allSheets.length > 1) {
                Spreadsheet.renderSheetTabs();
            }
        } catch (error) {
            alert('Error reading Excel file. Please try again.');
            Spreadsheet.clear();
            console.error(error);
        }
    },

    // Open file with File System Access API
    async openWithHandle() {
        if (!('showOpenFilePicker' in window)) {
            DOM.fileInput.click();
            return;
        }
        
        try {
            const [handle] = await window.showOpenFilePicker({
                types: [
                    {
                        description: 'Spreadsheet Files',
                        accept: {
                            'text/csv': ['.csv'],
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                            'application/vnd.ms-excel': ['.xls']
                        }
                    },
                    {
                        description: 'Document Files',
                        accept: {
                            'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
                            'application/msword': ['.doc']
                        }
                    }
                ]
            });
            
            const file = await handle.getFile();
            const ext = '.' + file.name.split('.').pop().toLowerCase();
            
            if (['.doc', '.docx'].includes(ext)) {
                AppState.docFileHandle = handle;
                AppState.fileHandle = null;
            } else {
                AppState.fileHandle = handle;
                AppState.docFileHandle = null;
            }
            
            this.process(file);
            
        } catch (err) {
            if (err.name !== 'AbortError') {
                console.error('Error opening file:', err);
                DOM.fileInput.click();
            }
        }
    },

    // Save spreadsheet file
    async save() {
        const state = AppState;
        
        if (!state.currentFile && !state.fileHandle) {
            Utils.showToast('No file to save', 'error');
            return;
        }
        
        this.showSaveStatus('saving', 'Saving...');
        
        try {
            const wb = XLSX.utils.book_new();
            
            state.allSheets.forEach(sheet => {
                const sheetData = [sheet.headers, ...sheet.data];
                const ws = XLSX.utils.aoa_to_sheet(sheetData);
                XLSX.utils.book_append_sheet(wb, ws, sheet.name);
            });
            
            const fileName = state.currentFile?.name || 'data.xlsx';
            const ext = fileName.split('.').pop().toLowerCase();
            
            if (state.fileHandle) {
                await this.saveWithHandle(wb, ext);
            } else {
                await this.saveAsDownload(wb, fileName, ext);
            }
            
            state.hasUnsavedChanges = false;
            DOM.editIndicator.style.display = 'none';
            this.showSaveStatus('saved', 'Saved');
            state.lastSaveTime = new Date();
            
            setTimeout(() => {
                DOM.saveStatus.style.display = 'none';
            }, 2000);
            
        } catch (err) {
            console.error('Save error:', err);
            this.showSaveStatus('error', 'Save failed');
            Utils.showToast('Failed to save file: ' + err.message, 'error');
        }
    },

    // Save using File System Access API
    async saveWithHandle(wb, ext) {
        const state = AppState;
        const writable = await state.fileHandle.createWritable();
        
        let content;
        let bookType;
        
        if (ext === 'csv') {
            content = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
            await writable.write(content);
        } else {
            bookType = ext === 'xls' ? 'biff8' : 'xlsx';
            const buffer = XLSX.write(wb, { bookType: bookType, type: 'array' });
            await writable.write(buffer);
        }
        
        await writable.close();
    },

    // Save as download
    async saveAsDownload(wb, fileName, ext) {
        if (ext === 'csv') {
            const csv = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
            Utils.downloadBlob(csv, fileName, 'text/csv;charset=utf-8;');
        } else {
            const bookType = ext === 'xls' ? 'biff8' : 'xlsx';
            XLSX.writeFile(wb, fileName, { bookType: bookType });
        }
    },

    // Show save status
    showSaveStatus(status, text) {
        DOM.saveStatus.className = 'save-status ' + status;
        document.getElementById('saveText').textContent = text;
        DOM.saveStatus.style.display = 'flex';
    },

    // Trigger auto-save
    triggerAutoSave() {
        const state = AppState;
        
        if (!state.autoSaveEnabled || !state.hasUnsavedChanges) return;
        
        if (state.autoSaveTimeout) {
            clearTimeout(state.autoSaveTimeout);
        }
        
        state.autoSaveTimeout = setTimeout(async () => {
            if (state.autoSaveEnabled && state.hasUnsavedChanges) {
                await this.save();
            }
        }, 1000);
    },

    // Export data
    exportData(format, sheetIndices) {
        const state = AppState;
        if (state.allSheets.length === 0) return;
        
        const baseName = state.currentFile ? state.currentFile.name.replace(/\.[^/.]+$/, '') : 'export';
        const sheetsToExport = sheetIndices.map(i => state.allSheets[i]);
        
        const wb = XLSX.utils.book_new();
        
        sheetsToExport.forEach(sheet => {
            const sheetData = [sheet.headers, ...sheet.data];
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(wb, ws, sheet.name);
        });
        
        let filename;
        
        switch (format) {
            case 'csv':
                if (sheetsToExport.length === 1) {
                    filename = `${baseName}.csv`;
                    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
                    Utils.downloadBlob(csv, filename, 'text/csv;charset=utf-8;');
                } else {
                    sheetsToExport.forEach(sheet => {
                        const wb2 = XLSX.utils.book_new();
                        const sheetData = [sheet.headers, ...sheet.data];
                        const ws = XLSX.utils.aoa_to_sheet(sheetData);
                        XLSX.utils.book_append_sheet(wb2, ws, sheet.name);
                        
                        const fname = `${baseName}_${sheet.name}.csv`;
                        const csv = XLSX.utils.sheet_to_csv(ws);
                        Utils.downloadBlob(csv, fname, 'text/csv;charset=utf-8;');
                    });
                    Utils.showExportSuccess(`${sheetsToExport.length} CSV files`);
                }
                return;
                
            case 'xlsx':
                filename = `${baseName}.xlsx`;
                XLSX.writeFile(wb, filename, { bookType: 'xlsx' });
                break;
                
            case 'xls':
                filename = `${baseName}.xls`;
                XLSX.writeFile(wb, filename, { bookType: 'biff8' });
                break;
                
            case 'json':
                filename = `${baseName}.json`;
                this.exportAsJSON(sheetsToExport, filename);
                return;
                
            case 'html':
                filename = `${baseName}.html`;
                this.exportAsHTML(sheetsToExport, filename);
                return;
                
            default:
                return;
        }
        
        Utils.showExportSuccess(filename);
    },

    // Export as JSON
    exportAsJSON(sheets, filename) {
        let jsonData;
        
        if (sheets.length === 1) {
            const sheet = sheets[0];
            jsonData = {
                sheet: sheet.name,
                headers: sheet.headers,
                data: sheet.data.map(row => {
                    const obj = {};
                    sheet.headers.forEach((h, i) => {
                        obj[h || `Column${i + 1}`] = row[i] || '';
                    });
                    return obj;
                })
            };
        } else {
            jsonData = {
                sheets: sheets.map(sheet => ({
                    name: sheet.name,
                    headers: sheet.headers,
                    data: sheet.data.map(row => {
                        const obj = {};
                        sheet.headers.forEach((h, i) => {
                            obj[h || `Column${i + 1}`] = row[i] || '';
                        });
                        return obj;
                    })
                }))
            };
        }
        
        Utils.downloadBlob(JSON.stringify(jsonData, null, 2), filename, 'application/json;charset=utf-8;');
    },

    // Export as HTML
    exportAsHTML(sheets, filename) {
        let tabsHtml = '';
        let contentHtml = '';
        
        if (sheets.length > 1) {
            tabsHtml = `
            <div class="tabs">
                ${sheets.map((sheet, i) => `
                    <button class="tab-btn ${i === 0 ? 'active' : ''}" onclick="showSheet(${i})">${sheet.name}</button>
                `).join('')}
            </div>`;
        }
        
        sheets.forEach((sheet, index) => {
            const wb = XLSX.utils.book_new();
            const sheetData = [sheet.headers, ...sheet.data];
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            const tableHtml = XLSX.utils.sheet_to_html(ws);
            
            contentHtml += `
            <div class="sheet-content" id="sheet${index}" style="${index === 0 ? '' : 'display: none;'}">
                <h2>${sheet.name}</h2>
                ${tableHtml}
            </div>`;
        });
        
        const fullHtml = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>${filename}</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }
        .tabs { display: flex; gap: 5px; margin-bottom: 20px; }
        .tab-btn { padding: 10px 20px; border: none; background: #e0e0e0; cursor: pointer; border-radius: 8px 8px 0 0; }
        .tab-btn.active { background: #667eea; color: white; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background: #667eea; color: white; }
        tr:nth-child(even) { background: #f9f9f9; }
    </style>
</head>
<body>
${tabsHtml}
${contentHtml}
<script>
function showSheet(index) {
    document.querySelectorAll('.sheet-content').forEach((el, i) => {
        el.style.display = i === index ? 'block' : 'none';
    });
    document.querySelectorAll('.tab-btn').forEach((btn, i) => {
        btn.classList.toggle('active', i === index);
    });
}
</script>
</body>
</html>`;
        Utils.downloadBlob(fullHtml, filename, 'text/html;charset=utf-8;');
    }
};

// Export for use in other modules
window.FileHandler = FileHandler;