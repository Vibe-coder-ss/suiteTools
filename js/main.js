/**
 * Main Application Entry Point
 * Initializes all modules and sets up event listeners
 */

// Initialize application when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    // Initialize DOM references
    DOM.init();
    
    // Initialize theme
    Theme.init();
    
    // Setup all event listeners
    setupEventListeners();
});

// Setup all event listeners
function setupEventListeners() {
    // Theme toggle
    DOM.themeToggle.addEventListener('click', () => Theme.toggle());
    
    // Global keyboard shortcuts
    document.addEventListener('keydown', handleGlobalKeydown);
    
    // File handling
    setupFileHandlers();
    
    // Spreadsheet handlers
    setupSpreadsheetHandlers();
    
    // SQL handlers
    setupSQLHandlers();
    
    // Filter handlers
    setupFilterHandlers();
    
    // Document handlers
    setupDocumentHandlers();
    
    // Export handlers
    setupExportHandlers();
    
    // Create new handlers
    setupCreateHandlers();
}

// Global keyboard shortcuts
function handleGlobalKeydown(e) {
    // Ctrl/Cmd + S to save
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        
        if (e.target === DOM.docEditor) {
            return; // Handled by document editor
        }
        
        if (AppState.currentMode === 'document') {
            DocumentEditor.save();
        } else if (AppState.currentMode === 'spreadsheet') {
            FileHandler.save();
        }
        return false;
    }
}

// File handling event listeners
function setupFileHandlers() {
    // Drag and drop
    DOM.dropZone.addEventListener('click', (e) => {
        if (e.target.tagName !== 'LABEL' && !e.target.closest('label')) {
            DOM.fileInput.click();
        }
    });
    
    DOM.dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        DOM.dropZone.classList.add('dragover');
    });
    
    DOM.dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        DOM.dropZone.classList.remove('dragover');
    });
    
    DOM.dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        DOM.dropZone.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            FileHandler.process(files[0]);
        }
    });
    
    // File input
    DOM.fileInput.addEventListener('change', (e) => {
        const files = e.target.files;
        if (files.length > 0) {
            if ('showOpenFilePicker' in window) {
                AppState.fileHandle = null;
            }
            FileHandler.process(files[0]);
        }
    });
    
    // Browse button
    document.getElementById('browseBtn').addEventListener('click', () => {
        if ('showOpenFilePicker' in window) {
            FileHandler.openWithHandle();
        } else {
            DOM.fileInput.click();
        }
    });
    
    // Clear button
    DOM.clearBtn.addEventListener('click', () => Spreadsheet.clear());
}

// Spreadsheet event listeners
function setupSpreadsheetHandlers() {
    // Add row/column buttons
    DOM.addRowBtn.addEventListener('click', () => Spreadsheet.addRow());
    DOM.addColBtn.addEventListener('click', () => Spreadsheet.addColumn());
    
    // Save and undo buttons
    DOM.saveBtn.addEventListener('click', () => FileHandler.save());
    DOM.undoAllBtn.addEventListener('click', () => Spreadsheet.undoAllChanges());
    
    // Auto-save toggle
    DOM.autosaveCheckbox.addEventListener('change', (e) => {
        AppState.autoSaveEnabled = e.target.checked;
        if (AppState.autoSaveEnabled && AppState.hasUnsavedChanges) {
            FileHandler.triggerAutoSave();
        }
    });
    
    // Search
    DOM.searchInput.addEventListener('input', () => Filter.handleSearch());
    
    // Table cell and header interactions
    DOM.tableContainer.addEventListener('click', handleTableClick);
    DOM.tableContainer.addEventListener('dblclick', handleTableDoubleClick);
}

// Handle table clicks
function handleTableClick(e) {
    // Delete button
    const deleteBtn = e.target.closest('.delete-row-btn');
    if (deleteBtn) {
        const originalRowIndex = parseInt(deleteBtn.dataset.originalRow);
        Spreadsheet.deleteRow(originalRowIndex);
        return;
    }
    
    // Header click for sorting
    const th = e.target.closest('th.sortable');
    if (th && !th.classList.contains('editing')) {
        const colIndex = parseInt(th.dataset.col);
        if (!isNaN(colIndex)) {
            if (AppState.sortColumn === colIndex) {
                if (AppState.sortDirection === 'asc') {
                    AppState.sortDirection = 'desc';
                } else if (AppState.sortDirection === 'desc') {
                    AppState.sortDirection = 'none';
                    AppState.sortColumn = -1;
                }
            } else {
                AppState.sortColumn = colIndex;
                AppState.sortDirection = 'asc';
            }
            Spreadsheet.sortAndRender();
        }
        return;
    }
    
    // Cell click for selection
    const cell = e.target.closest('td:not(.row-num)');
    if (cell && !cell.classList.contains('editing')) {
        document.querySelectorAll('.data-table td.selected').forEach(td => {
            td.classList.remove('selected');
        });
        cell.classList.add('selected');
    }
}

// Handle table double clicks
function handleTableDoubleClick(e) {
    // Header double-click for renaming
    const th = e.target.closest('th.sortable');
    if (th && !th.classList.contains('editing')) {
        if (AppState.hasSqlResults) {
            Utils.showToast('Cannot rename columns in SQL results. Reset first.', 'warning');
            return;
        }
        const colIndex = parseInt(th.dataset.col);
        if (!isNaN(colIndex)) {
            e.stopPropagation();
            Spreadsheet.startHeaderEdit(th, colIndex);
        }
        return;
    }
    
    // Cell double-click for editing
    const cell = e.target.closest('td:not(.row-num)');
    if (cell && !cell.classList.contains('editing')) {
        if (!cell.classList.contains('editable')) {
            Utils.showToast('Cannot edit filtered or SQL results. Clear search/reset to edit.', 'warning');
            return;
        }
        Spreadsheet.startCellEdit(cell);
    }
}

// SQL event listeners
function setupSQLHandlers() {
    DOM.sqlRunBtn.addEventListener('click', () => SQL.run());
    DOM.sqlExamplesBtn.addEventListener('click', () => SQL.toggleExamplesPanel());
    DOM.sqlResetBtn.addEventListener('click', () => SQL.reset());
    DOM.sqlExpandBtn.addEventListener('click', () => SQL.toggleExpand());
    
    // SQL input handlers
    DOM.sqlInput.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'Enter') {
            e.preventDefault();
            SQL.run();
        }
        if (e.key === 'Escape' && DOM.sqlEditorWrapper.classList.contains('expanded')) {
            SQL.toggleExpand();
        }
    });
    
    DOM.sqlInput.addEventListener('input', () => {
        SQL.autoResizeEditor();
        SQL.updateLineCount();
    });
    
    DOM.sqlInput.addEventListener('paste', () => {
        setTimeout(() => {
            SQL.autoResizeEditor();
            SQL.updateLineCount();
        }, 0);
    });
    
    // Examples panel click handler
    DOM.sqlExamplesPanel.addEventListener('click', (e) => {
        const item = e.target.closest('.example-item');
        if (item) {
            DOM.sqlInput.value = item.dataset.query;
            DOM.sqlExamplesPanel.style.display = 'none';
            DOM.sqlInput.focus();
        }
        
        const badge = e.target.closest('.table-badge');
        if (badge) {
            const tableName = badge.dataset.table;
            const cursorPos = DOM.sqlInput.selectionStart;
            const currentValue = DOM.sqlInput.value;
            DOM.sqlInput.value = currentValue.slice(0, cursorPos) + tableName + currentValue.slice(cursorPos);
            DOM.sqlInput.focus();
            DOM.sqlInput.setSelectionRange(cursorPos + tableName.length, cursorPos + tableName.length);
        }
    });
    
    // Export SQL results
    DOM.sqlExportResultsBtn.addEventListener('click', () => {
        // TODO: Implement SQL results export modal
        if (AppState.hasSqlResults && AppState.sqlResultData) {
            const format = prompt('Export format (csv, json, xlsx):', 'csv');
            if (format) {
                // Simple export of SQL results
                const sheet = {
                    name: 'SQL_Results',
                    headers: AppState.sqlResultData.headers,
                    data: AppState.sqlResultData.data
                };
                FileHandler.exportData(format, [0]);
            }
        }
    });
}

// Filter event listeners
function setupFilterHandlers() {
    DOM.filterToggleBtn.addEventListener('click', () => Filter.togglePanel());
    DOM.filterApplyBtn.addEventListener('click', () => Filter.apply());
    DOM.filterClearBtn.addEventListener('click', () => Filter.clearAll());
    DOM.filterOperator.addEventListener('change', () => Filter.updateValueVisibility());
    DOM.filterValue.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') Filter.apply();
    });
}

// Document editor event listeners
function setupDocumentHandlers() {
    // Toolbar buttons
    DOM.docToolbar.addEventListener('click', (e) => {
        const btn = e.target.closest('.doc-tool-btn');
        if (btn && btn.dataset.command) {
            e.preventDefault();
            document.execCommand(btn.dataset.command, false, null);
            DOM.docEditor.focus();
            DocumentEditor.updateToolbarState();
            DocumentEditor.markAsEdited();
        }
    });
    
    // Font family
    DOM.fontSelect.addEventListener('change', (e) => {
        document.execCommand('fontName', false, e.target.value);
        DOM.docEditor.focus();
        DocumentEditor.markAsEdited();
    });
    
    // Font size
    DOM.sizeSelect.addEventListener('change', (e) => {
        document.execCommand('fontSize', false, e.target.value);
        DOM.docEditor.focus();
        DocumentEditor.markAsEdited();
    });
    
    // Text color
    DOM.textColorPicker.addEventListener('input', (e) => {
        document.execCommand('foreColor', false, e.target.value);
        DOM.docEditor.focus();
        DocumentEditor.markAsEdited();
    });
    
    // Background color
    DOM.bgColorPicker.addEventListener('input', (e) => {
        document.execCommand('hiliteColor', false, e.target.value);
        DOM.docEditor.focus();
        DocumentEditor.markAsEdited();
    });
    
    // Insert link
    DOM.insertLinkBtn.addEventListener('click', () => {
        const url = prompt('Enter URL:', 'https://');
        if (url) {
            document.execCommand('createLink', false, url);
            DOM.docEditor.focus();
            DocumentEditor.markAsEdited();
        }
    });
    
    // Insert image
    DOM.insertImageBtn.addEventListener('click', () => {
        const url = prompt('Enter image URL:', 'https://');
        if (url) {
            document.execCommand('insertImage', false, url);
            DOM.docEditor.focus();
            DocumentEditor.markAsEdited();
        }
    });
    
    // Editor input handler
    DOM.docEditor.addEventListener('input', () => {
        DocumentEditor.markAsEdited();
        DocumentEditor.updateStats();
    });
    
    // Toolbar state updates
    DOM.docEditor.addEventListener('mouseup', () => DocumentEditor.updateToolbarState());
    DOM.docEditor.addEventListener('keyup', () => DocumentEditor.updateToolbarState());
    
    // Editor keyboard shortcuts
    DOM.docEditor.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
            e.preventDefault();
            e.stopPropagation();
            setTimeout(() => DocumentEditor.save(), 0);
            return false;
        }
        
        if ((e.ctrlKey || e.metaKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
            setTimeout(() => {
                DocumentEditor.markAsEdited();
                DocumentEditor.updateToolbarState();
            }, 10);
        }
    });
    
    // Save and export buttons
    DOM.docSaveBtn.addEventListener('click', () => DocumentEditor.save());
    DOM.docExportBtn.addEventListener('click', () => DocumentEditor.export());
}

// Export modal event listeners
function setupExportHandlers() {
    DOM.exportBtn.addEventListener('click', openExportModal);
    DOM.modalClose.addEventListener('click', closeExportModal);
    DOM.exportCancel.addEventListener('click', closeExportModal);
    DOM.exportConfirm.addEventListener('click', confirmExport);
    
    DOM.exportModal.addEventListener('click', (e) => {
        if (e.target === DOM.exportModal) closeExportModal();
    });
    
    DOM.formatGrid.addEventListener('click', (e) => {
        const option = e.target.closest('.format-option');
        if (option) {
            document.querySelectorAll('.format-option').forEach(o => o.classList.remove('selected'));
            option.classList.add('selected');
            AppState.selectedFormat = option.dataset.format;
            updateFormatNote();
        }
    });
    
    document.querySelectorAll('input[name="sheetOption"]').forEach(radio => {
        radio.addEventListener('change', updateFormatNote);
    });
}

// Create new handlers
function setupCreateHandlers() {
    // Dropdown toggle
    DOM.createDropdownBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        DOM.createDropdownMenu.classList.toggle('show');
    });
    
    document.addEventListener('click', () => {
        DOM.createDropdownMenu.classList.remove('show');
    });
    
    // Create options
    DOM.createSheetBtn.addEventListener('click', () => {
        DOM.createDropdownMenu.classList.remove('show');
        Spreadsheet.createEmpty();
    });
    
    DOM.createDocBtn.addEventListener('click', () => {
        DOM.createDropdownMenu.classList.remove('show');
        DocumentEditor.createEmpty();
    });
    
    // Large buttons
    DOM.createNewLargeBtn.addEventListener('click', () => Spreadsheet.createEmpty());
    DOM.createDocLargeBtn.addEventListener('click', () => DocumentEditor.createEmpty());
}

// Export modal functions
function openExportModal() {
    const state = AppState;
    if (state.allSheets.length === 0) return;
    
    DOM.currentSheetNameBadge.textContent = state.allSheets[state.currentSheetIndex].name;
    
    if (state.allSheets.length > 1) {
        DOM.sheetCountBadge.textContent = `${state.allSheets.length} sheets`;
        DOM.allSheetsOption.style.display = 'flex';
        DOM.sheetsChecklist.style.display = 'block';
        
        DOM.sheetCheckboxes.innerHTML = state.allSheets.map((sheet, index) => `
            <label class="checkbox-option">
                <input type="checkbox" name="sheetCheck" value="${index}" checked>
                <span>${sheet.name}</span>
                <span style="color: #888; font-size: 0.8rem;">(${sheet.data.length} rows)</span>
            </label>
        `).join('');
    } else {
        DOM.allSheetsOption.style.display = 'none';
        DOM.sheetsChecklist.style.display = 'none';
    }
    
    state.selectedFormat = 'xlsx';
    document.querySelectorAll('.format-option').forEach(o => {
        o.classList.toggle('selected', o.dataset.format === 'xlsx');
    });
    document.querySelector('input[name="sheetOption"][value="current"]').checked = true;
    
    updateFormatNote();
    DOM.exportModal.classList.add('show');
}

function closeExportModal() {
    DOM.exportModal.classList.remove('show');
}

function updateFormatNote() {
    const state = AppState;
    const sheetOption = document.querySelector('input[name="sheetOption"]:checked').value;
    const isMultiSheet = sheetOption === 'all' || sheetOption === 'selected';
    
    if (state.selectedFormat === 'csv') {
        if (isMultiSheet && state.allSheets.length > 1) {
            DOM.formatNote.textContent = '⚠️ CSV format will create separate files for each sheet';
            DOM.formatNote.style.display = 'block';
        } else {
            DOM.formatNote.style.display = 'none';
        }
    } else if (state.selectedFormat === 'json' || state.selectedFormat === 'html') {
        if (isMultiSheet && state.allSheets.length > 1) {
            DOM.formatNote.textContent = `ℹ️ ${state.selectedFormat.toUpperCase()} will include all selected sheets`;
            DOM.formatNote.style.display = 'block';
        } else {
            DOM.formatNote.style.display = 'none';
        }
    } else {
        DOM.formatNote.style.display = 'none';
    }
}

function confirmExport() {
    const state = AppState;
    const sheetOption = document.querySelector('input[name="sheetOption"]:checked').value;
    
    let sheetsToExport = [];
    
    if (sheetOption === 'current') {
        sheetsToExport = [state.currentSheetIndex];
    } else if (sheetOption === 'all') {
        sheetsToExport = state.allSheets.map((_, i) => i);
    } else if (sheetOption === 'selected') {
        sheetsToExport = Array.from(document.querySelectorAll('input[name="sheetCheck"]:checked'))
            .map(cb => parseInt(cb.value));
        
        if (sheetsToExport.length === 0) {
            alert('Please select at least one sheet to export');
            return;
        }
    }
    
    FileHandler.exportData(state.selectedFormat, sheetsToExport);
    closeExportModal();
}