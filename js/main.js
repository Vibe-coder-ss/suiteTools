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
    
    // App logo click - return to home
    const appLogo = document.getElementById('appLogo');
    if (appLogo) {
        appLogo.addEventListener('click', () => goHome());
    }
    
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
    
    // Slides handlers
    setupSlidesHandlers();
    
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
        } else if (AppState.currentMode === 'slides') {
            SlidesEditor.save();
        }
        return false;
    }
    
    // Arrow keys for slides navigation (only when not editing)
    if (AppState.currentMode === 'slides' && !e.target.matches('input, textarea, [contenteditable]')) {
        if (e.key === 'ArrowLeft') {
            e.preventDefault();
            SlidesEditor.prevSlide();
        } else if (e.key === 'ArrowRight') {
            e.preventDefault();
            SlidesEditor.nextSlide();
        } else if (e.key === ' ') {
            e.preventDefault();
            SlidesEditor.nextSlide();
        } else if (e.key === 'Delete' || e.key === 'Backspace') {
            // Only delete if not in an editable field
            e.preventDefault();
            SlidesEditor.deleteSlide();
        }
    }
    
    // Ctrl+D to duplicate slide
    if (AppState.currentMode === 'slides' && (e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') {
        e.preventDefault();
        SlidesEditor.duplicateSlide();
    }
    
    // F5 to start slideshow (in slides mode)
    if (AppState.currentMode === 'slides' && e.key === 'F5') {
        e.preventDefault();
        SlidesEditor.startSlideshow();
    }
    
    // Escape to blur current editable element
    if (e.key === 'Escape' && e.target.matches('[contenteditable]')) {
        e.target.blur();
    }
}

// Check if there are any unsaved changes across all modes
function hasAnyUnsavedChanges() {
    return AppState.hasUnsavedChanges || 
           AppState.docHasUnsavedChanges || 
           AppState.slidesHasUnsavedChanges;
}

// Go to home page (show drop zone, hide all editors)
async function goHome() {
    // Check for unsaved changes
    if (hasAnyUnsavedChanges()) {
        const result = await Utils.confirm(
            'You have unsaved changes. Would you like to save before returning home?',
            {
                title: 'Unsaved Changes',
                confirmText: 'Save & Go Home',
                cancelText: 'Discard Changes',
                showThirdOption: true,
                thirdOptionText: 'Cancel'
            }
        );
        
        if (result === 'third' || result === null) {
            // User clicked Cancel or closed the dialog - stay on current page
            return;
        }
        
        if (result === true) {
            // User wants to save first
            try {
                if (AppState.currentMode === 'spreadsheet' && AppState.hasUnsavedChanges) {
                    await FileHandler.save();
                } else if (AppState.currentMode === 'document' && AppState.docHasUnsavedChanges) {
                    await DocumentEditor.save();
                } else if (AppState.currentMode === 'slides' && AppState.slidesHasUnsavedChanges) {
                    await SlidesEditor.save();
                }
            } catch (err) {
                console.error('Error saving:', err);
                // Continue to home even if save failed (user already chose to go home)
            }
        }
        // If result === false, user wants to discard changes - continue to home
    }
    
    // Reset the application state
    resetToHome();
}

// Reset the app to the home/landing state
function resetToHome() {
    // Clear spreadsheet state
    AppState.data = [];
    AppState.columns = [];
    AppState.fileName = null;
    AppState.fileHandle = null;
    AppState.hasUnsavedChanges = false;
    AppState.originalSheetData = [];
    AppState.allSheets = {};
    AppState.currentSheetName = null;
    
    // Clear document state
    AppState.docFile = null;
    AppState.docFileHandle = null;
    AppState.docHasUnsavedChanges = false;
    
    // Clear slides state
    AppState.slidesData = [];
    AppState.currentSlideIndex = 0;
    AppState.slidesFile = null;
    AppState.slidesFileHandle = null;
    AppState.slidesHasUnsavedChanges = false;
    
    // Reset mode
    AppState.currentMode = null;
    
    // Hide all spreadsheet panels and info bars
    if (DOM.tableContainer) DOM.tableContainer.style.display = 'none';
    if (DOM.fileInfo) DOM.fileInfo.style.display = 'none';
    if (DOM.editIndicator) DOM.editIndicator.style.display = 'none';
    if (DOM.quickFilterPanel) DOM.quickFilterPanel.style.display = 'none';
    if (DOM.sqlSection) DOM.sqlSection.style.display = 'none';
    if (DOM.clearBtn) DOM.clearBtn.style.display = 'none';
    if (DOM.sheetTabs) DOM.sheetTabs.innerHTML = '';
    
    // Hide document panel
    if (DOM.documentContainer) DOM.documentContainer.style.display = 'none';
    if (DOM.docInfo) DOM.docInfo.style.display = 'none';
    if (DOM.docEditIndicator) DOM.docEditIndicator.style.display = 'none';
    
    // Hide slides panel
    if (DOM.slidesContainer) DOM.slidesContainer.style.display = 'none';
    if (DOM.slidesInfo) DOM.slidesInfo.style.display = 'none';
    if (DOM.slidesEditIndicator) DOM.slidesEditIndicator.style.display = 'none';
    
    // Clear table content
    const table = document.getElementById('dataTable');
    if (table) table.innerHTML = '';
    
    // Clear document content
    if (DOM.docEditor) DOM.docEditor.innerHTML = '';
    
    // Clear slides content
    const slideCanvas = document.getElementById('slideCanvas');
    if (slideCanvas) slideCanvas.innerHTML = '';
    const thumbnailsContainer = document.getElementById('thumbnailsContainer');
    if (thumbnailsContainer) thumbnailsContainer.innerHTML = '';
    
    // Reset tabs
    const tabs = document.querySelectorAll('.mode-tab');
    tabs.forEach(tab => tab.classList.remove('active'));
    
    // Show upload section (which contains the drop zone)
    if (DOM.uploadSection) DOM.uploadSection.style.display = 'flex';
    
    // Reset file input
    if (DOM.fileInput) DOM.fileInput.value = '';
    
    Utils.showToast('Returned to home', 'info');
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
    DOM.sqlExportResultsBtn.addEventListener('click', async () => {
        if (AppState.hasSqlResults && AppState.sqlResultData) {
            const format = await Utils.prompt('Enter export format:', 'csv', {
                title: 'Export SQL Results',
                icon: '⬇️',
                okText: 'Export',
                placeholder: 'csv, json, or xlsx'
            });
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

// Slides event listeners
function setupSlidesHandlers() {
    // Navigation buttons
    if (DOM.prevSlideBtn) {
        DOM.prevSlideBtn.addEventListener('click', () => SlidesEditor.prevSlide());
    }
    if (DOM.nextSlideBtn) {
        DOM.nextSlideBtn.addEventListener('click', () => SlidesEditor.nextSlide());
    }
    
    // Save and export buttons
    if (DOM.slidesSaveBtn) {
        DOM.slidesSaveBtn.addEventListener('click', () => SlidesEditor.save());
    }
    if (DOM.slidesExportBtn) {
        DOM.slidesExportBtn.addEventListener('click', () => SlidesEditor.export());
    }
    
    // Slideshow button
    const slideshowBtn = document.getElementById('slideshowBtn');
    if (slideshowBtn) {
        slideshowBtn.addEventListener('click', () => SlidesEditor.startSlideshow());
    }
    
    // Add slide button
    const addSlideBtn = document.getElementById('addSlideBtn');
    if (addSlideBtn) {
        addSlideBtn.addEventListener('click', () => SlidesEditor.addSlide());
    }
    
    // Duplicate slide button
    const duplicateSlideBtn = document.getElementById('duplicateSlideBtn');
    if (duplicateSlideBtn) {
        duplicateSlideBtn.addEventListener('click', () => SlidesEditor.duplicateSlide());
    }
    
    // Move slide up button
    const moveSlideUpBtn = document.getElementById('moveSlideUpBtn');
    if (moveSlideUpBtn) {
        moveSlideUpBtn.addEventListener('click', () => SlidesEditor.moveSlideUp());
    }
    
    // Move slide down button
    const moveSlideDownBtn = document.getElementById('moveSlideDownBtn');
    if (moveSlideDownBtn) {
        moveSlideDownBtn.addEventListener('click', () => SlidesEditor.moveSlideDown());
    }
    
    // Delete slide button
    const deleteSlideBtn = document.getElementById('deleteSlideBtn');
    if (deleteSlideBtn) {
        deleteSlideBtn.addEventListener('click', () => SlidesEditor.deleteSlide());
    }
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
    DOM.insertLinkBtn.addEventListener('click', async () => {
        const url = await Utils.prompt('Enter URL:', 'https://', {
            title: 'Insert Link',
            icon: '🔗',
            okText: 'Insert',
            placeholder: 'https://example.com'
        });
        if (url) {
            document.execCommand('createLink', false, url);
            DOM.docEditor.focus();
            DocumentEditor.markAsEdited();
        }
    });
    
    // Insert image
    DOM.insertImageBtn.addEventListener('click', async () => {
        const url = await Utils.prompt('Enter image URL:', 'https://', {
            title: 'Insert Image',
            icon: '🖼️',
            okText: 'Insert',
            placeholder: 'https://example.com/image.jpg'
        });
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
    
    DOM.createSlidesBtn.addEventListener('click', () => {
        DOM.createDropdownMenu.classList.remove('show');
        SlidesEditor.createEmpty();
    });
    
    // Large buttons
    DOM.createNewLargeBtn.addEventListener('click', () => Spreadsheet.createEmpty());
    DOM.createDocLargeBtn.addEventListener('click', () => DocumentEditor.createEmpty());
    DOM.createSlidesLargeBtn.addEventListener('click', () => SlidesEditor.createEmpty());
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
            Utils.alert('Please select at least one sheet to export', {
                title: 'No Sheets Selected',
                icon: '📊'
            });
            return;
        }
    }
    
    FileHandler.exportData(state.selectedFormat, sheetsToExport);
    closeExportModal();
}