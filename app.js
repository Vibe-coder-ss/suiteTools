// Theme handling
const themeToggle = document.getElementById('themeToggle');

// Initialize theme from localStorage or system preference
function initTheme() {
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme) {
        document.documentElement.setAttribute('data-theme', savedTheme);
    } else if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
        document.documentElement.setAttribute('data-theme', 'dark');
    }
}

// Toggle theme
function toggleTheme() {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
}

// Initialize theme on load
initTheme();

// Theme toggle event listener
themeToggle.addEventListener('click', toggleTheme);

// Global keyboard shortcuts
document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + S to save
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        
        // Skip if already handled by document editor
        if (e.target === docEditor) {
            return;
        }
        
        if (currentMode === 'document') {
            saveDocument();
        } else if (currentMode === 'spreadsheet') {
            saveFile();
        }
        return false;
    }
});

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const rowCount = document.getElementById('rowCount');
const clearBtn = document.getElementById('clearBtn');
const searchInput = document.getElementById('searchInput');
const tableContainer = document.getElementById('tableContainer');
const uploadSection = document.getElementById('uploadSection');
const sheetTabs = document.getElementById('sheetTabs');
const exportBtn = document.getElementById('exportBtn');
const exportModal = document.getElementById('exportModal');
const modalClose = document.getElementById('modalClose');
const exportCancel = document.getElementById('exportCancel');
const exportConfirm = document.getElementById('exportConfirm');
const formatGrid = document.getElementById('formatGrid');
const formatNote = document.getElementById('formatNote');
const sheetCheckboxes = document.getElementById('sheetCheckboxes');
const currentSheetNameBadge = document.getElementById('currentSheetName');
const sheetCountBadge = document.getElementById('sheetCountBadge');
const allSheetsOption = document.getElementById('allSheetsOption');
const sheetsChecklist = document.getElementById('sheetsChecklist');

// SQL elements
const sqlSection = document.getElementById('sqlSection');
const sqlInput = document.getElementById('sqlInput');
const sqlRunBtn = document.getElementById('sqlRunBtn');
const sqlExamplesBtn = document.getElementById('sqlExamplesBtn');
const sqlExamplesPanel = document.getElementById('sqlExamplesPanel');
const sqlStatus = document.getElementById('sqlStatus');

let selectedFormat = 'xlsx';
let sqlTablesRegistered = false;  // Track if SQL tables are set up
let hasSqlResults = false;        // Track if we have SQL query results
let sqlResultData = null;         // Store SQL result data for export
let hasUnsavedChanges = false;    // Track if data has been edited
let originalSheetData = [];       // Store original data for undo
let fileHandle = null;            // Store file handle for saving
let autoSaveEnabled = false;      // Auto-save toggle state
let autoSaveTimeout = null;       // Debounce auto-save
let lastSaveTime = null;          // Track last save time
let sortColumn = -1;              // Current sort column index (-1 = no sort)
let sortDirection = 'none';       // 'asc', 'desc', or 'none'
let activeFilters = [];           // Store active quick filters

let currentData = [];
let headers = [];
let allSheets = [];        // Store all sheets data
let currentSheetIndex = 0; // Track active sheet
let currentFile = null;    // Store file reference
let currentMode = 'none';  // 'spreadsheet', 'document', or 'none'
let docHasUnsavedChanges = false; // Track document edits
let originalDocContent = ''; // Store original doc content for undo
let docFileHandle = null; // Store document file handle for saving

// Event Listeners
dropZone.addEventListener('click', (e) => {
    // Prevent double-trigger when clicking on the label/button
    if (e.target.tagName !== 'LABEL' && !e.target.closest('label')) {
        fileInput.click();
    }
});
dropZone.addEventListener('dragover', handleDragOver);
dropZone.addEventListener('dragleave', handleDragLeave);
dropZone.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);

// Browse button - use File System Access API if available
document.getElementById('browseBtn').addEventListener('click', () => {
    if ('showOpenFilePicker' in window) {
        openFileWithHandle();
    } else {
        fileInput.click();
    }
});
clearBtn.addEventListener('click', clearData);
searchInput.addEventListener('input', handleSearch);

// Export modal handlers
exportBtn.addEventListener('click', openExportModal);
modalClose.addEventListener('click', closeExportModal);
exportCancel.addEventListener('click', closeExportModal);
exportConfirm.addEventListener('click', confirmExport);

// Create New dropdown handlers
const createDropdownBtn = document.getElementById('createDropdownBtn');
const createDropdownMenu = document.getElementById('createDropdownMenu');

createDropdownBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    createDropdownMenu.classList.toggle('show');
});

// Close dropdown when clicking outside
document.addEventListener('click', () => {
    createDropdownMenu.classList.remove('show');
});

// Create new spreadsheet
document.getElementById('createSheetBtn').addEventListener('click', () => {
    createDropdownMenu.classList.remove('show');
    createEmptySpreadsheet();
});

// Create new document
document.getElementById('createDocBtn').addEventListener('click', () => {
    createDropdownMenu.classList.remove('show');
    createEmptyDocument();
});

// Large buttons in upload section
document.getElementById('createNewLargeBtn').addEventListener('click', createEmptySpreadsheet);
document.getElementById('createDocLargeBtn').addEventListener('click', createEmptyDocument);

// Add column button
document.getElementById('addColBtn').addEventListener('click', addNewColumn);

// Document editor elements
const documentContainer = document.getElementById('documentContainer');
const docEditor = document.getElementById('docEditor');
const docInfo = document.getElementById('docInfo');
const docToolbar = document.getElementById('docToolbar');

// Document toolbar handlers
docToolbar.addEventListener('click', (e) => {
    const btn = e.target.closest('.doc-tool-btn');
    if (btn && btn.dataset.command) {
        e.preventDefault();
        document.execCommand(btn.dataset.command, false, null);
        docEditor.focus();
        updateToolbarState();
        markDocAsEdited();
    }
});

// Font family select
document.getElementById('fontSelect').addEventListener('change', (e) => {
    document.execCommand('fontName', false, e.target.value);
    docEditor.focus();
    markDocAsEdited();
});

// Font size select
document.getElementById('sizeSelect').addEventListener('change', (e) => {
    document.execCommand('fontSize', false, e.target.value);
    docEditor.focus();
    markDocAsEdited();
});

// Text color picker
document.getElementById('textColorPicker').addEventListener('input', (e) => {
    document.execCommand('foreColor', false, e.target.value);
    docEditor.focus();
    markDocAsEdited();
});

// Background color picker
document.getElementById('bgColorPicker').addEventListener('input', (e) => {
    document.execCommand('hiliteColor', false, e.target.value);
    docEditor.focus();
    markDocAsEdited();
});

// Insert link button
document.getElementById('insertLinkBtn').addEventListener('click', () => {
    const url = prompt('Enter URL:', 'https://');
    if (url) {
        document.execCommand('createLink', false, url);
        docEditor.focus();
        markDocAsEdited();
    }
});

// Insert image button
document.getElementById('insertImageBtn').addEventListener('click', () => {
    const url = prompt('Enter image URL:', 'https://');
    if (url) {
        document.execCommand('insertImage', false, url);
        docEditor.focus();
        markDocAsEdited();
    }
});

// Document editor input handler
docEditor.addEventListener('input', () => {
    markDocAsEdited();
    updateDocStats();
});

// Document editor keyboard shortcuts
docEditor.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + S - prevent default browser save and save document
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        e.stopPropagation();
        // Use setTimeout to ensure any pending input events complete first
        setTimeout(() => saveDocument(), 0);
        return false;
    }
    
    // Let other formatting shortcuts work naturally (Ctrl+B, I, U)
    // They're handled by the browser's contenteditable, but mark as edited
    if ((e.ctrlKey || e.metaKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) {
        setTimeout(() => {
            markDocAsEdited();
            updateToolbarState();
        }, 10);
    }
});

// Update toolbar state on selection change
docEditor.addEventListener('mouseup', updateToolbarState);
docEditor.addEventListener('keyup', updateToolbarState);

// Document save button
document.getElementById('docSaveBtn').addEventListener('click', saveDocument);

// Document export button
document.getElementById('docExportBtn').addEventListener('click', exportDocument);

// Close modal on overlay click
exportModal.addEventListener('click', (e) => {
    if (e.target === exportModal) closeExportModal();
});

// Format selection
formatGrid.addEventListener('click', (e) => {
    const option = e.target.closest('.format-option');
    if (option) {
        document.querySelectorAll('.format-option').forEach(o => o.classList.remove('selected'));
        option.classList.add('selected');
        selectedFormat = option.dataset.format;
        updateFormatNote();
    }
});

// Sheet option change - update note
document.querySelectorAll('input[name="sheetOption"]').forEach(radio => {
    radio.addEventListener('change', updateFormatNote);
});

// Undo all changes button
document.getElementById('undoAllBtn').addEventListener('click', undoAllChanges);

// Add row button
document.getElementById('addRowBtn').addEventListener('click', addNewRow);

// Save button
document.getElementById('saveBtn').addEventListener('click', saveFile);

// Auto-save toggle
document.getElementById('autosaveCheckbox').addEventListener('change', (e) => {
    autoSaveEnabled = e.target.checked;
    if (autoSaveEnabled && hasUnsavedChanges) {
        triggerAutoSave();
    }
});

// Table cell editing - delegated event listener
tableContainer.addEventListener('click', handleTableClick);
tableContainer.addEventListener('dblclick', handleCellDoubleClick);

// Column header click for sorting
tableContainer.addEventListener('click', handleHeaderClick);

// Column header double-click for renaming
tableContainer.addEventListener('dblclick', handleHeaderDoubleClick);

// Quick Filter controls
const filterToggleBtn = document.getElementById('filterToggleBtn');
const quickFilterPanel = document.getElementById('quickFilterPanel');
const filterColumn = document.getElementById('filterColumn');
const filterOperator = document.getElementById('filterOperator');
const filterValue = document.getElementById('filterValue');
const filterApplyBtn = document.getElementById('filterApplyBtn');
const filterClearBtn = document.getElementById('filterClearBtn');

filterToggleBtn.addEventListener('click', toggleFilterPanel);
filterApplyBtn.addEventListener('click', applyQuickFilter);
filterClearBtn.addEventListener('click', clearAllFilters);
filterOperator.addEventListener('change', updateFilterValueVisibility);
filterValue.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') applyQuickFilter();
});

// Handle all table clicks
function handleTableClick(e) {
    // Check if delete button was clicked
    const deleteBtn = e.target.closest('.delete-row-btn');
    if (deleteBtn) {
        const originalRowIndex = parseInt(deleteBtn.dataset.originalRow);
        deleteRow(originalRowIndex);
        return;
    }
    
    // Otherwise handle cell click
    handleCellClick(e);
}

// SQL event listeners
sqlRunBtn.addEventListener('click', runSqlQuery);
sqlExamplesBtn.addEventListener('click', toggleExamplesPanel);
document.getElementById('sqlResetBtn').addEventListener('click', resetToOriginalData);
document.getElementById('sqlExportResultsBtn').addEventListener('click', openExportResultsModal);

// SQL editor expand/collapse
const sqlEditorWrapper = document.getElementById('sqlEditorWrapper');
const sqlExpandBtn = document.getElementById('sqlExpandBtn');

sqlExpandBtn.addEventListener('click', toggleSqlEditorExpand);

// Run SQL on Ctrl+Enter, Escape to collapse
sqlInput.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'Enter') {
        e.preventDefault();
        runSqlQuery();
    }
    if (e.key === 'Escape' && sqlEditorWrapper.classList.contains('expanded')) {
        toggleSqlEditorExpand();
    }
});

// Auto-resize SQL editor and update line count
const sqlLineCount = document.getElementById('sqlLineCount');

sqlInput.addEventListener('input', handleSqlEditorInput);
sqlInput.addEventListener('paste', () => setTimeout(handleSqlEditorInput, 0));
sqlInput.addEventListener('keyup', updateSqlLineCount);
sqlInput.addEventListener('click', updateSqlLineCount);

function handleSqlEditorInput() {
    autoResizeSqlEditor();
    updateSqlLineCount();
}

function autoResizeSqlEditor() {
    // Don't auto-resize in expanded mode
    if (sqlEditorWrapper.classList.contains('expanded')) return;
    
    // Reset height to auto to get the correct scrollHeight
    sqlInput.style.height = 'auto';
    
    // Calculate new height (with min and max limits)
    const minHeight = 60;
    const maxHeight = 300;
    const newHeight = Math.min(Math.max(sqlInput.scrollHeight, minHeight), maxHeight);
    
    sqlInput.style.height = newHeight + 'px';
}

function updateSqlLineCount() {
    const text = sqlInput.value;
    if (!text) {
        sqlLineCount.textContent = '';
        return;
    }
    
    const lines = text.split('\n').length;
    const chars = text.length;
    sqlLineCount.textContent = `${lines} line${lines !== 1 ? 's' : ''} • ${chars} char${chars !== 1 ? 's' : ''}`;
}

function toggleSqlEditorExpand() {
    const isExpanded = sqlEditorWrapper.classList.toggle('expanded');
    
    if (isExpanded) {
        // Store original height and expand
        sqlInput.dataset.originalHeight = sqlInput.style.height;
        sqlInput.style.height = '';
        document.body.style.overflow = 'hidden';
    } else {
        // Restore original height
        sqlInput.style.height = sqlInput.dataset.originalHeight || '';
        document.body.style.overflow = '';
        autoResizeSqlEditor();
    }
    
    sqlInput.focus();
}

// Example click handler
document.getElementById('sqlExamplesPanel').addEventListener('click', (e) => {
    const item = e.target.closest('.example-item');
    if (item) {
        sqlInput.value = item.dataset.query;
        sqlExamplesPanel.style.display = 'none';
        sqlInput.focus();
    }
    
    // Table badge click - insert table name
    const badge = e.target.closest('.table-badge');
    if (badge) {
        const tableName = badge.dataset.table;
        // Insert at cursor or append
        const cursorPos = sqlInput.selectionStart;
        const currentValue = sqlInput.value;
        sqlInput.value = currentValue.slice(0, cursorPos) + tableName + currentValue.slice(cursorPos);
        sqlInput.focus();
        sqlInput.setSelectionRange(cursorPos + tableName.length, cursorPos + tableName.length);
    }
});

// Drag and Drop handlers
function handleDragOver(e) {
    e.preventDefault();
    dropZone.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    dropZone.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        // Try to get file handle for saving (File System Access API)
        if ('showOpenFilePicker' in window) {
            // Store that we don't have a handle from input element
            fileHandle = null;
        }
        processFile(files[0]);
    }
}

// Open file with File System Access API (enables saving back)
async function openFileWithHandle() {
    if (!('showOpenFilePicker' in window)) {
        // Fallback to regular file input
        fileInput.click();
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
        
        // Store appropriate handle based on file type
        if (['.doc', '.docx'].includes(ext)) {
            docFileHandle = handle;
            fileHandle = null;
        } else {
            fileHandle = handle;
            docFileHandle = null;
        }
        
        processFile(file);
        
    } catch (err) {
        if (err.name !== 'AbortError') {
            console.error('Error opening file:', err);
            // Fallback to regular file input
            fileInput.click();
        }
    }
}

// Process uploaded file
function processFile(file) {
    const spreadsheetTypes = ['.csv', '.xlsx', '.xls'];
    const documentTypes = ['.doc', '.docx'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    const allValidTypes = [...spreadsheetTypes, ...documentTypes];
    
    if (!allValidTypes.includes(fileExtension)) {
        alert('Please upload a supported file (.csv, .xlsx, .xls, .doc, .docx)');
        return;
    }

    // Show loading state
    uploadSection.style.display = 'none';

    const reader = new FileReader();

    if (spreadsheetTypes.includes(fileExtension)) {
        // Handle spreadsheet files
        tableContainer.style.display = 'block';
        tableContainer.innerHTML = '<div class="loading">Processing file</div>';
        documentContainer.style.display = 'none';
        
        if (fileExtension === '.csv') {
            reader.onload = (e) => {
                const text = e.target.result;
                parseCSV(text, file);
            };
            reader.readAsText(file);
        } else {
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                parseExcel(data, file);
            };
            reader.readAsArrayBuffer(file);
        }
    } else if (documentTypes.includes(fileExtension)) {
        // Handle document files
        documentContainer.style.display = 'flex';
        tableContainer.style.display = 'none';
        
        reader.onload = (e) => {
            const arrayBuffer = e.target.result;
            parseDocument(arrayBuffer, file);
        };
        reader.readAsArrayBuffer(file);
    }
}

// Parse CSV file
function parseCSV(text, file) {
    const lines = text.split(/\r\n|\n/).filter(line => line.trim());
    
    if (lines.length === 0) {
        alert('The file appears to be empty');
        clearData();
        return;
    }

    // Smart delimiter detection
    const delimiter = detectDelimiter(lines[0]);
    
    headers = parseCSVLine(lines[0], delimiter);
    currentData = lines.slice(1).map(line => parseCSVLine(line, delimiter));

    // For CSV, store as single sheet
    allSheets = [{ name: 'Sheet1', headers: headers, data: currentData }];
    currentFile = file;
    currentSheetIndex = 0;
    
    // Hide sheet tabs for CSV (single sheet)
    sheetTabs.style.display = 'none';

    displayData(file);
}

// Detect CSV delimiter
function detectDelimiter(line) {
    const delimiters = [',', ';', '\t', '|'];
    let maxCount = 0;
    let bestDelimiter = ',';

    delimiters.forEach(d => {
        const count = (line.match(new RegExp(escapeRegex(d), 'g')) || []).length;
        if (count > maxCount) {
            maxCount = count;
            bestDelimiter = d;
        }
    });

    return bestDelimiter;
}

// Parse CSV line handling quoted values
function parseCSVLine(line, delimiter) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        
        if (char === '"') {
            if (inQuotes && line[i + 1] === '"') {
                current += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === delimiter && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current.trim());

    return result;
}

// Parse Excel file
function parseExcel(data, file) {
    try {
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Parse all sheets
        allSheets = workbook.SheetNames.map(sheetName => {
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

        // Check if all sheets are empty
        if (allSheets.every(sheet => sheet.data.length === 0 && sheet.headers.length === 0)) {
            alert('The file appears to be empty');
            clearData();
            return;
        }

        currentFile = file;
        currentSheetIndex = 0;
        loadSheet(0);
        displayData(file);
        
        // Show sheet tabs if multiple sheets
        if (allSheets.length > 1) {
            renderSheetTabs();
        }
    } catch (error) {
        alert('Error reading Excel file. Please try again.');
        clearData();
        console.error(error);
    }
}

// Load a specific sheet
function loadSheet(index) {
    if (index < 0 || index >= allSheets.length) return;
    
    currentSheetIndex = index;
    headers = [...allSheets[index].headers];
    // Load fresh data from sheet (includes any edits made)
    currentData = allSheets[index].data.map(row => [...row]);
    
    // Update active tab
    document.querySelectorAll('.sheet-tab').forEach((tab, i) => {
        tab.classList.toggle('active', i === index);
    });
    
    // Reset sort and filter state
    resetSort();
    activeFilters = [];
    updateActiveFiltersUI();
    
    // Update row count
    rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
    
    // Clear search and SQL, update placeholder and tables hint
    searchInput.value = '';
    sqlInput.value = '';
    sqlStatus.className = 'sql-status';
    hasSqlResults = false;
    sqlResultData = null;
    hideSqlResultBar();
    updateSqlTablesHint();
    updateSqlPlaceholder();
    
    renderTable(currentData);
    
    // Show edit indicator if there are unsaved changes
    document.getElementById('editIndicator').style.display = hasUnsavedChanges ? 'flex' : 'none';
}

// Render sheet tabs
function renderSheetTabs() {
    sheetTabs.innerHTML = '';
    sheetTabs.style.display = 'flex';
    
    allSheets.forEach((sheet, index) => {
        const tab = document.createElement('button');
        tab.className = 'sheet-tab' + (index === currentSheetIndex ? ' active' : '');
        tab.textContent = sheet.name;
        tab.addEventListener('click', () => loadSheet(index));
        sheetTabs.appendChild(tab);
    });
}

// Display data in table
function displayData(file) {
    // Set mode
    currentMode = 'spreadsheet';
    
    // Update file info
    const canSaveBack = fileHandle !== null;
    fileName.textContent = '📄 ' + file.name + (canSaveBack ? ' ✓' : '');
    fileName.title = canSaveBack ? 'File can be saved back to original' : 'File will be downloaded as new file when saving';
    fileSize.textContent = formatFileSize(file.size);
    rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
    
    // Show file info, SQL section, and table; hide upload section and document
    fileInfo.style.display = 'flex';
    docInfo.style.display = 'none';
    sqlSection.style.display = 'block';
    tableContainer.style.display = 'block';
    documentContainer.style.display = 'none';
    uploadSection.style.display = 'none';
    clearBtn.style.display = 'inline-block';
    
    // Reset edit state for new file
    hasUnsavedChanges = false;
    originalSheetData = [];
    autoSaveEnabled = false;
    resetSort();
    document.getElementById('autosaveCheckbox').checked = false;
    document.getElementById('editIndicator').style.display = 'none';
    document.getElementById('saveStatus').style.display = 'none';
    
    // Show/hide autosave toggle based on File System Access API support
    const autosaveToggle = document.getElementById('autosaveToggle');
    autosaveToggle.style.display = canSaveBack ? 'flex' : 'none';
    autosaveToggle.title = canSaveBack 
        ? 'Auto-save changes to original file' 
        : 'Open file using "Browse Files" button to enable auto-save';
    
    // Register SQL tables and update placeholder
    registerSqlTables();
    updateSqlPlaceholder();
    
    // Update filter columns if panel is open
    if (quickFilterPanel.style.display !== 'none') {
        populateFilterColumns();
    }

    renderTable(currentData);
}

// Update SQL placeholder with actual column names
function updateSqlPlaceholder() {
    if (headers.length > 0) {
        const cols = headers.slice(0, 3).map((h, i) => getColumnName(h, i)).join(', ');
        const firstCol = getColumnName(headers[0], 0);
        
        if (allSheets.length > 1) {
            const tables = allSheets.slice(0, 2).map(s => getTableName(s.name)).join(', ');
            sqlInput.placeholder = `SELECT * FROM data WHERE ${firstCol} = 'value' -- Tables: ${tables}...`;
        } else {
            sqlInput.placeholder = `SELECT ${cols}... FROM data WHERE ${firstCol} = 'value' ORDER BY ${firstCol}`;
        }
    }
}

// Render table with data
function renderTable(data, isFiltered = false) {
    if (data.length === 0 && headers.length === 0) {
        tableContainer.innerHTML = '<p class="placeholder-text">No data to display</p>';
        return;
    }

    // Check if this is the original unfiltered data (editable)
    // When isFiltered is false and no SQL results, the data is editable
    const isOriginalData = !isFiltered && !hasSqlResults;
    
    let html = '<div class="table-wrapper"><table class="data-table" id="dataTable">';
    
    // Header with sort indicators
    html += '<thead><tr><th class="row-num">#</th>';
    headers.forEach((header, colIndex) => {
        const isSorted = sortColumn === colIndex;
        const sortClass = isSorted ? `sortable sorted-${sortDirection}` : 'sortable';
        const sortIcon = isSorted 
            ? (sortDirection === 'asc' ? '↑' : '↓') 
            : '⇅';
        html += `<th data-col="${colIndex}" class="${sortClass}" title="Click to sort • Double-click to rename">${escapeHtml(header)}<span class="sort-icon">${sortIcon}</span></th>`;
    });
    html += '</tr></thead>';

    // Body
    html += '<tbody>';
    if (data.length === 0) {
        // No rows yet - show helpful message
        html += `<tr class="empty-row-message">
            <td colspan="${headers.length + 1}" class="empty-table-cell">
                <span>No rows yet.</span>
                <button class="inline-add-row-btn" id="inlineAddRowBtn">➕ Add Row</button>
            </td>
        </tr>`;
    } else {
        data.forEach((row, displayIndex) => {
            // Find original row index in the sheet data
            const originalRowIndex = isOriginalData ? displayIndex : findOriginalRowIndex(row);
            const canEdit = isOriginalData && originalRowIndex !== -1;
            
            html += `<tr data-row="${displayIndex}" data-original-row="${originalRowIndex}">`;
            html += `<td class="row-num">
                <span class="row-number">${displayIndex + 1}</span>
                ${canEdit ? `<button class="delete-row-btn" data-original-row="${originalRowIndex}" title="Delete row">🗑️</button>` : ''}
            </td>`;
            headers.forEach((_, colIndex) => {
                const value = row[colIndex] || '';
                const editAttr = canEdit ? `data-original-row="${originalRowIndex}"` : '';
                const editTitle = canEdit ? 'Double-click to edit' : 'Cannot edit filtered/SQL results';
                const editClass = canEdit ? 'editable' : 'readonly';
                html += `<td data-row="${displayIndex}" data-col="${colIndex}" ${editAttr} class="${editClass}" title="${editTitle}">${escapeHtml(value)}</td>`;
            });
            html += '</tr>';
        });
    }
    html += '</tbody></table></div>';

    tableContainer.innerHTML = html;
    
    // Bind inline add row button if present
    const inlineAddRowBtn = document.getElementById('inlineAddRowBtn');
    if (inlineAddRowBtn) {
        inlineAddRowBtn.addEventListener('click', addNewRow);
    }
}

// Find original row index by matching row data
function findOriginalRowIndex(row) {
    const sheetData = allSheets[currentSheetIndex]?.data;
    if (!sheetData) return -1;
    
    for (let i = 0; i < sheetData.length; i++) {
        if (arraysEqual(sheetData[i], row)) {
            return i;
        }
    }
    return -1;
}

// Check if two arrays are equal
function arraysEqual(a, b) {
    if (!a || !b) return false;
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) {
        if (a[i] !== b[i]) return false;
    }
    return true;
}

// Handle header click for sorting
function handleHeaderClick(e) {
    const th = e.target.closest('th.sortable');
    if (!th) return;
    
    // Don't sort if we're in edit mode
    if (th.classList.contains('editing')) return;
    
    const colIndex = parseInt(th.dataset.col);
    if (isNaN(colIndex)) return;
    
    // Toggle sort direction
    if (sortColumn === colIndex) {
        // Cycle through: asc -> desc -> none
        if (sortDirection === 'asc') {
            sortDirection = 'desc';
        } else if (sortDirection === 'desc') {
            sortDirection = 'none';
            sortColumn = -1;
        }
    } else {
        sortColumn = colIndex;
        sortDirection = 'asc';
    }
    
    // Sort and re-render
    sortAndRenderTable();
}

// Handle header double-click for renaming
function handleHeaderDoubleClick(e) {
    const th = e.target.closest('th.sortable');
    if (!th || th.classList.contains('editing')) return;
    
    if (hasSqlResults) {
        showToast('Cannot rename columns in SQL results. Reset first.', 'warning');
        return;
    }
    
    const colIndex = parseInt(th.dataset.col);
    if (isNaN(colIndex)) return;
    
    e.stopPropagation();
    startHeaderEdit(th, colIndex);
}

// Start editing a header
function startHeaderEdit(th, colIndex) {
    const currentValue = headers[colIndex] || '';
    
    // Store original data for undo if not already stored
    if (!hasUnsavedChanges) {
        storeOriginalData();
    }
    
    th.classList.add('editing');
    
    // Remove sort icon temporarily
    const sortIcon = th.querySelector('.sort-icon');
    const sortIconHtml = sortIcon ? sortIcon.outerHTML : '';
    
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'header-editor';
    input.value = currentValue;
    
    th.innerHTML = '';
    th.appendChild(input);
    input.focus();
    input.select();
    
    // Handle blur (save)
    input.addEventListener('blur', () => {
        finishHeaderEdit(th, input, colIndex, sortIconHtml);
    });
    
    // Handle keyboard
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            input.blur();
        } else if (e.key === 'Escape') {
            // Cancel edit
            th.classList.remove('editing');
            th.innerHTML = escapeHtml(currentValue) + sortIconHtml;
        } else if (e.key === 'Tab') {
            e.preventDefault();
            input.blur();
            // Move to next/prev column header
            const nextColIndex = e.shiftKey ? colIndex - 1 : colIndex + 1;
            if (nextColIndex >= 0 && nextColIndex < headers.length) {
                const nextTh = document.querySelector(`th[data-col="${nextColIndex}"]`);
                if (nextTh) {
                    setTimeout(() => startHeaderEdit(nextTh, nextColIndex), 50);
                }
            }
        }
    });
    
    // Prevent click from triggering sort
    input.addEventListener('click', (e) => {
        e.stopPropagation();
    });
}

// Finish editing a header
function finishHeaderEdit(th, input, colIndex, sortIconHtml) {
    const newValue = input.value.trim();
    const oldValue = headers[colIndex] || '';
    
    // Use default name if empty
    const finalValue = newValue || `Column ${String.fromCharCode(65 + colIndex)}`;
    
    // Update headers
    headers[colIndex] = finalValue;
    allSheets[currentSheetIndex].headers[colIndex] = finalValue;
    
    // Update cell display
    th.classList.remove('editing');
    
    // Determine sort state for this column
    const isSorted = sortColumn === colIndex;
    const sortClass = isSorted ? `sortable sorted-${sortDirection}` : 'sortable';
    const sortIcon = isSorted 
        ? (sortDirection === 'asc' ? '↑' : '↓') 
        : '⇅';
    
    th.innerHTML = escapeHtml(finalValue) + `<span class="sort-icon">${sortIcon}</span>`;
    
    // Mark as changed if value changed
    if (finalValue !== oldValue) {
        markAsEdited();
        sqlTablesRegistered = false;
        
        // Update SQL placeholder with new column names
        updateSqlPlaceholder();
    }
}

// Sort data and render table
function sortAndRenderTable() {
    // Get data to sort based on current state
    let dataToSort;
    
    if (hasSqlResults) {
        // Sort SQL results
        dataToSort = sqlResultData.data.map(row => [...row]);
    } else if (activeFilters.length > 0) {
        // Sort already filtered data - reapply filters first
        dataToSort = allSheets[currentSheetIndex].data.map(row => [...row]);
        activeFilters.forEach(filter => {
            dataToSort = dataToSort.filter(row => {
                const cellValue = (row[filter.columnIndex] || '').toString();
                return evaluateFilter(cellValue, filter.operator, filter.value);
            });
        });
    } else {
        // Sort current sheet data
        dataToSort = allSheets[currentSheetIndex].data.map(row => [...row]);
    }
    
    if (sortColumn >= 0 && sortDirection !== 'none') {
        dataToSort.sort((a, b) => {
            let valA = a[sortColumn] || '';
            let valB = b[sortColumn] || '';
            
            // Try numeric comparison first
            const numA = parseFloat(valA);
            const numB = parseFloat(valB);
            
            if (!isNaN(numA) && !isNaN(numB)) {
                return sortDirection === 'asc' ? numA - numB : numB - numA;
            }
            
            // String comparison (case-insensitive)
            valA = valA.toString().toLowerCase();
            valB = valB.toString().toLowerCase();
            
            if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    }
    
    currentData = dataToSort;
    const isFiltered = hasSqlResults || activeFilters.length > 0;
    renderTable(currentData, isFiltered);
    
    // Update row count
    const total = allSheets[currentSheetIndex]?.data.length || currentData.length;
    let statusText = '';
    
    if (hasSqlResults) {
        statusText = `${currentData.length} rows × ${headers.length} columns (SQL result)`;
    } else if (activeFilters.length > 0) {
        statusText = `Showing ${currentData.length} of ${total} rows (filtered)`;
    } else {
        statusText = `${currentData.length} rows × ${headers.length} columns`;
    }
    
    const sortInfo = sortColumn >= 0 && sortDirection !== 'none' 
        ? ` • Sorted by ${headers[sortColumn]} ${sortDirection === 'asc' ? '↑' : '↓'}` 
        : '';
    rowCount.textContent = statusText + sortInfo;
}

// Reset sort state
function resetSort() {
    sortColumn = -1;
    sortDirection = 'none';
}

// Toggle filter panel
function toggleFilterPanel() {
    const isVisible = quickFilterPanel.style.display !== 'none';
    quickFilterPanel.style.display = isVisible ? 'none' : 'block';
    filterToggleBtn.classList.toggle('active', !isVisible);
    
    if (!isVisible) {
        populateFilterColumns();
    }
}

// Populate filter column dropdown
function populateFilterColumns() {
    filterColumn.innerHTML = '<option value="">Select Column...</option>';
    headers.forEach((header, index) => {
        const cleanName = header || `Column ${index + 1}`;
        filterColumn.innerHTML += `<option value="${index}">${cleanName}</option>`;
    });
}

// Update filter value visibility based on operator
function updateFilterValueVisibility() {
    const operator = filterOperator.value;
    const valueGroup = document.querySelector('.filter-value-group');
    
    if (operator === 'is_empty' || operator === 'is_not_empty') {
        valueGroup.style.display = 'none';
    } else {
        valueGroup.style.display = 'block';
    }
}

// Apply quick filter
function applyQuickFilter() {
    const colIndex = parseInt(filterColumn.value);
    const operator = filterOperator.value;
    const value = filterValue.value.trim();
    
    if (isNaN(colIndex)) {
        showToast('Please select a column', 'warning');
        return;
    }
    
    if (!['is_empty', 'is_not_empty'].includes(operator) && !value) {
        showToast('Please enter a filter value', 'warning');
        return;
    }
    
    // Create filter object
    const filter = {
        id: Date.now(),
        columnIndex: colIndex,
        columnName: headers[colIndex] || `Column ${colIndex + 1}`,
        operator: operator,
        value: value,
        operatorLabel: getOperatorLabel(operator)
    };
    
    // Add to active filters
    activeFilters.push(filter);
    
    // Apply filters and re-render
    applyAllFilters();
    
    // Update UI
    updateActiveFiltersUI();
    
    // Clear input
    filterValue.value = '';
    filterColumn.value = '';
    
    showToast('Filter applied', 'success');
}

// Get operator label
function getOperatorLabel(operator) {
    const labels = {
        'equals': '=',
        'not_equals': '≠',
        'contains': 'contains',
        'not_contains': 'not contains',
        'starts_with': 'starts with',
        'ends_with': 'ends with',
        'greater_than': '>',
        'less_than': '<',
        'greater_equal': '≥',
        'less_equal': '≤',
        'is_empty': 'is empty',
        'is_not_empty': 'is not empty'
    };
    return labels[operator] || operator;
}

// Apply all active filters
function applyAllFilters() {
    // Start with original sheet data
    let filteredData = allSheets[currentSheetIndex].data.map(row => [...row]);
    
    // Apply each filter
    activeFilters.forEach(filter => {
        filteredData = filteredData.filter(row => {
            const cellValue = (row[filter.columnIndex] || '').toString();
            return evaluateFilter(cellValue, filter.operator, filter.value);
        });
    });
    
    currentData = filteredData;
    hasSqlResults = false;
    hideSqlResultBar();
    
    // Apply sort if active
    if (sortColumn >= 0 && sortDirection !== 'none') {
        sortAndRenderTable();
    } else {
        renderTable(currentData, activeFilters.length > 0);
    }
    
    // Update row count
    const total = allSheets[currentSheetIndex].data.length;
    if (activeFilters.length > 0) {
        rowCount.textContent = `Showing ${currentData.length} of ${total} rows (filtered)`;
    } else {
        rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
    }
}

// Evaluate filter condition
function evaluateFilter(cellValue, operator, filterValue) {
    const cellLower = cellValue.toLowerCase();
    const filterLower = filterValue.toLowerCase();
    const cellNum = parseFloat(cellValue);
    const filterNum = parseFloat(filterValue);
    
    switch (operator) {
        case 'equals':
            return cellLower === filterLower;
        case 'not_equals':
            return cellLower !== filterLower;
        case 'contains':
            return cellLower.includes(filterLower);
        case 'not_contains':
            return !cellLower.includes(filterLower);
        case 'starts_with':
            return cellLower.startsWith(filterLower);
        case 'ends_with':
            return cellLower.endsWith(filterLower);
        case 'greater_than':
            return !isNaN(cellNum) && !isNaN(filterNum) && cellNum > filterNum;
        case 'less_than':
            return !isNaN(cellNum) && !isNaN(filterNum) && cellNum < filterNum;
        case 'greater_equal':
            return !isNaN(cellNum) && !isNaN(filterNum) && cellNum >= filterNum;
        case 'less_equal':
            return !isNaN(cellNum) && !isNaN(filterNum) && cellNum <= filterNum;
        case 'is_empty':
            return cellValue.trim() === '';
        case 'is_not_empty':
            return cellValue.trim() !== '';
        default:
            return true;
    }
}

// Update active filters UI
function updateActiveFiltersUI() {
    const activeFiltersDiv = document.getElementById('activeFilters');
    const filterTags = document.getElementById('filterTags');
    
    if (activeFilters.length === 0) {
        activeFiltersDiv.style.display = 'none';
        return;
    }
    
    activeFiltersDiv.style.display = 'flex';
    filterTags.innerHTML = activeFilters.map(filter => {
        const valueDisplay = ['is_empty', 'is_not_empty'].includes(filter.operator) 
            ? '' 
            : ` "${filter.value}"`;
        return `
            <div class="filter-tag">
                <span class="filter-tag-text">${filter.columnName} ${filter.operatorLabel}${valueDisplay}</span>
                <button class="filter-tag-remove" data-filter-id="${filter.id}" title="Remove filter">✕</button>
            </div>
        `;
    }).join('');
    
    // Add remove event listeners
    filterTags.querySelectorAll('.filter-tag-remove').forEach(btn => {
        btn.addEventListener('click', () => {
            removeFilter(parseInt(btn.dataset.filterId));
        });
    });
}

// Remove a specific filter
function removeFilter(filterId) {
    activeFilters = activeFilters.filter(f => f.id !== filterId);
    applyAllFilters();
    updateActiveFiltersUI();
    showToast('Filter removed', 'info');
}

// Clear all filters
function clearAllFilters() {
    activeFilters = [];
    filterColumn.value = '';
    filterValue.value = '';
    filterOperator.value = 'equals';
    
    // Reload original data
    currentData = allSheets[currentSheetIndex].data.map(row => [...row]);
    
    // Apply sort if active
    if (sortColumn >= 0 && sortDirection !== 'none') {
        sortAndRenderTable();
    } else {
        renderTable(currentData);
    }
    
    updateActiveFiltersUI();
    rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
    
    showToast('All filters cleared', 'info');
}

// Handle single click on cell (select)
function handleCellClick(e) {
    const cell = e.target.closest('td:not(.row-num)');
    if (!cell || cell.classList.contains('editing')) return;
    
    // Visual feedback
    document.querySelectorAll('.data-table td.selected').forEach(td => {
        td.classList.remove('selected');
    });
    cell.classList.add('selected');
}

// Handle double click on cell (edit)
function handleCellDoubleClick(e) {
    const cell = e.target.closest('td:not(.row-num)');
    if (!cell || cell.classList.contains('editing')) return;
    
    // Check if cell is editable
    if (!cell.classList.contains('editable')) {
        showToast('Cannot edit filtered or SQL results. Clear search/reset to edit.', 'warning');
        return;
    }
    
    startCellEdit(cell);
}

// Start editing a cell
function startCellEdit(cell) {
    const displayRowIndex = parseInt(cell.dataset.row);
    const originalRowIndex = parseInt(cell.dataset.originalRow);
    const colIndex = parseInt(cell.dataset.col);
    
    // Get value from the original sheet data
    const currentValue = allSheets[currentSheetIndex]?.data[originalRowIndex]?.[colIndex] || '';
    
    // Store original data for undo if not already stored
    if (!hasUnsavedChanges) {
        storeOriginalData();
    }
    
    cell.classList.add('editing');
    
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'cell-editor';
    input.value = currentValue;
    
    cell.innerHTML = '';
    cell.appendChild(input);
    input.focus();
    input.select();
    
    // Handle blur (save)
    input.addEventListener('blur', () => {
        finishCellEdit(cell, input, originalRowIndex, colIndex, displayRowIndex);
    });
    
    // Handle keyboard
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            input.blur();
            // Move to next row
            const nextRow = cell.parentElement.nextElementSibling;
            if (nextRow) {
                const nextCell = nextRow.querySelector(`td[data-col="${colIndex}"].editable`);
                if (nextCell) {
                    setTimeout(() => startCellEdit(nextCell), 50);
                }
            }
        } else if (e.key === 'Tab') {
            e.preventDefault();
            input.blur();
            // Move to next column
            const nextColIndex = e.shiftKey ? colIndex - 1 : colIndex + 1;
            const nextCell = cell.parentElement.querySelector(`td[data-col="${nextColIndex}"].editable`);
            if (nextCell) {
                setTimeout(() => startCellEdit(nextCell), 50);
            }
        } else if (e.key === 'Escape') {
            // Cancel edit
            cell.classList.remove('editing');
            cell.innerHTML = escapeHtml(currentValue);
            cell.title = 'Double-click to edit';
        }
    });
}

// Finish editing a cell
function finishCellEdit(cell, input, originalRowIndex, colIndex, displayRowIndex) {
    const newValue = input.value;
    const oldValue = allSheets[currentSheetIndex]?.data[originalRowIndex]?.[colIndex] || '';
    
    // Update the original sheet data (this is the source of truth)
    if (allSheets[currentSheetIndex]?.data[originalRowIndex]) {
        allSheets[currentSheetIndex].data[originalRowIndex][colIndex] = newValue;
    }
    
    // Also update currentData for display consistency
    if (currentData[displayRowIndex]) {
        currentData[displayRowIndex][colIndex] = newValue;
    }
    
    // Update cell display
    cell.classList.remove('editing');
    cell.innerHTML = escapeHtml(newValue);
    cell.title = 'Double-click to edit';
    
    // Mark as changed if value changed
    if (newValue !== oldValue) {
        markAsEdited();
        cell.classList.add('cell-changed');
        
        // Re-register SQL tables with new data
        sqlTablesRegistered = false;
    }
}

// Store original data for undo
function storeOriginalData() {
    originalSheetData = allSheets.map(sheet => ({
        name: sheet.name,
        headers: [...sheet.headers],
        data: sheet.data.map(row => [...row])
    }));
}

// Mark document as edited
function markAsEdited() {
    hasUnsavedChanges = true;
    document.getElementById('editIndicator').style.display = 'flex';
    
    // Trigger auto-save if enabled
    if (autoSaveEnabled) {
        triggerAutoSave();
    }
}

// Undo all changes
function undoAllChanges() {
    if (!hasUnsavedChanges || originalSheetData.length === 0) return;
    
    if (!confirm('Undo all changes and restore original data?')) return;
    
    // Restore original data
    allSheets = originalSheetData.map(sheet => ({
        name: sheet.name,
        headers: [...sheet.headers],
        data: sheet.data.map(row => [...row])
    }));
    
    // Reload current sheet
    headers = [...allSheets[currentSheetIndex].headers];
    currentData = allSheets[currentSheetIndex].data.map(row => [...row]);
    
    // Re-render
    renderTable(currentData);
    rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
    
    // Reset edit state
    hasUnsavedChanges = false;
    document.getElementById('editIndicator').style.display = 'none';
    sqlTablesRegistered = false;
    
    showToast('All changes undone', 'success');
}

// Delete a row
function deleteRow(originalRowIndex) {
    if (hasSqlResults) {
        showToast('Cannot delete rows from SQL results. Reset first.', 'warning');
        return;
    }
    
    const sheetData = allSheets[currentSheetIndex]?.data;
    if (!sheetData || originalRowIndex < 0 || originalRowIndex >= sheetData.length) return;
    
    // Store original data for undo if not already stored
    if (!hasUnsavedChanges) {
        storeOriginalData();
    }
    
    // Remove from original sheet data
    allSheets[currentSheetIndex].data.splice(originalRowIndex, 1);
    
    // Reload current data from sheet
    currentData = allSheets[currentSheetIndex].data.map(row => [...row]);
    
    // Re-render
    renderTable(currentData);
    rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
    
    // Mark as edited
    markAsEdited();
    sqlTablesRegistered = false;
    
    showToast('Row deleted', 'info');
}

// Add new row
function addNewRow() {
    if (hasSqlResults) {
        showToast('Cannot add rows to SQL results. Reset first.', 'warning');
        return;
    }
    
    if (allSheets.length === 0) return;
    
    // Check if there are columns first
    if (headers.length === 0) {
        showToast('Add columns first before adding rows', 'warning');
        return;
    }
    
    // Clear search first
    if (searchInput.value.trim()) {
        searchInput.value = '';
    }
    
    // Store original data for undo if not already stored
    if (!hasUnsavedChanges) {
        storeOriginalData();
    }
    
    // Create empty row
    const newRow = headers.map(() => '');
    
    // Add to data (currentData and allSheets[].data should be same reference for new sheets)
    currentData.push(newRow);
    allSheets[currentSheetIndex].data = currentData;
    
    // Re-render
    renderTable(currentData);
    updateRowColCount();
    
    // Mark as edited
    markAsEdited();
    
    // Scroll to bottom and start editing first cell of new row
    const tableWrapper = document.querySelector('.table-wrapper');
    if (tableWrapper) {
        tableWrapper.scrollTop = tableWrapper.scrollHeight;
    }
    
    // Start editing first cell of new row
    setTimeout(() => {
        const lastRow = document.querySelector(`tr[data-row="${currentData.length - 1}"]`);
        if (lastRow) {
            const firstCell = lastRow.querySelector('td[data-col="0"].editable');
            if (firstCell) {
                startCellEdit(firstCell);
            }
        }
    }, 100);
    
    sqlTablesRegistered = false;
    showToast('New row added', 'success');
}

// Show toast notification
function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast-notification ${type}`;
    toast.innerHTML = message;
    document.body.appendChild(toast);
    
    setTimeout(() => toast.classList.add('show'), 10);
    
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => document.body.removeChild(toast), 300);
    }, 3000);
}

// Save file to disk
async function saveFile() {
    if (!currentFile && !fileHandle) {
        showToast('No file to save', 'error');
        return;
    }
    
    showSaveStatus('saving', 'Saving...');
    
    try {
        // Create workbook from current data
        const wb = XLSX.utils.book_new();
        
        allSheets.forEach(sheet => {
            const sheetData = [sheet.headers, ...sheet.data];
            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(wb, ws, sheet.name);
        });
        
        // Determine file type from current file
        const fileName = currentFile?.name || 'data.xlsx';
        const ext = fileName.split('.').pop().toLowerCase();
        
        if (fileHandle) {
            // Use File System Access API to save back to original file
            await saveWithFileHandle(wb, ext);
        } else {
            // Download as new file
            await saveAsDownload(wb, fileName, ext);
        }
        
        // Mark as saved
        hasUnsavedChanges = false;
        document.getElementById('editIndicator').style.display = 'none';
        showSaveStatus('saved', 'Saved');
        lastSaveTime = new Date();
        
        // Hide save status after 2 seconds
        setTimeout(() => {
            document.getElementById('saveStatus').style.display = 'none';
        }, 2000);
        
    } catch (err) {
        console.error('Save error:', err);
        showSaveStatus('error', 'Save failed');
        showToast('Failed to save file: ' + err.message, 'error');
    }
}

// Save using File System Access API (write back to original file)
async function saveWithFileHandle(wb, ext) {
    const writable = await fileHandle.createWritable();
    
    let content;
    let bookType;
    
    if (ext === 'csv') {
        // For CSV, only save current sheet
        content = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
        await writable.write(content);
    } else {
        bookType = ext === 'xls' ? 'biff8' : 'xlsx';
        const buffer = XLSX.write(wb, { bookType: bookType, type: 'array' });
        await writable.write(buffer);
    }
    
    await writable.close();
}

// Save as download (fallback)
async function saveAsDownload(wb, fileName, ext) {
    if (ext === 'csv') {
        const csv = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
        downloadBlob(csv, fileName, 'text/csv;charset=utf-8;');
    } else {
        const bookType = ext === 'xls' ? 'biff8' : 'xlsx';
        XLSX.writeFile(wb, fileName, { bookType: bookType });
    }
}

// Show save status
function showSaveStatus(status, text) {
    const saveStatus = document.getElementById('saveStatus');
    const saveText = document.getElementById('saveText');
    
    saveStatus.className = 'save-status ' + status;
    saveText.textContent = text;
    saveStatus.style.display = 'flex';
}

// Trigger auto-save with debounce
function triggerAutoSave() {
    if (!autoSaveEnabled || !hasUnsavedChanges) return;
    
    // Clear existing timeout
    if (autoSaveTimeout) {
        clearTimeout(autoSaveTimeout);
    }
    
    // Debounce: wait 1 second after last change before saving
    autoSaveTimeout = setTimeout(async () => {
        if (autoSaveEnabled && hasUnsavedChanges) {
            await saveFile();
        }
    }, 1000);
}

// Search functionality
function handleSearch() {
    const query = searchInput.value.toLowerCase().trim();
    
    // Start with original sheet data
    let dataToSearch = allSheets[currentSheetIndex].data.map(row => [...row]);
    
    // Apply active quick filters first
    if (activeFilters.length > 0) {
        activeFilters.forEach(filter => {
            dataToSearch = dataToSearch.filter(row => {
                const cellValue = (row[filter.columnIndex] || '').toString();
                return evaluateFilter(cellValue, filter.operator, filter.value);
            });
        });
    }
    
    // Reset sort when searching
    resetSort();
    
    if (!query) {
        currentData = dataToSearch;
        renderTable(currentData, activeFilters.length > 0);
        if (activeFilters.length > 0) {
            rowCount.textContent = `Showing ${currentData.length} of ${allSheets[currentSheetIndex].data.length} rows (filtered)`;
        } else {
            rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
        }
        return;
    }

    // Apply text search
    const searchResults = dataToSearch.filter(row => 
        row.some(cell => cell.toLowerCase().includes(query))
    );

    currentData = searchResults;
    
    renderTable(searchResults, true);  // Pass isFiltered = true
    
    // Update row count for search results
    const baseCount = activeFilters.length > 0 ? dataToSearch.length : allSheets[currentSheetIndex].data.length;
    rowCount.textContent = `Showing ${searchResults.length} of ${baseCount} rows (search${activeFilters.length > 0 ? ' + filter' : ''})`;
}

// Toggle SQL examples panel
function toggleExamplesPanel() {
    const isVisible = sqlExamplesPanel.style.display !== 'none';
    sqlExamplesPanel.style.display = isVisible ? 'none' : 'block';
    
    if (!isVisible) {
        updateExamplesPanel();
    }
}

// Get clean table name from sheet name
function getTableName(sheetName) {
    // Clean sheet name for SQL (remove spaces, special chars, ensure valid identifier)
    let name = sheetName.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^[0-9]/, '_$&');
    return name || 'sheet';
}

// Get clean column name
function getColumnName(header, index) {
    if (!header) return `col_${index}`;
    return header.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^[0-9]/, '_$&') || `col_${index}`;
}

// Register all sheets as SQL tables
function registerSqlTables() {
    // Clear existing tables
    allSheets.forEach(sheet => {
        const tableName = getTableName(sheet.name);
        try {
            alasql(`DROP TABLE IF EXISTS ${tableName}`);
        } catch (e) {}
    });
    try {
        alasql('DROP TABLE IF EXISTS data');
    } catch (e) {}
    
    // Register each sheet as a table
    allSheets.forEach((sheet, index) => {
        const tableName = getTableName(sheet.name);
        
        // Create array of objects from data
        const tableData = sheet.data.map(row => {
            const obj = {};
            sheet.headers.forEach((h, i) => {
                const cleanHeader = getColumnName(h, i);
                // Try to convert to number if possible
                let value = row[i] || '';
                if (value !== '' && !isNaN(value) && !isNaN(parseFloat(value))) {
                    value = parseFloat(value);
                }
                obj[cleanHeader] = value;
            });
            return obj;
        });
        
        // Register the table
        alasql(`CREATE TABLE ${tableName}`);
        alasql.tables[tableName].data = tableData;
        
        // Also register current sheet as "data" for convenience
        if (index === currentSheetIndex) {
            alasql('CREATE TABLE data');
            alasql.tables.data.data = [...tableData];
        }
    });
    
    sqlTablesRegistered = true;
    updateSqlTablesHint();
}

// Update SQL tables hint in header
function updateSqlTablesHint() {
    const hint = document.getElementById('sqlTablesHint');
    if (allSheets.length === 1) {
        const tableName = getTableName(allSheets[0].name);
        hint.innerHTML = `Tables: <code>data</code> <span class="table-separator">or</span> <code>${tableName}</code>`;
    } else {
        const tableNames = allSheets.map(s => `<code>${getTableName(s.name)}</code>`).join(' ');
        hint.innerHTML = `Tables: <code>data</code> (current) ${tableNames}`;
    }
}

// Update examples panel with dynamic table names
function updateExamplesPanel() {
    const tablesListPanel = document.getElementById('tablesListPanel');
    const basicExamples = document.getElementById('basicExamples');
    const joinExamples = document.getElementById('joinExamples');
    const joinExamplesSection = document.getElementById('joinExamplesSection');
    
    // Build tables list
    tablesListPanel.innerHTML = allSheets.map((sheet, index) => {
        const tableName = getTableName(sheet.name);
        const isActive = index === currentSheetIndex;
        const cols = sheet.headers.slice(0, 3).map((h, i) => getColumnName(h, i)).join(', ');
        return `
            <div class="table-badge ${isActive ? 'active' : ''}" data-table="${tableName}">
                <code>${tableName}</code>
                <span class="table-info">${sheet.data.length} rows • ${sheet.headers.length} cols</span>
            </div>
        `;
    }).join('');
    
    // Current table for examples
    const currentTable = getTableName(allSheets[currentSheetIndex].name);
    const currentCols = allSheets[currentSheetIndex].headers;
    const col1 = currentCols[0] ? getColumnName(currentCols[0], 0) : 'column1';
    const col2 = currentCols[1] ? getColumnName(currentCols[1], 1) : 'column2';
    
    // Basic examples
    basicExamples.innerHTML = `
        <div class="example-item" data-query="SELECT * FROM data">
            <strong>Select All (Current Sheet)</strong>
            <code>SELECT * FROM data</code>
        </div>
        <div class="example-item" data-query="SELECT * FROM data LIMIT 10">
            <strong>Limit Rows</strong>
            <code>SELECT * FROM data LIMIT 10</code>
        </div>
        <div class="example-item" data-query="SELECT * FROM data WHERE ${col1} = 'value'">
            <strong>Filter Rows</strong>
            <code>SELECT * FROM data WHERE ${col1} = 'value'</code>
        </div>
        <div class="example-item" data-query="SELECT * FROM data ORDER BY ${col1} DESC">
            <strong>Sort Data</strong>
            <code>SELECT * FROM data ORDER BY ${col1} DESC</code>
        </div>
        <div class="example-item example-modify" data-query="UPDATE data SET ${col1} = 'new_value' WHERE ${col2} = 'condition'">
            <strong>✏️ Update Rows</strong>
            <code>UPDATE data SET ${col1} = 'new' WHERE ...</code>
        </div>
        <div class="example-item example-modify" data-query="INSERT INTO data (${col1}, ${col2}) VALUES ('value1', 'value2')">
            <strong>➕ Insert Row</strong>
            <code>INSERT INTO data VALUES (...)</code>
        </div>
        <div class="example-item example-modify" data-query="DELETE FROM data WHERE ${col1} = 'value'">
            <strong>🗑️ Delete Rows</strong>
            <code>DELETE FROM data WHERE ...</code>
        </div>
        <div class="example-item" data-query="SELECT ${col1}, COUNT(*) as count FROM data GROUP BY ${col1}">
            <strong>Group & Count</strong>
            <code>SELECT ${col1}, COUNT(*) FROM data GROUP BY ...</code>
        </div>
    `;
    
    // JOIN examples (only if multiple sheets)
    if (allSheets.length > 1) {
        joinExamplesSection.style.display = 'block';
        
        const table1 = getTableName(allSheets[0].name);
        const table2 = getTableName(allSheets[1].name);
        const t1Col = allSheets[0].headers[0] ? getColumnName(allSheets[0].headers[0], 0) : 'id';
        const t2Col = allSheets[1].headers[0] ? getColumnName(allSheets[1].headers[0], 0) : 'id';
        
        joinExamples.innerHTML = `
            <div class="example-item" data-query="SELECT * FROM ${table1} a JOIN ${table2} b ON a.${t1Col} = b.${t2Col}">
                <strong>Inner Join</strong>
                <code>SELECT * FROM ${table1} a JOIN ${table2} b ON a.${t1Col} = b.${t2Col}</code>
            </div>
            <div class="example-item" data-query="SELECT * FROM ${table1} a LEFT JOIN ${table2} b ON a.${t1Col} = b.${t2Col}">
                <strong>Left Join</strong>
                <code>SELECT * FROM ${table1} a LEFT JOIN ${table2} b ON a.${t1Col} = b.${t2Col}</code>
            </div>
            <div class="example-item" data-query="SELECT a.*, b.* FROM ${table1} a, ${table2} b WHERE a.${t1Col} = b.${t2Col}">
                <strong>Cross Join with Filter</strong>
                <code>SELECT a.*, b.* FROM ${table1} a, ${table2} b WHERE a.${t1Col} = b.${t2Col}</code>
            </div>
            <div class="example-item" data-query="SELECT * FROM ${table1} UNION SELECT * FROM ${table2}">
                <strong>Union (Combine Rows)</strong>
                <code>SELECT * FROM ${table1} UNION SELECT * FROM ${table2}</code>
            </div>
        `;
    } else {
        joinExamplesSection.style.display = 'none';
    }
}

// Convert data to SQL-friendly format and run query
function runSqlQuery() {
    const query = sqlInput.value.trim();
    
    if (!query) {
        showSqlStatus('Please enter a SQL query', 'error');
        return;
    }
    
    try {
        // Detect query type
        const queryType = detectQueryType(query);
        
        // Ensure tables are registered
        if (!sqlTablesRegistered) {
            registerSqlTables();
        }
        
        // Re-register "data" table to point to current sheet
        const currentSheet = allSheets[currentSheetIndex];
        const currentTableData = currentSheet.data.map(row => {
            const obj = {};
            currentSheet.headers.forEach((h, i) => {
                const cleanHeader = getColumnName(h, i);
                let value = row[i] || '';
                if (value !== '' && !isNaN(value) && !isNaN(parseFloat(value))) {
                    value = parseFloat(value);
                }
                obj[cleanHeader] = value;
            });
            return obj;
        });
        
        try { alasql('DROP TABLE IF EXISTS data'); } catch(e) {}
        alasql('CREATE TABLE data');
        alasql.tables.data.data = currentTableData;
        
        // Execute the query
        const startTime = performance.now();
        const result = alasql(query);
        const endTime = performance.now();
        const execTime = (endTime - startTime).toFixed(2);
        
        // Handle based on query type
        if (queryType === 'SELECT') {
            handleSelectResult(result, query, execTime, currentSheet);
        } else if (queryType === 'UPDATE' || queryType === 'DELETE' || queryType === 'INSERT') {
            handleModifyResult(result, query, execTime, queryType);
        } else {
            hasSqlResults = false;
            hideSqlResultBar();
            showSqlStatus(`✓ Query executed in ${execTime}ms — Result: ${JSON.stringify(result)}`, 'success');
        }
        
    } catch (error) {
        showSqlStatus(`✗ Error: ${error.message}`, 'error');
        console.error('SQL Error:', error);
    }
}

// Detect SQL query type
function detectQueryType(query) {
    const upperQuery = query.trim().toUpperCase();
    if (upperQuery.startsWith('SELECT')) return 'SELECT';
    if (upperQuery.startsWith('UPDATE')) return 'UPDATE';
    if (upperQuery.startsWith('DELETE')) return 'DELETE';
    if (upperQuery.startsWith('INSERT')) return 'INSERT';
    return 'OTHER';
}

// Handle SELECT query result
function handleSelectResult(result, query, execTime, currentSheet) {
    if (Array.isArray(result)) {
        // Update headers and data from result
        if (result.length > 0) {
            headers = Object.keys(result[0]);
            currentData = result.map(row => headers.map(h => {
                const val = row[h];
                return val !== null && val !== undefined ? val.toString() : '';
            }));
        } else {
            headers = currentSheet.headers.map((h, i) => getColumnName(h, i));
            currentData = [];
        }
        
        // Store SQL results for export
        sqlResultData = {
            headers: [...headers],
            data: currentData.map(row => [...row]),
            query: query,
            rowCount: result.length
        };
        hasSqlResults = true;
        
        // Show result bar with export option
        showSqlResultBar(result.length, headers.length, execTime);
        
        renderTable(currentData, true);
        showSqlStatus(`✓ Query executed in ${execTime}ms — ${result.length} rows returned`, 'success');
        rowCount.textContent = `${result.length} rows × ${headers.length} columns (SQL result)`;
    }
}

// Handle UPDATE, DELETE, INSERT query result
function handleModifyResult(result, query, execTime, queryType) {
    // Store original data for undo if not already stored
    if (!hasUnsavedChanges) {
        storeOriginalData();
    }
    
    // Determine which table was modified
    const tableName = extractTableName(query, queryType);
    const sheetIndex = findSheetIndexByTableName(tableName);
    
    if (sheetIndex === -1) {
        showSqlStatus(`✗ Error: Table "${tableName}" not found`, 'error');
        return;
    }
    
    // Get the modified data from AlaSQL
    const modifiedData = alasql.tables[tableName]?.data || [];
    
    // Convert back to array format and update sheet
    const sheet = allSheets[sheetIndex];
    const newData = modifiedData.map(row => {
        return sheet.headers.map((h, i) => {
            const cleanHeader = getColumnName(h, i);
            const val = row[cleanHeader];
            return val !== null && val !== undefined ? val.toString() : '';
        });
    });
    
    // Update the sheet data
    allSheets[sheetIndex].data = newData;
    
    // If current sheet was modified, update display
    if (sheetIndex === currentSheetIndex) {
        headers = [...sheet.headers];
        currentData = newData.map(row => [...row]);
        renderTable(currentData);
        rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
    }
    
    // Mark as edited
    markAsEdited();
    sqlTablesRegistered = false;
    hasSqlResults = false;
    hideSqlResultBar();
    
    // Show success message
    let affectedRows = typeof result === 'number' ? result : (Array.isArray(result) ? result.length : 1);
    const actionText = queryType === 'UPDATE' ? 'updated' : (queryType === 'DELETE' ? 'deleted' : 'inserted');
    showSqlStatus(`✓ ${queryType} executed in ${execTime}ms — ${affectedRows} row(s) ${actionText}`, 'success');
    
    showToast(`${affectedRows} row(s) ${actionText}`, 'success');
}

// Extract table name from query
function extractTableName(query, queryType) {
    const upperQuery = query.toUpperCase();
    let tableName = 'data';
    
    try {
        if (queryType === 'UPDATE') {
            // UPDATE tablename SET ...
            const match = query.match(/UPDATE\s+(\w+)/i);
            if (match) tableName = match[1];
        } else if (queryType === 'DELETE') {
            // DELETE FROM tablename ...
            const match = query.match(/DELETE\s+FROM\s+(\w+)/i);
            if (match) tableName = match[1];
        } else if (queryType === 'INSERT') {
            // INSERT INTO tablename ...
            const match = query.match(/INSERT\s+INTO\s+(\w+)/i);
            if (match) tableName = match[1];
        }
    } catch (e) {
        console.error('Error extracting table name:', e);
    }
    
    return tableName;
}

// Find sheet index by table name
function findSheetIndexByTableName(tableName) {
    // Check if it's the "data" alias (current sheet)
    if (tableName.toLowerCase() === 'data') {
        return currentSheetIndex;
    }
    
    // Find by sheet name
    for (let i = 0; i < allSheets.length; i++) {
        if (getTableName(allSheets[i].name).toLowerCase() === tableName.toLowerCase()) {
            return i;
        }
    }
    
    return -1;
}

// Show SQL status message
function showSqlStatus(message, type) {
    sqlStatus.textContent = message;
    sqlStatus.className = 'sql-status ' + type;
    
    // Auto-hide success messages after 5 seconds
    if (type === 'success') {
        setTimeout(() => {
            sqlStatus.className = 'sql-status';
        }, 5000);
    }
}

// Reset to original data
function resetToOriginalData() {
    // Reload from current sheet data (which may have edits)
    headers = [...allSheets[currentSheetIndex].headers];
    currentData = allSheets[currentSheetIndex].data.map(row => [...row]);
    
    // Clear SQL and search state
    sqlInput.value = '';
    sqlStatus.className = 'sql-status';
    sqlLineCount.textContent = '';
    searchInput.value = '';
    
    // Reset SQL results
    hasSqlResults = false;
    sqlResultData = null;
    hideSqlResultBar();
    
    // Reset editor height
    sqlInput.style.height = '';
    autoResizeSqlEditor();
    
    // Re-render table
    renderTable(currentData);
    rowCount.textContent = `${currentData.length} rows × ${headers.length} columns`;
}

// Show SQL result bar
function showSqlResultBar(rows, cols, execTime) {
    const resultBar = document.getElementById('sqlResultBar');
    const resultText = document.getElementById('sqlResultText');
    resultText.textContent = `Query returned ${rows} rows × ${cols} columns in ${execTime}ms`;
    resultBar.style.display = 'flex';
}

// Hide SQL result bar
function hideSqlResultBar() {
    document.getElementById('sqlResultBar').style.display = 'none';
}

// Open export modal for SQL results
function openExportResultsModal() {
    if (!hasSqlResults || !sqlResultData) {
        alert('No query results to export. Run a SQL query first.');
        return;
    }
    
    // Create a temporary sheet structure for the export modal
    const resultSheet = {
        name: 'SQL_Results',
        headers: sqlResultData.headers,
        data: sqlResultData.data
    };
    
    // Store original sheets and temporarily replace with results
    const originalSheets = allSheets;
    const originalIndex = currentSheetIndex;
    
    allSheets = [resultSheet];
    currentSheetIndex = 0;
    
    // Update modal for results export
    currentSheetNameBadge.textContent = 'SQL Results';
    sheetCountBadge.textContent = `${sqlResultData.rowCount} rows`;
    allSheetsOption.style.display = 'none';
    sheetsChecklist.style.display = 'none';
    document.querySelector('input[name="sheetOption"][value="current"]').checked = true;
    
    // Reset format selection
    selectedFormat = 'csv';  // Default to CSV for query results
    document.querySelectorAll('.format-option').forEach(o => {
        o.classList.toggle('selected', o.dataset.format === 'csv');
    });
    
    formatNote.innerHTML = `<strong>Exporting SQL Query Results:</strong><br><code style="font-size: 0.8rem; color: #666;">${truncateQuery(sqlResultData.query)}</code>`;
    formatNote.style.display = 'block';
    formatNote.style.background = '#e8f5e9';
    formatNote.style.color = '#2e7d32';
    
    // Show modal
    exportModal.classList.add('show');
    
    // Override export confirm to use results
    const originalConfirm = confirmExport;
    exportConfirm.onclick = function() {
        exportSqlResults();
        
        // Restore original sheets
        allSheets = originalSheets;
        currentSheetIndex = originalIndex;
        
        // Restore original confirm handler
        exportConfirm.onclick = originalConfirm;
        
        closeExportModal();
    };
    
    // Restore on cancel/close
    const restoreSheets = () => {
        allSheets = originalSheets;
        currentSheetIndex = originalIndex;
        exportConfirm.onclick = originalConfirm;
    };
    
    modalClose.onclick = () => { restoreSheets(); closeExportModal(); };
    exportCancel.onclick = () => { restoreSheets(); closeExportModal(); };
}

// Truncate long query for display
function truncateQuery(query) {
    const maxLen = 100;
    if (query.length <= maxLen) return query;
    return query.substring(0, maxLen) + '...';
}

// Export SQL query results
function exportSqlResults() {
    if (!sqlResultData) return;
    
    const baseName = 'sql_results_' + new Date().toISOString().slice(0, 10);
    
    // Create workbook with results
    const wb = XLSX.utils.book_new();
    const sheetData = [sqlResultData.headers, ...sqlResultData.data];
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(wb, ws, 'SQL_Results');
    
    let filename;
    
    switch (selectedFormat) {
        case 'csv':
            filename = `${baseName}.csv`;
            const csv = XLSX.utils.sheet_to_csv(ws);
            downloadBlob(csv, filename, 'text/csv;charset=utf-8;');
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
            const jsonData = {
                query: sqlResultData.query,
                exported: new Date().toISOString(),
                rowCount: sqlResultData.rowCount,
                headers: sqlResultData.headers,
                data: sqlResultData.data.map(row => {
                    const obj = {};
                    sqlResultData.headers.forEach((h, i) => {
                        obj[h] = row[i] || '';
                    });
                    return obj;
                })
            };
            downloadBlob(JSON.stringify(jsonData, null, 2), filename, 'application/json;charset=utf-8;');
            return;
            
        case 'html':
            filename = `${baseName}.html`;
            const htmlTable = XLSX.utils.sheet_to_html(ws);
            const fullHtml = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>SQL Query Results</title>
    <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .query-info { background: #f5f5f5; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .query-info code { display: block; margin-top: 10px; padding: 10px; background: #fff; border-radius: 4px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background: #667eea; color: white; }
        tr:nth-child(even) { background: #f9f9f9; }
        tr:hover { background: #f0f3ff; }
    </style>
</head>
<body>
    <div class="query-info">
        <strong>SQL Query:</strong>
        <code>${sqlResultData.query.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</code>
    </div>
    <p><strong>${sqlResultData.rowCount} rows</strong> exported on ${new Date().toLocaleString()}</p>
    ${htmlTable}
</body>
</html>`;
            downloadBlob(fullHtml, filename, 'text/html;charset=utf-8;');
            return;
            
        default:
            return;
    }
    
    showExportSuccess(filename);
}

// Open export modal
function openExportModal() {
    if (allSheets.length === 0) return;
    
    // Update current sheet name badge
    currentSheetNameBadge.textContent = allSheets[currentSheetIndex].name;
    
    // Update sheet count badge and visibility
    if (allSheets.length > 1) {
        sheetCountBadge.textContent = `${allSheets.length} sheets`;
        allSheetsOption.style.display = 'flex';
        sheetsChecklist.style.display = 'block';
        
        // Build sheet checkboxes
        sheetCheckboxes.innerHTML = allSheets.map((sheet, index) => `
            <label class="checkbox-option">
                <input type="checkbox" name="sheetCheck" value="${index}" checked>
                <span>${sheet.name}</span>
                <span style="color: #888; font-size: 0.8rem;">(${sheet.data.length} rows)</span>
            </label>
        `).join('');
    } else {
        allSheetsOption.style.display = 'none';
        sheetsChecklist.style.display = 'none';
    }
    
    // Reset to defaults
    selectedFormat = 'xlsx';
    document.querySelectorAll('.format-option').forEach(o => {
        o.classList.toggle('selected', o.dataset.format === 'xlsx');
    });
    document.querySelector('input[name="sheetOption"][value="current"]').checked = true;
    
    updateFormatNote();
    exportModal.classList.add('show');
}

// Close export modal
function closeExportModal() {
    exportModal.classList.remove('show');
}

// Update format note based on selection
function updateFormatNote() {
    const sheetOption = document.querySelector('input[name="sheetOption"]:checked').value;
    const isMultiSheet = sheetOption === 'all' || sheetOption === 'selected';
    
    if (selectedFormat === 'csv') {
        if (isMultiSheet && allSheets.length > 1) {
            formatNote.textContent = '⚠️ CSV format will create separate files for each sheet (downloaded as ZIP)';
            formatNote.style.display = 'block';
        } else {
            formatNote.textContent = '';
            formatNote.style.display = 'none';
        }
    } else if (selectedFormat === 'json' || selectedFormat === 'html') {
        if (isMultiSheet && allSheets.length > 1) {
            formatNote.textContent = `ℹ️ ${selectedFormat.toUpperCase()} will include all selected sheets in one file`;
            formatNote.style.display = 'block';
        } else {
            formatNote.style.display = 'none';
        }
    } else {
        formatNote.style.display = 'none';
    }
}

// Confirm and execute export
function confirmExport() {
    const sheetOption = document.querySelector('input[name="sheetOption"]:checked').value;
    
    // Determine which sheets to export
    let sheetsToExport = [];
    
    if (sheetOption === 'current') {
        sheetsToExport = [currentSheetIndex];
    } else if (sheetOption === 'all') {
        sheetsToExport = allSheets.map((_, i) => i);
    } else if (sheetOption === 'selected') {
        sheetsToExport = Array.from(document.querySelectorAll('input[name="sheetCheck"]:checked'))
            .map(cb => parseInt(cb.value));
        
        if (sheetsToExport.length === 0) {
            alert('Please select at least one sheet to export');
            return;
        }
    }
    
    exportData(selectedFormat, sheetsToExport);
    closeExportModal();
}

// Export data in various formats
function exportData(format, sheetIndices) {
    if (allSheets.length === 0) return;
    
    // Get base filename without extension
    const baseName = currentFile ? currentFile.name.replace(/\.[^/.]+$/, '') : 'export';
    
    // Get sheets to export
    const sheetsToExport = sheetIndices.map(i => allSheets[i]);
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    
    sheetsToExport.forEach(sheet => {
        const sheetData = [sheet.headers, ...sheet.data];
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    
    // Generate and download file based on format
    let filename;
    
    switch (format) {
        case 'csv':
            if (sheetsToExport.length === 1) {
                filename = `${baseName}.csv`;
                downloadCSV(wb, filename);
            } else {
                downloadMultiCSV(sheetsToExport, baseName);
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
            downloadJSON(sheetsToExport, filename);
            return;
            
        case 'html':
            filename = `${baseName}.html`;
            downloadHTML(sheetsToExport, filename);
            return;
            
        default:
            return;
    }
    
    showExportSuccess(filename);
}

// Download multiple CSVs (one per sheet)
function downloadMultiCSV(sheets, baseName) {
    sheets.forEach(sheet => {
        const wb = XLSX.utils.book_new();
        const sheetData = [sheet.headers, ...sheet.data];
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(wb, ws, sheet.name);
        
        const filename = `${baseName}_${sheet.name}.csv`;
        const csv = XLSX.utils.sheet_to_csv(ws);
        downloadBlob(csv, filename, 'text/csv;charset=utf-8;');
    });
    
    showExportSuccess(`${sheets.length} CSV files`);
}

// Download as CSV
function downloadCSV(wb, filename) {
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
    downloadBlob(csv, filename, 'text/csv;charset=utf-8;');
}

// Download as JSON (supports multiple sheets)
function downloadJSON(sheets, filename) {
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
    
    const json = JSON.stringify(jsonData, null, 2);
    downloadBlob(json, filename, 'application/json;charset=utf-8;');
}

// Download as HTML (supports multiple sheets)
function downloadHTML(sheets, filename) {
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
        .tabs { display: flex; gap: 5px; margin-bottom: 20px; flex-wrap: wrap; }
        .tab-btn { 
            padding: 10px 20px; 
            border: none; 
            background: #e0e0e0; 
            cursor: pointer; 
            border-radius: 8px 8px 0 0;
            font-weight: 500;
        }
        .tab-btn.active { background: #667eea; color: white; }
        .tab-btn:hover { background: #d0d0d0; }
        .tab-btn.active:hover { background: #5a6fd6; }
        h2 { color: #333; margin-bottom: 15px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 30px; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background: #667eea; color: white; }
        tr:nth-child(even) { background: #f9f9f9; }
        tr:hover { background: #f0f3ff; }
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
    downloadBlob(fullHtml, filename, 'text/html;charset=utf-8;');
}

// Helper to download blob
function downloadBlob(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showExportSuccess(filename);
}

// Show export success notification
function showExportSuccess(filename) {
    // Create toast notification
    const toast = document.createElement('div');
    toast.className = 'toast-notification';
    toast.innerHTML = `✅ Exported: <strong>${filename}</strong>`;
    document.body.appendChild(toast);
    
    // Animate in
    setTimeout(() => toast.classList.add('show'), 10);
    
    // Remove after 3 seconds
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => document.body.removeChild(toast), 300);
    }, 3000);
}

// Clear all data
function clearData() {
    // Warn about unsaved changes
    if (hasUnsavedChanges || docHasUnsavedChanges) {
        if (!confirm('You have unsaved changes. Are you sure you want to clear?')) {
            return;
        }
    }
    
    // Clear spreadsheet data
    currentData = [];
    headers = [];
    allSheets = [];
    currentSheetIndex = 0;
    currentFile = null;
    fileHandle = null;
    sqlTablesRegistered = false;
    hasSqlResults = false;
    sqlResultData = null;
    hasUnsavedChanges = false;
    originalSheetData = [];
    autoSaveEnabled = false;
    activeFilters = [];
    resetSort();
    
    // Clear document data
    currentMode = 'none';
    docHasUnsavedChanges = false;
    originalDocContent = '';
    docFileHandle = null;
    docEditor.innerHTML = '<p></p>';
    
    document.getElementById('autosaveCheckbox').checked = false;
    document.getElementById('quickFilterPanel').style.display = 'none';
    document.getElementById('filterToggleBtn').classList.remove('active');
    document.getElementById('activeFilters').style.display = 'none';
    fileInput.value = '';
    searchInput.value = '';
    sqlInput.value = '';
    
    // Hide all content sections; show upload section
    fileInfo.style.display = 'none';
    docInfo.style.display = 'none';
    document.getElementById('editIndicator').style.display = 'none';
    document.getElementById('docEditIndicator').style.display = 'none';
    document.getElementById('saveStatus').style.display = 'none';
    sqlSection.style.display = 'none';
    sqlStatus.className = 'sql-status';
    sqlExamplesPanel.style.display = 'none';
    hideSqlResultBar();
    tableContainer.style.display = 'none';
    tableContainer.innerHTML = '';
    documentContainer.style.display = 'none';
    sheetTabs.style.display = 'none';
    sheetTabs.innerHTML = '';
    uploadSection.style.display = 'flex';
    clearBtn.style.display = 'none';
}

// Utility functions
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ==========================================
// DOCUMENT EDITOR FUNCTIONALITY
// ==========================================

// Parse document file (DOCX)
async function parseDocument(arrayBuffer, file) {
    try {
        const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        const html = result.value;
        
        // Display the document
        currentFile = file;
        currentMode = 'document';
        originalDocContent = html;
        docHasUnsavedChanges = false;
        
        docEditor.innerHTML = html || '<p>Start typing your document here...</p>';
        
        displayDocument(file);
        updateDocStats();
        
        if (result.messages.length > 0) {
            console.log('Document conversion messages:', result.messages);
        }
        
    } catch (error) {
        console.error('Error parsing document:', error);
        alert('Error reading document file. Please try again.');
        clearData();
    }
}

// Display document UI
function displayDocument(file) {
    // Update document info bar
    const canSaveBack = docFileHandle !== null;
    document.getElementById('docFileName').textContent = '📄 ' + file.name + (canSaveBack ? ' ✓' : '');
    document.getElementById('docFileSize').textContent = formatFileSize(file.size);
    
    // Show document UI, hide spreadsheet UI
    docInfo.style.display = 'flex';
    fileInfo.style.display = 'none';
    sqlSection.style.display = 'none';
    tableContainer.style.display = 'none';
    documentContainer.style.display = 'flex';
    uploadSection.style.display = 'none';
    clearBtn.style.display = 'inline-block';
    sheetTabs.style.display = 'none';
    
    // Reset edit indicator
    document.getElementById('docEditIndicator').style.display = 'none';
    
    // Focus editor
    docEditor.focus();
}

// Create empty document
function createEmptyDocument() {
    // Warn about unsaved changes
    if (hasUnsavedChanges || docHasUnsavedChanges) {
        if (!confirm('You have unsaved changes. Creating a new document will clear them. Continue?')) {
            return;
        }
    }
    
    // Clear spreadsheet data
    clearDataSilent();
    
    // Set up new document
    currentMode = 'document';
    currentFile = { name: 'new_document.docx', size: 0 };
    originalDocContent = '';
    docHasUnsavedChanges = true;
    docFileHandle = null; // No handle for new files
    
    // Set empty content with placeholder
    docEditor.innerHTML = '<p></p>';
    
    // Update UI
    document.getElementById('docFileName').textContent = '📄 new_document.docx (New)';
    document.getElementById('docFileSize').textContent = 'New Document';
    
    // Show document UI
    docInfo.style.display = 'flex';
    fileInfo.style.display = 'none';
    sqlSection.style.display = 'none';
    tableContainer.style.display = 'none';
    documentContainer.style.display = 'flex';
    uploadSection.style.display = 'none';
    clearBtn.style.display = 'inline-block';
    sheetTabs.style.display = 'none';
    
    // Show edit indicator for new docs
    document.getElementById('docEditIndicator').style.display = 'flex';
    
    // Focus editor
    docEditor.focus();
    updateDocStats();
    
    showToast('New document created! Start typing to add content.', 'success');
}

// Mark document as edited
function markDocAsEdited() {
    if (!docHasUnsavedChanges) {
        docHasUnsavedChanges = true;
        document.getElementById('docEditIndicator').style.display = 'flex';
    }
}

// Update document statistics
function updateDocStats() {
    const text = docEditor.innerText || '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    
    document.getElementById('docWordCount').textContent = `${words} word${words !== 1 ? 's' : ''}`;
    document.getElementById('docCharCount').textContent = `${chars} character${chars !== 1 ? 's' : ''}`;
}

// Update toolbar button states based on current selection
function updateToolbarState() {
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
}

// Save document
async function saveDocument() {
    if (currentMode !== 'document') {
        return;
    }
    
    try {
        // Check if we have a valid file handle to save to
        if (docFileHandle !== null && docFileHandle !== undefined) {
            // Try to save to existing file
            try {
                await saveDocumentToHandle();
                onDocumentSaveSuccess();
                return;
            } catch (handleError) {
                console.warn('Could not save to existing handle, will prompt for new location:', handleError);
                // Handle might be invalid, try with picker
                docFileHandle = null;
            }
        }
        
        // No valid handle - need to create new file
        if ('showSaveFilePicker' in window) {
            await saveDocumentWithPicker();
            onDocumentSaveSuccess();
        } else {
            // Fallback - just download the file
            await exportDocumentAsDocx();
            onDocumentSaveSuccess();
        }
        
    } catch (error) {
        if (error.name === 'AbortError') {
            // User cancelled the save dialog
            return;
        }
        console.error('Save error:', error);
        showToast('Failed to save document: ' + error.message, 'error');
    }
}

// Called when document save is successful
function onDocumentSaveSuccess() {
    docHasUnsavedChanges = false;
    document.getElementById('docEditIndicator').style.display = 'none';
    showToast('Document saved!', 'success');
    
    // Update file name display
    const docFileNameEl = document.getElementById('docFileName');
    if (currentFile) {
        docFileNameEl.textContent = '📄 ' + currentFile.name + (docFileHandle ? ' ✓' : '');
    }
}

// Save document using File System Access API picker (first time save)
async function saveDocumentWithPicker() {
    const fileName = currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
    
    const handle = await window.showSaveFilePicker({
        suggestedName: `${fileName}.docx`,
        types: [{
            description: 'Word Document',
            accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] }
        }]
    });
    
    // Store the handle globally for future saves
    docFileHandle = handle;
    currentFile = { name: handle.name, size: 0 };
    
    // Actually save the content
    await saveDocumentToHandle();
    
    // Update the file name display
    document.getElementById('docFileName').textContent = '📄 ' + handle.name + ' ✓';
    document.getElementById('docFileSize').textContent = 'Saved';
}

// Save document to existing file handle
async function saveDocumentToHandle() {
    if (!docFileHandle) {
        throw new Error('No file handle available');
    }
    
    const { Document, Packer, Paragraph, TextRun, HeadingLevel } = docx;
    
    // Get HTML content and convert to DOCX
    const htmlContent = docEditor.innerHTML;
    const paragraphs = htmlToDocxParagraphs(htmlContent);
    
    const doc = new Document({
        sections: [{
            properties: {},
            children: paragraphs
        }]
    });
    
    const blob = await Packer.toBlob(doc);
    
    // Write to file handle
    const writable = await docFileHandle.createWritable();
    await writable.write(blob);
    await writable.close();
}

// Export document
function exportDocument() {
    if (!currentFile) {
        showToast('No document to export', 'error');
        return;
    }
    
    const format = prompt('Export format:\n1. DOCX (Word)\n2. HTML\n3. TXT (Plain Text)\n\nEnter 1, 2, or 3:', '1');
    
    switch (format) {
        case '1':
            exportDocumentAsDocx();
            break;
        case '2':
            exportDocumentAsHtml();
            break;
        case '3':
            exportDocumentAsText();
            break;
        default:
            if (format !== null) {
                showToast('Invalid format selected', 'warning');
            }
    }
}

// Export as DOCX
async function exportDocumentAsDocx() {
    const fileName = currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
    
    // Get HTML content and convert to DOCX using docx library
    const htmlContent = docEditor.innerHTML;
    
    // Create a simple DOCX document
    const { Document, Packer, Paragraph, TextRun, HeadingLevel } = docx;
    
    // Parse HTML and create paragraphs
    const paragraphs = htmlToDocxParagraphs(htmlContent);
    
    const doc = new Document({
        sections: [{
            properties: {},
            children: paragraphs
        }]
    });
    
    const blob = await Packer.toBlob(doc);
    downloadBlobFile(blob, `${fileName}.docx`);
}

// Convert HTML to DOCX paragraphs
function htmlToDocxParagraphs(html) {
    const { Paragraph, TextRun, HeadingLevel } = docx;
    const paragraphs = [];
    
    // Create a temporary element to parse HTML
    const temp = document.createElement('div');
    temp.innerHTML = html;
    
    // Process each child node
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
                    parseInlineElements(node, runs);
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
                    // Process children
                    node.childNodes.forEach(child => processNode(child));
            }
        }
    };
    
    temp.childNodes.forEach(child => processNode(child));
    
    // Ensure at least one paragraph
    if (paragraphs.length === 0) {
        paragraphs.push(new Paragraph({ children: [new TextRun('')] }));
    }
    
    return paragraphs;
}

// Parse inline elements for text runs
function parseInlineElements(node, runs) {
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
}

// Export as HTML
function exportDocumentAsHtml() {
    const fileName = currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
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
${docEditor.innerHTML}
</body>
</html>`;
    
    downloadBlob(htmlContent, `${fileName}.html`, 'text/html;charset=utf-8;');
}

// Export as plain text
function exportDocumentAsText() {
    const fileName = currentFile?.name?.replace(/\.[^/.]+$/, '') || 'document';
    const textContent = docEditor.innerText || '';
    
    downloadBlob(textContent, `${fileName}.txt`, 'text/plain;charset=utf-8;');
}

// Helper to download blob as file
function downloadBlobFile(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showExportSuccess(filename);
}

// ==========================================
// CREATE NEW SPREADSHEET FUNCTIONALITY
// ==========================================

// Create empty spreadsheet with default columns
function createEmptySpreadsheet() {
    // Warn about unsaved changes
    if (hasUnsavedChanges || docHasUnsavedChanges) {
        if (!confirm('You have unsaved changes. Creating a new spreadsheet will clear them. Continue?')) {
            return;
        }
    }
    
    // Clear any existing data
    clearDataSilent();
    
    // Set mode
    currentMode = 'spreadsheet';
    
    // Default configuration - start completely empty
    const sheetName = 'Sheet1';
    const fileName = 'new_spreadsheet.xlsx';
    
    // Set up empty sheet - no columns, no rows
    headers = [];
    currentData = [];
    
    // Create sheet structure
    allSheets = [{
        name: sheetName,
        headers: headers,
        data: currentData
    }];
    
    currentSheetIndex = 0;
    
    // Create a virtual file object
    currentFile = {
        name: fileName,
        size: 0
    };
    
    // No file handle for new files - will download on save
    fileHandle = null;
    
    // Display the new spreadsheet
    displayNewSpreadsheet(fileName);
    
    showToast('New spreadsheet created! Add columns and rows to get started.', 'success');
}

// Display the newly created spreadsheet
function displayNewSpreadsheet(fileName) {
    // Update file info bar
    const fileNameEl = document.getElementById('fileName');
    fileNameEl.textContent = '📄 ' + fileName + ' (New)';
    fileNameEl.title = 'New file - will be downloaded when saving';
    
    document.getElementById('fileSize').textContent = 'New File';
    updateRowColCount();
    
    // Show UI elements
    fileInfo.style.display = 'flex';
    sqlSection.style.display = 'block';
    tableContainer.style.display = 'block';
    uploadSection.style.display = 'none';
    clearBtn.style.display = 'inline-block';
    
    // Hide sheet tabs for single sheet
    sheetTabs.style.display = 'none';
    
    // New files start as "unsaved" since user needs to save them
    hasUnsavedChanges = true;
    document.getElementById('editIndicator').style.display = 'flex';
    document.getElementById('saveStatus').style.display = 'none';
    
    // Can't auto-save new files (no file handle)
    const autosaveToggle = document.getElementById('autosaveToggle');
    autosaveToggle.style.display = 'none';
    autosaveToggle.title = 'Auto-save not available for new files. Save to create the file first.';
    document.getElementById('autosaveCheckbox').checked = false;
    autoSaveEnabled = false;
    
    // Reset sort and filter states
    resetSort();
    activeFilters = [];
    updateActiveFiltersUI();
    
    // Render the empty state or table
    renderTableOrEmpty();
}

// Update row/column count display
function updateRowColCount() {
    const colCount = headers.length;
    const rowCount = currentData.length;
    
    if (colCount === 0 && rowCount === 0) {
        document.getElementById('rowCount').textContent = 'Empty sheet';
    } else {
        document.getElementById('rowCount').textContent = `${rowCount} rows × ${colCount} columns`;
    }
}

// Render table or show empty state message
function renderTableOrEmpty() {
    if (headers.length === 0) {
        // Show empty state with instructions
        tableContainer.innerHTML = `
            <div class="empty-sheet-message">
                <div class="empty-icon">📋</div>
                <h3>Empty Spreadsheet</h3>
                <p>Start building your spreadsheet:</p>
                <div class="empty-actions">
                    <button class="empty-action-btn" id="emptyAddColBtn">
                        <span>➕</span> Add Column
                    </button>
                </div>
                <p class="empty-hint">Add columns first, then add rows to enter data</p>
            </div>
        `;
        
        // Bind the button
        document.getElementById('emptyAddColBtn').addEventListener('click', addNewColumn);
        
        // Hide SQL section when empty
        sqlSection.style.display = 'none';
    } else {
        // Show SQL section
        sqlSection.style.display = 'block';
        registerSqlTables();
        updateSqlPlaceholder();
        
        // Render the table
        renderTable(currentData);
    }
}

// Add new column
function addNewColumn() {
    if (allSheets.length === 0) {
        showToast('No spreadsheet loaded', 'warning');
        return;
    }
    
    if (hasSqlResults) {
        showToast('Cannot add columns to SQL results. Reset first.', 'warning');
        return;
    }
    
    // Generate new column name
    const newColIndex = headers.length;
    const newColName = `Column ${String.fromCharCode(65 + (newColIndex % 26))}${newColIndex >= 26 ? Math.floor(newColIndex / 26) : ''}`;
    
    // Prompt for column name
    const colName = prompt('Enter column name:', newColName);
    if (colName === null) return; // User cancelled
    
    const finalColName = colName.trim() || newColName;
    
    // Store original data for undo if not already stored
    if (!hasUnsavedChanges) {
        storeOriginalData();
    }
    
    // Add to headers
    headers.push(finalColName);
    allSheets[currentSheetIndex].headers = headers;
    
    // Add empty cell to each existing row
    allSheets[currentSheetIndex].data.forEach(row => {
        row.push('');
    });
    currentData.forEach(row => {
        row.push('');
    });
    
    // Re-render (handles empty state too)
    renderTableOrEmpty();
    updateRowColCount();
    
    // Mark as edited
    markAsEdited();
    sqlTablesRegistered = false;
    
    showToast(`Added column: ${finalColName}`, 'success');
}

// Clear data without confirmation (for internal use)
function clearDataSilent() {
    // Clear spreadsheet data
    currentData = [];
    headers = [];
    allSheets = [];
    currentSheetIndex = 0;
    currentFile = null;
    fileHandle = null;
    sqlTablesRegistered = false;
    hasSqlResults = false;
    sqlResultData = null;
    hasUnsavedChanges = false;
    originalSheetData = [];
    autoSaveEnabled = false;
    activeFilters = [];
    resetSort();
    
    // Clear document data
    docHasUnsavedChanges = false;
    originalDocContent = '';
    docFileHandle = null;
    
    document.getElementById('autosaveCheckbox').checked = false;
    document.getElementById('quickFilterPanel').style.display = 'none';
    document.getElementById('filterToggleBtn').classList.remove('active');
    document.getElementById('activeFilters').style.display = 'none';
    fileInput.value = '';
    searchInput.value = '';
    sqlInput.value = '';
    
    // Reset UI elements
    document.getElementById('editIndicator').style.display = 'none';
    document.getElementById('docEditIndicator').style.display = 'none';
    document.getElementById('saveStatus').style.display = 'none';
    sqlStatus.className = 'sql-status';
    sqlExamplesPanel.style.display = 'none';
    hideSqlResultBar();
    tableContainer.innerHTML = '';
    sheetTabs.style.display = 'none';
    sheetTabs.innerHTML = '';
    docInfo.style.display = 'none';
    documentContainer.style.display = 'none';
}
