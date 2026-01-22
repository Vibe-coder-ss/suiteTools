/**
 * Spreadsheet Module
 * Handles spreadsheet creation, editing, and rendering
 */

const Spreadsheet = {
    // Create empty spreadsheet
    async createEmpty() {
        const state = AppState;
        
        if (state.hasUnsavedChanges || state.docHasUnsavedChanges || state.slidesHasUnsavedChanges) {
            const confirmed = await Utils.confirm('You have unsaved changes. Creating a new spreadsheet will clear them. Continue?', {
                title: 'Unsaved Changes',
                icon: '⚠️',
                okText: 'Continue',
                cancelText: 'Cancel',
                danger: true
            });
            if (!confirmed) return;
        }
        
        // Clear any existing data
        this.clearSilent();
        
        // Set mode
        state.currentMode = 'spreadsheet';
        
        // Default configuration - start completely empty
        const sheetName = 'Sheet1';
        const fileName = 'new_spreadsheet.xlsx';
        
        // Set up empty sheet - no columns, no rows
        state.headers = [];
        state.currentData = [];
        
        // Create sheet structure
        state.allSheets = [{
            name: sheetName,
            headers: state.headers,
            data: state.currentData
        }];
        
        state.currentSheetIndex = 0;
        state.currentFile = { name: fileName, size: 0 };
        state.fileHandle = null;
        
        // Display the new spreadsheet
        this.displayNew(fileName);
        
        Utils.showToast('New spreadsheet created! Add columns and rows to get started.', 'success');
    },

    // Display newly created spreadsheet
    displayNew(fileName) {
        const state = AppState;
        
        DOM.fileName.textContent = '📄 ' + fileName + ' (New)';
        DOM.fileName.title = 'New file - will be downloaded when saving';
        
        DOM.fileSize.textContent = 'New File';
        this.updateRowColCount();
        
        // Show UI elements
        DOM.fileInfo.style.display = 'flex';
        DOM.sqlSection.style.display = 'block';
        DOM.tableContainer.style.display = 'block';
        DOM.uploadSection.style.display = 'none';
        DOM.clearBtn.style.display = 'inline-block';
        DOM.sheetTabs.style.display = 'none';
        DOM.docInfo.style.display = 'none';
        DOM.documentContainer.style.display = 'none';
        
        // New files start as "unsaved"
        state.hasUnsavedChanges = true;
        DOM.editIndicator.style.display = 'flex';
        DOM.saveStatus.style.display = 'none';
        
        // Can't auto-save new files
        DOM.autosaveToggle.style.display = 'none';
        DOM.autosaveCheckbox.checked = false;
        state.autoSaveEnabled = false;
        
        // Reset sort and filter states
        this.resetSort();
        state.activeFilters = [];
        Filter.updateActiveFiltersUI();
        
        // Render the empty state or table
        this.renderTableOrEmpty();
    },

    // Update row/column count display
    updateRowColCount() {
        const state = AppState;
        const colCount = state.headers.length;
        const rowCount = state.currentData.length;
        
        if (colCount === 0 && rowCount === 0) {
            DOM.rowCount.textContent = 'Empty sheet';
        } else {
            DOM.rowCount.textContent = `${rowCount} rows × ${colCount} columns`;
        }
    },

    // Render table or show empty state message
    renderTableOrEmpty() {
        const state = AppState;
        
        if (state.headers.length === 0) {
            DOM.tableContainer.innerHTML = `
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
            
            document.getElementById('emptyAddColBtn').addEventListener('click', () => this.addColumn());
            DOM.sqlSection.style.display = 'none';
        } else {
            DOM.sqlSection.style.display = 'block';
            SQL.registerTables();
            SQL.updatePlaceholder();
            this.renderTable(state.currentData);
        }
    },

    // Render table with data
    renderTable(data, isFiltered = false) {
        const state = AppState;
        
        if (data.length === 0 && state.headers.length === 0) {
            DOM.tableContainer.innerHTML = '<p class="placeholder-text">No data to display</p>';
            return;
        }

        const isOriginalData = !isFiltered && !state.hasSqlResults;
        
        let html = '<div class="table-wrapper"><table class="data-table" id="dataTable">';
        
        // Header with sort indicators
        html += '<thead><tr><th class="row-num">#</th>';
        state.headers.forEach((header, colIndex) => {
            const isSorted = state.sortColumn === colIndex;
            const sortClass = isSorted ? `sortable sorted-${state.sortDirection}` : 'sortable';
            const sortIcon = isSorted 
                ? (state.sortDirection === 'asc' ? '↑' : '↓') 
                : '⇅';
            html += `<th data-col="${colIndex}" class="${sortClass}" title="Click to sort • Double-click to rename">${Utils.escapeHtml(header)}<span class="sort-icon">${sortIcon}</span></th>`;
        });
        html += '</tr></thead>';

        // Body
        html += '<tbody>';
        if (data.length === 0) {
            html += `<tr class="empty-row-message">
                <td colspan="${state.headers.length + 1}" class="empty-table-cell">
                    <span>No rows yet.</span>
                    <button class="inline-add-row-btn" id="inlineAddRowBtn">➕ Add Row</button>
                </td>
            </tr>`;
        } else {
            data.forEach((row, displayIndex) => {
                const originalRowIndex = isOriginalData ? displayIndex : this.findOriginalRowIndex(row);
                const canEdit = isOriginalData && originalRowIndex !== -1;
                
                html += `<tr data-row="${displayIndex}" data-original-row="${originalRowIndex}">`;
                html += `<td class="row-num">
                    <span class="row-number">${displayIndex + 1}</span>
                    ${canEdit ? `<button class="delete-row-btn" data-original-row="${originalRowIndex}" title="Delete row">🗑️</button>` : ''}
                </td>`;
                state.headers.forEach((_, colIndex) => {
                    const value = row[colIndex] || '';
                    const editAttr = canEdit ? `data-original-row="${originalRowIndex}"` : '';
                    const editTitle = canEdit ? 'Double-click to edit' : 'Cannot edit filtered/SQL results';
                    const editClass = canEdit ? 'editable' : 'readonly';
                    html += `<td data-row="${displayIndex}" data-col="${colIndex}" ${editAttr} class="${editClass}" title="${editTitle}">${Utils.escapeHtml(value)}</td>`;
                });
                html += '</tr>';
            });
        }
        html += '</tbody></table></div>';

        DOM.tableContainer.innerHTML = html;
        
        // Bind inline add row button if present
        const inlineAddRowBtn = document.getElementById('inlineAddRowBtn');
        if (inlineAddRowBtn) {
            inlineAddRowBtn.addEventListener('click', () => this.addRow());
        }
    },

    // Find original row index by matching row data
    findOriginalRowIndex(row) {
        const state = AppState;
        const sheetData = state.allSheets[state.currentSheetIndex]?.data;
        if (!sheetData) return -1;
        
        for (let i = 0; i < sheetData.length; i++) {
            if (Utils.arraysEqual(sheetData[i], row)) {
                return i;
            }
        }
        return -1;
    },

    // Add new column
    async addColumn() {
        const state = AppState;
        
        if (state.allSheets.length === 0) {
            Utils.showToast('No spreadsheet loaded', 'warning');
            return;
        }
        
        if (state.hasSqlResults) {
            Utils.showToast('Cannot add columns to SQL results. Reset first.', 'warning');
            return;
        }
        
        const newColIndex = state.headers.length;
        const newColName = `Column ${String.fromCharCode(65 + (newColIndex % 26))}${newColIndex >= 26 ? Math.floor(newColIndex / 26) : ''}`;
        
        const colName = await Utils.prompt('Enter column name:', newColName, {
            title: 'Add Column',
            icon: '➕',
            okText: 'Add',
            placeholder: 'Column name'
        });
        if (colName === null) return;
        
        const finalColName = colName.trim() || newColName;
        
        if (!state.hasUnsavedChanges) {
            this.storeOriginalData();
        }
        
        state.headers.push(finalColName);
        state.allSheets[state.currentSheetIndex].headers = state.headers;
        
        state.allSheets[state.currentSheetIndex].data.forEach(row => {
            row.push('');
        });
        state.currentData.forEach(row => {
            row.push('');
        });
        
        this.renderTableOrEmpty();
        this.updateRowColCount();
        this.markAsEdited();
        state.sqlTablesRegistered = false;
        
        Utils.showToast(`Added column: ${finalColName}`, 'success');
    },

    // Add new row
    addRow() {
        const state = AppState;
        
        if (state.hasSqlResults) {
            Utils.showToast('Cannot add rows to SQL results. Reset first.', 'warning');
            return;
        }
        
        if (state.allSheets.length === 0) return;
        
        if (state.headers.length === 0) {
            Utils.showToast('Add columns first before adding rows', 'warning');
            return;
        }
        
        if (DOM.searchInput.value.trim()) {
            DOM.searchInput.value = '';
        }
        
        if (!state.hasUnsavedChanges) {
            this.storeOriginalData();
        }
        
        const newRow = state.headers.map(() => '');
        state.currentData.push(newRow);
        state.allSheets[state.currentSheetIndex].data = state.currentData;
        
        this.renderTable(state.currentData);
        this.updateRowColCount();
        this.markAsEdited();
        
        const tableWrapper = document.querySelector('.table-wrapper');
        if (tableWrapper) {
            tableWrapper.scrollTop = tableWrapper.scrollHeight;
        }
        
        setTimeout(() => {
            const lastRow = document.querySelector(`tr[data-row="${state.currentData.length - 1}"]`);
            if (lastRow) {
                const firstCell = lastRow.querySelector('td[data-col="0"].editable');
                if (firstCell) {
                    this.startCellEdit(firstCell);
                }
            }
        }, 100);
        
        state.sqlTablesRegistered = false;
        Utils.showToast('New row added', 'success');
    },

    // Delete row
    deleteRow(originalRowIndex) {
        const state = AppState;
        
        if (state.hasSqlResults) {
            Utils.showToast('Cannot delete rows from SQL results. Reset first.', 'warning');
            return;
        }
        
        const sheetData = state.allSheets[state.currentSheetIndex]?.data;
        if (!sheetData || originalRowIndex < 0 || originalRowIndex >= sheetData.length) return;
        
        if (!state.hasUnsavedChanges) {
            this.storeOriginalData();
        }
        
        state.allSheets[state.currentSheetIndex].data.splice(originalRowIndex, 1);
        state.currentData = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
        
        this.renderTable(state.currentData);
        this.updateRowColCount();
        this.markAsEdited();
        state.sqlTablesRegistered = false;
        
        Utils.showToast('Row deleted', 'info');
    },

    // Start cell editing
    startCellEdit(cell) {
        const state = AppState;
        const displayRowIndex = parseInt(cell.dataset.row);
        const originalRowIndex = parseInt(cell.dataset.originalRow);
        const colIndex = parseInt(cell.dataset.col);
        
        const currentValue = state.allSheets[state.currentSheetIndex]?.data[originalRowIndex]?.[colIndex] || '';
        
        if (!state.hasUnsavedChanges) {
            this.storeOriginalData();
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
        
        input.addEventListener('blur', () => {
            this.finishCellEdit(cell, input, originalRowIndex, colIndex, displayRowIndex);
        });
        
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                input.blur();
                const nextRow = cell.parentElement.nextElementSibling;
                if (nextRow) {
                    const nextCell = nextRow.querySelector(`td[data-col="${colIndex}"].editable`);
                    if (nextCell) {
                        setTimeout(() => this.startCellEdit(nextCell), 50);
                    }
                }
            } else if (e.key === 'Tab') {
                e.preventDefault();
                input.blur();
                const nextColIndex = e.shiftKey ? colIndex - 1 : colIndex + 1;
                const nextCell = cell.parentElement.querySelector(`td[data-col="${nextColIndex}"].editable`);
                if (nextCell) {
                    setTimeout(() => this.startCellEdit(nextCell), 50);
                }
            } else if (e.key === 'Escape') {
                cell.classList.remove('editing');
                cell.innerHTML = Utils.escapeHtml(currentValue);
                cell.title = 'Double-click to edit';
            }
        });
    },

    // Finish cell editing
    finishCellEdit(cell, input, originalRowIndex, colIndex, displayRowIndex) {
        const state = AppState;
        const newValue = input.value;
        const oldValue = state.allSheets[state.currentSheetIndex]?.data[originalRowIndex]?.[colIndex] || '';
        
        if (state.allSheets[state.currentSheetIndex]?.data[originalRowIndex]) {
            state.allSheets[state.currentSheetIndex].data[originalRowIndex][colIndex] = newValue;
        }
        
        if (state.currentData[displayRowIndex]) {
            state.currentData[displayRowIndex][colIndex] = newValue;
        }
        
        cell.classList.remove('editing');
        cell.innerHTML = Utils.escapeHtml(newValue);
        cell.title = 'Double-click to edit';
        
        if (newValue !== oldValue) {
            this.markAsEdited();
            cell.classList.add('cell-changed');
            state.sqlTablesRegistered = false;
        }
    },

    // Start header editing
    startHeaderEdit(th, colIndex) {
        const state = AppState;
        const currentValue = state.headers[colIndex] || '';
        
        if (!state.hasUnsavedChanges) {
            this.storeOriginalData();
        }
        
        th.classList.add('editing');
        
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
        
        input.addEventListener('blur', () => {
            this.finishHeaderEdit(th, input, colIndex, sortIconHtml);
        });
        
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                input.blur();
            } else if (e.key === 'Escape') {
                th.classList.remove('editing');
                th.innerHTML = Utils.escapeHtml(currentValue) + sortIconHtml;
            } else if (e.key === 'Tab') {
                e.preventDefault();
                input.blur();
                const nextColIndex = e.shiftKey ? colIndex - 1 : colIndex + 1;
                if (nextColIndex >= 0 && nextColIndex < state.headers.length) {
                    const nextTh = document.querySelector(`th[data-col="${nextColIndex}"]`);
                    if (nextTh) {
                        setTimeout(() => this.startHeaderEdit(nextTh, nextColIndex), 50);
                    }
                }
            }
        });
        
        input.addEventListener('click', (e) => {
            e.stopPropagation();
        });
    },

    // Finish header editing
    finishHeaderEdit(th, input, colIndex, sortIconHtml) {
        const state = AppState;
        const newValue = input.value.trim();
        const oldValue = state.headers[colIndex] || '';
        
        const finalValue = newValue || `Column ${String.fromCharCode(65 + colIndex)}`;
        
        state.headers[colIndex] = finalValue;
        state.allSheets[state.currentSheetIndex].headers[colIndex] = finalValue;
        
        th.classList.remove('editing');
        
        const isSorted = state.sortColumn === colIndex;
        const sortIcon = isSorted 
            ? (state.sortDirection === 'asc' ? '↑' : '↓') 
            : '⇅';
        
        th.innerHTML = Utils.escapeHtml(finalValue) + `<span class="sort-icon">${sortIcon}</span>`;
        
        if (finalValue !== oldValue) {
            this.markAsEdited();
            state.sqlTablesRegistered = false;
            SQL.updatePlaceholder();
        }
    },

    // Store original data for undo
    storeOriginalData() {
        const state = AppState;
        state.originalSheetData = state.allSheets.map(sheet => ({
            name: sheet.name,
            headers: [...sheet.headers],
            data: sheet.data.map(row => [...row])
        }));
    },

    // Mark document as edited
    markAsEdited() {
        const state = AppState;
        state.hasUnsavedChanges = true;
        DOM.editIndicator.style.display = 'flex';
        
        if (state.autoSaveEnabled) {
            FileHandler.triggerAutoSave();
        }
    },

    // Undo all changes
    async undoAllChanges() {
        const state = AppState;
        if (!state.hasUnsavedChanges || state.originalSheetData.length === 0) return;
        
        const confirmed = await Utils.confirm('Undo all changes and restore original data?', {
            title: 'Undo All Changes',
            icon: '↩️',
            okText: 'Undo All',
            cancelText: 'Cancel',
            danger: true
        });
        if (!confirmed) return;
        
        state.allSheets = state.originalSheetData.map(sheet => ({
            name: sheet.name,
            headers: [...sheet.headers],
            data: sheet.data.map(row => [...row])
        }));
        
        state.headers = [...state.allSheets[state.currentSheetIndex].headers];
        state.currentData = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
        
        this.renderTable(state.currentData);
        this.updateRowColCount();
        
        state.hasUnsavedChanges = false;
        DOM.editIndicator.style.display = 'none';
        state.sqlTablesRegistered = false;
        
        Utils.showToast('All changes undone', 'success');
    },

    // Reset sort state
    resetSort() {
        AppState.sortColumn = -1;
        AppState.sortDirection = 'none';
    },

    // Sort and render table
    sortAndRender() {
        const state = AppState;
        let dataToSort;
        
        if (state.hasSqlResults) {
            dataToSort = state.sqlResultData.data.map(row => [...row]);
        } else if (state.activeFilters.length > 0) {
            dataToSort = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
            state.activeFilters.forEach(filter => {
                dataToSort = dataToSort.filter(row => {
                    const cellValue = (row[filter.columnIndex] || '').toString();
                    return Filter.evaluate(cellValue, filter.operator, filter.value);
                });
            });
        } else {
            dataToSort = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
        }
        
        if (state.sortColumn >= 0 && state.sortDirection !== 'none') {
            dataToSort.sort((a, b) => {
                let valA = a[state.sortColumn] || '';
                let valB = b[state.sortColumn] || '';
                
                const numA = parseFloat(valA);
                const numB = parseFloat(valB);
                
                if (!isNaN(numA) && !isNaN(numB)) {
                    return state.sortDirection === 'asc' ? numA - numB : numB - numA;
                }
                
                valA = valA.toString().toLowerCase();
                valB = valB.toString().toLowerCase();
                
                if (valA < valB) return state.sortDirection === 'asc' ? -1 : 1;
                if (valA > valB) return state.sortDirection === 'asc' ? 1 : -1;
                return 0;
            });
        }
        
        state.currentData = dataToSort;
        const isFiltered = state.hasSqlResults || state.activeFilters.length > 0;
        this.renderTable(state.currentData, isFiltered);
        
        const total = state.allSheets[state.currentSheetIndex]?.data.length || state.currentData.length;
        let statusText = '';
        
        if (state.hasSqlResults) {
            statusText = `${state.currentData.length} rows × ${state.headers.length} columns (SQL result)`;
        } else if (state.activeFilters.length > 0) {
            statusText = `Showing ${state.currentData.length} of ${total} rows (filtered)`;
        } else {
            statusText = `${state.currentData.length} rows × ${state.headers.length} columns`;
        }
        
        const sortInfo = state.sortColumn >= 0 && state.sortDirection !== 'none' 
            ? ` • Sorted by ${state.headers[state.sortColumn]} ${state.sortDirection === 'asc' ? '↑' : '↓'}` 
            : '';
        DOM.rowCount.textContent = statusText + sortInfo;
    },

    // Load a specific sheet
    loadSheet(index) {
        const state = AppState;
        if (index < 0 || index >= state.allSheets.length) return;
        
        state.currentSheetIndex = index;
        state.headers = [...state.allSheets[index].headers];
        state.currentData = state.allSheets[index].data.map(row => [...row]);
        
        document.querySelectorAll('.sheet-tab').forEach((tab, i) => {
            tab.classList.toggle('active', i === index);
        });
        
        this.resetSort();
        state.activeFilters = [];
        Filter.updateActiveFiltersUI();
        
        this.updateRowColCount();
        
        DOM.searchInput.value = '';
        DOM.sqlInput.value = '';
        DOM.sqlStatus.className = 'sql-status';
        state.hasSqlResults = false;
        state.sqlResultData = null;
        SQL.hideResultBar();
        SQL.updateTablesHint();
        SQL.updatePlaceholder();
        
        this.renderTable(state.currentData);
        
        DOM.editIndicator.style.display = state.hasUnsavedChanges ? 'flex' : 'none';
    },

    // Render sheet tabs
    renderSheetTabs() {
        const state = AppState;
        DOM.sheetTabs.innerHTML = '';
        DOM.sheetTabs.style.display = 'flex';
        
        state.allSheets.forEach((sheet, index) => {
            const tab = document.createElement('button');
            tab.className = 'sheet-tab' + (index === state.currentSheetIndex ? ' active' : '');
            tab.textContent = sheet.name;
            tab.addEventListener('click', () => this.loadSheet(index));
            DOM.sheetTabs.appendChild(tab);
        });
    },

    // Display loaded data
    displayData(file) {
        const state = AppState;
        state.currentMode = 'spreadsheet';
        
        const canSaveBack = state.fileHandle !== null;
        DOM.fileName.textContent = '📄 ' + file.name + (canSaveBack ? ' ✓' : '');
        DOM.fileName.title = canSaveBack ? 'File can be saved back to original' : 'File will be downloaded as new file when saving';
        DOM.fileSize.textContent = Utils.formatFileSize(file.size);
        DOM.rowCount.textContent = `${state.currentData.length} rows × ${state.headers.length} columns`;
        
        DOM.fileInfo.style.display = 'flex';
        DOM.docInfo.style.display = 'none';
        DOM.sqlSection.style.display = 'block';
        DOM.tableContainer.style.display = 'block';
        DOM.documentContainer.style.display = 'none';
        DOM.uploadSection.style.display = 'none';
        DOM.clearBtn.style.display = 'inline-block';
        
        state.hasUnsavedChanges = false;
        state.originalSheetData = [];
        state.autoSaveEnabled = false;
        this.resetSort();
        DOM.autosaveCheckbox.checked = false;
        DOM.editIndicator.style.display = 'none';
        DOM.saveStatus.style.display = 'none';
        
        DOM.autosaveToggle.style.display = canSaveBack ? 'flex' : 'none';
        DOM.autosaveToggle.title = canSaveBack 
            ? 'Auto-save changes to original file' 
            : 'Open file using "Browse Files" button to enable auto-save';
        
        SQL.registerTables();
        SQL.updatePlaceholder();
        
        if (DOM.quickFilterPanel.style.display !== 'none') {
            Filter.populateColumns();
        }

        this.renderTable(state.currentData);
    },

    // Clear data silently (no confirmation)
    clearSilent() {
        const state = AppState;
        
        state.currentData = [];
        state.headers = [];
        state.allSheets = [];
        state.currentSheetIndex = 0;
        state.currentFile = null;
        state.fileHandle = null;
        state.sqlTablesRegistered = false;
        state.hasSqlResults = false;
        state.sqlResultData = null;
        state.hasUnsavedChanges = false;
        state.originalSheetData = [];
        state.autoSaveEnabled = false;
        state.activeFilters = [];
        this.resetSort();
        
        state.docHasUnsavedChanges = false;
        state.originalDocContent = '';
        state.docFileHandle = null;
        
        state.slidesHasUnsavedChanges = false;
        state.slidesData = [];
        state.currentSlideIndex = 0;
        state.slidesFileHandle = null;
        state.slidesFile = null;
        
        DOM.autosaveCheckbox.checked = false;
        DOM.quickFilterPanel.style.display = 'none';
        DOM.filterToggleBtn.classList.remove('active');
        DOM.activeFiltersDiv.style.display = 'none';
        DOM.fileInput.value = '';
        DOM.searchInput.value = '';
        DOM.sqlInput.value = '';
        
        DOM.editIndicator.style.display = 'none';
        DOM.docEditIndicator.style.display = 'none';
        if (DOM.slidesEditIndicator) DOM.slidesEditIndicator.style.display = 'none';
        DOM.saveStatus.style.display = 'none';
        DOM.sqlStatus.className = 'sql-status';
        DOM.sqlExamplesPanel.style.display = 'none';
        SQL.hideResultBar();
        DOM.tableContainer.innerHTML = '';
        DOM.sheetTabs.style.display = 'none';
        DOM.sheetTabs.innerHTML = '';
        DOM.docInfo.style.display = 'none';
        DOM.documentContainer.style.display = 'none';
        if (DOM.slidesInfo) DOM.slidesInfo.style.display = 'none';
        if (DOM.slidesContainer) DOM.slidesContainer.style.display = 'none';
    },

    // Clear data with confirmation
    async clear() {
        const state = AppState;
        
        if (state.hasUnsavedChanges || state.docHasUnsavedChanges || state.slidesHasUnsavedChanges) {
            const confirmed = await Utils.confirm('You have unsaved changes. Are you sure you want to clear?', {
                title: 'Unsaved Changes',
                icon: '⚠️',
                okText: 'Clear',
                cancelText: 'Cancel',
                danger: true
            });
            if (!confirmed) return;
        }
        
        this.clearSilent();
        state.currentMode = 'none';
        DOM.docEditor.innerHTML = '<p></p>';
        
        DOM.fileInfo.style.display = 'none';
        DOM.docInfo.style.display = 'none';
        if (DOM.slidesInfo) DOM.slidesInfo.style.display = 'none';
        DOM.sqlSection.style.display = 'none';
        DOM.tableContainer.style.display = 'none';
        DOM.documentContainer.style.display = 'none';
        if (DOM.slidesContainer) DOM.slidesContainer.style.display = 'none';
        DOM.uploadSection.style.display = 'flex';
        DOM.clearBtn.style.display = 'none';
    }
};

// Export for use in other modules
window.Spreadsheet = Spreadsheet;