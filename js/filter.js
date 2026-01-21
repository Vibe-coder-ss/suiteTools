/**
 * Filter Module
 * Handles quick filtering and search functionality
 */

const Filter = {
    // Toggle filter panel
    togglePanel() {
        const isVisible = DOM.quickFilterPanel.style.display !== 'none';
        DOM.quickFilterPanel.style.display = isVisible ? 'none' : 'block';
        DOM.filterToggleBtn.classList.toggle('active', !isVisible);
        
        if (!isVisible) {
            this.populateColumns();
        }
    },

    // Populate filter column dropdown
    populateColumns() {
        const state = AppState;
        DOM.filterColumn.innerHTML = '<option value="">Select Column...</option>';
        state.headers.forEach((header, index) => {
            const cleanName = header || `Column ${index + 1}`;
            DOM.filterColumn.innerHTML += `<option value="${index}">${cleanName}</option>`;
        });
    },

    // Update filter value visibility based on operator
    updateValueVisibility() {
        const operator = DOM.filterOperator.value;
        const valueGroup = document.querySelector('.filter-value-group');
        
        if (operator === 'is_empty' || operator === 'is_not_empty') {
            valueGroup.style.display = 'none';
        } else {
            valueGroup.style.display = 'block';
        }
    },

    // Apply quick filter
    apply() {
        const state = AppState;
        const colIndex = parseInt(DOM.filterColumn.value);
        const operator = DOM.filterOperator.value;
        const value = DOM.filterValue.value.trim();
        
        if (isNaN(colIndex)) {
            Utils.showToast('Please select a column', 'warning');
            return;
        }
        
        if (!['is_empty', 'is_not_empty'].includes(operator) && !value) {
            Utils.showToast('Please enter a filter value', 'warning');
            return;
        }
        
        const filter = {
            id: Date.now(),
            columnIndex: colIndex,
            columnName: state.headers[colIndex] || `Column ${colIndex + 1}`,
            operator: operator,
            value: value,
            operatorLabel: this.getOperatorLabel(operator)
        };
        
        state.activeFilters.push(filter);
        this.applyAll();
        this.updateActiveFiltersUI();
        
        DOM.filterValue.value = '';
        DOM.filterColumn.value = '';
        
        Utils.showToast('Filter applied', 'success');
    },

    // Get operator label
    getOperatorLabel(operator) {
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
    },

    // Apply all active filters
    applyAll() {
        const state = AppState;
        
        let filteredData = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
        
        state.activeFilters.forEach(filter => {
            filteredData = filteredData.filter(row => {
                const cellValue = (row[filter.columnIndex] || '').toString();
                return this.evaluate(cellValue, filter.operator, filter.value);
            });
        });
        
        state.currentData = filteredData;
        state.hasSqlResults = false;
        SQL.hideResultBar();
        
        if (state.sortColumn >= 0 && state.sortDirection !== 'none') {
            Spreadsheet.sortAndRender();
        } else {
            Spreadsheet.renderTable(state.currentData, state.activeFilters.length > 0);
        }
        
        const total = state.allSheets[state.currentSheetIndex].data.length;
        if (state.activeFilters.length > 0) {
            DOM.rowCount.textContent = `Showing ${state.currentData.length} of ${total} rows (filtered)`;
        } else {
            DOM.rowCount.textContent = `${state.currentData.length} rows × ${state.headers.length} columns`;
        }
    },

    // Evaluate filter condition
    evaluate(cellValue, operator, filterValue) {
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
    },

    // Update active filters UI
    updateActiveFiltersUI() {
        const state = AppState;
        
        if (state.activeFilters.length === 0) {
            DOM.activeFiltersDiv.style.display = 'none';
            return;
        }
        
        DOM.activeFiltersDiv.style.display = 'flex';
        DOM.filterTags.innerHTML = state.activeFilters.map(filter => {
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
        
        DOM.filterTags.querySelectorAll('.filter-tag-remove').forEach(btn => {
            btn.addEventListener('click', () => {
                this.remove(parseInt(btn.dataset.filterId));
            });
        });
    },

    // Remove a specific filter
    remove(filterId) {
        const state = AppState;
        state.activeFilters = state.activeFilters.filter(f => f.id !== filterId);
        this.applyAll();
        this.updateActiveFiltersUI();
        Utils.showToast('Filter removed', 'info');
    },

    // Clear all filters
    clearAll() {
        const state = AppState;
        state.activeFilters = [];
        DOM.filterColumn.value = '';
        DOM.filterValue.value = '';
        DOM.filterOperator.value = 'equals';
        
        state.currentData = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
        
        if (state.sortColumn >= 0 && state.sortDirection !== 'none') {
            Spreadsheet.sortAndRender();
        } else {
            Spreadsheet.renderTable(state.currentData);
        }
        
        this.updateActiveFiltersUI();
        DOM.rowCount.textContent = `${state.currentData.length} rows × ${state.headers.length} columns`;
        
        Utils.showToast('All filters cleared', 'info');
    },

    // Handle search
    handleSearch() {
        const state = AppState;
        const query = DOM.searchInput.value.toLowerCase().trim();
        
        let dataToSearch = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
        
        if (state.activeFilters.length > 0) {
            state.activeFilters.forEach(filter => {
                dataToSearch = dataToSearch.filter(row => {
                    const cellValue = (row[filter.columnIndex] || '').toString();
                    return this.evaluate(cellValue, filter.operator, filter.value);
                });
            });
        }
        
        Spreadsheet.resetSort();
        
        if (!query) {
            state.currentData = dataToSearch;
            Spreadsheet.renderTable(state.currentData, state.activeFilters.length > 0);
            if (state.activeFilters.length > 0) {
                DOM.rowCount.textContent = `Showing ${state.currentData.length} of ${state.allSheets[state.currentSheetIndex].data.length} rows (filtered)`;
            } else {
                DOM.rowCount.textContent = `${state.currentData.length} rows × ${state.headers.length} columns`;
            }
            return;
        }

        const searchResults = dataToSearch.filter(row => 
            row.some(cell => cell.toLowerCase().includes(query))
        );

        state.currentData = searchResults;
        
        Spreadsheet.renderTable(searchResults, true);
        
        const baseCount = state.activeFilters.length > 0 ? dataToSearch.length : state.allSheets[state.currentSheetIndex].data.length;
        DOM.rowCount.textContent = `Showing ${searchResults.length} of ${baseCount} rows (search${state.activeFilters.length > 0 ? ' + filter' : ''})`;
    }
};

// Export for use in other modules
window.Filter = Filter;