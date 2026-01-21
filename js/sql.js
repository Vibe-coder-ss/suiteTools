/**
 * SQL Module
 * Handles SQL query functionality
 */

const SQL = {
    // Get clean table name from sheet name
    getTableName(sheetName) {
        let name = sheetName.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^[0-9]/, '_$&');
        return name || 'sheet';
    },

    // Get clean column name
    getColumnName(header, index) {
        if (!header) return `col_${index}`;
        return header.replace(/[^a-zA-Z0-9_]/g, '_').replace(/^[0-9]/, '_$&') || `col_${index}`;
    },

    // Register all sheets as SQL tables
    registerTables() {
        const state = AppState;
        
        // Clear existing tables
        state.allSheets.forEach(sheet => {
            const tableName = this.getTableName(sheet.name);
            try {
                alasql(`DROP TABLE IF EXISTS ${tableName}`);
            } catch (e) {}
        });
        try {
            alasql('DROP TABLE IF EXISTS data');
        } catch (e) {}
        
        // Register each sheet as a table
        state.allSheets.forEach((sheet, index) => {
            const tableName = this.getTableName(sheet.name);
            
            const tableData = sheet.data.map(row => {
                const obj = {};
                sheet.headers.forEach((h, i) => {
                    const cleanHeader = this.getColumnName(h, i);
                    let value = row[i] || '';
                    if (value !== '' && !isNaN(value) && !isNaN(parseFloat(value))) {
                        value = parseFloat(value);
                    }
                    obj[cleanHeader] = value;
                });
                return obj;
            });
            
            alasql(`CREATE TABLE ${tableName}`);
            alasql.tables[tableName].data = tableData;
            
            if (index === state.currentSheetIndex) {
                alasql('CREATE TABLE data');
                alasql.tables.data.data = [...tableData];
            }
        });
        
        state.sqlTablesRegistered = true;
        this.updateTablesHint();
    },

    // Update SQL tables hint in header
    updateTablesHint() {
        const state = AppState;
        const hint = DOM.sqlTablesHint;
        if (!hint) return;
        
        if (state.allSheets.length === 1) {
            const tableName = this.getTableName(state.allSheets[0].name);
            hint.innerHTML = `Tables: <code>data</code> <span class="table-separator">or</span> <code>${tableName}</code>`;
        } else if (state.allSheets.length > 1) {
            const tableNames = state.allSheets.map(s => `<code>${this.getTableName(s.name)}</code>`).join(' ');
            hint.innerHTML = `Tables: <code>data</code> (current) ${tableNames}`;
        } else {
            hint.innerHTML = '';
        }
    },

    // Update SQL placeholder with actual column names
    updatePlaceholder() {
        const state = AppState;
        if (state.headers.length > 0) {
            const cols = state.headers.slice(0, 3).map((h, i) => this.getColumnName(h, i)).join(', ');
            const firstCol = this.getColumnName(state.headers[0], 0);
            
            if (state.allSheets.length > 1) {
                const tables = state.allSheets.slice(0, 2).map(s => this.getTableName(s.name)).join(', ');
                DOM.sqlInput.placeholder = `SELECT * FROM data WHERE ${firstCol} = 'value' -- Tables: ${tables}...`;
            } else {
                DOM.sqlInput.placeholder = `SELECT ${cols}... FROM data WHERE ${firstCol} = 'value' ORDER BY ${firstCol}`;
            }
        }
    },

    // Run SQL query
    run() {
        const state = AppState;
        const query = DOM.sqlInput.value.trim();
        
        if (!query) {
            this.showStatus('Please enter a SQL query', 'error');
            return;
        }
        
        try {
            const queryType = this.detectQueryType(query);
            
            if (!state.sqlTablesRegistered) {
                this.registerTables();
            }
            
            // Re-register "data" table to point to current sheet
            const currentSheet = state.allSheets[state.currentSheetIndex];
            const currentTableData = currentSheet.data.map(row => {
                const obj = {};
                currentSheet.headers.forEach((h, i) => {
                    const cleanHeader = this.getColumnName(h, i);
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
            
            const startTime = performance.now();
            const result = alasql(query);
            const endTime = performance.now();
            const execTime = (endTime - startTime).toFixed(2);
            
            if (queryType === 'SELECT') {
                this.handleSelectResult(result, query, execTime, currentSheet);
            } else if (queryType === 'UPDATE' || queryType === 'DELETE' || queryType === 'INSERT') {
                this.handleModifyResult(result, query, execTime, queryType);
            } else {
                state.hasSqlResults = false;
                this.hideResultBar();
                this.showStatus(`✓ Query executed in ${execTime}ms — Result: ${JSON.stringify(result)}`, 'success');
            }
            
        } catch (error) {
            this.showStatus(`✗ Error: ${error.message}`, 'error');
            console.error('SQL Error:', error);
        }
    },

    // Detect SQL query type
    detectQueryType(query) {
        const upperQuery = query.trim().toUpperCase();
        if (upperQuery.startsWith('SELECT')) return 'SELECT';
        if (upperQuery.startsWith('UPDATE')) return 'UPDATE';
        if (upperQuery.startsWith('DELETE')) return 'DELETE';
        if (upperQuery.startsWith('INSERT')) return 'INSERT';
        return 'OTHER';
    },

    // Handle SELECT query result
    handleSelectResult(result, query, execTime, currentSheet) {
        const state = AppState;
        
        if (Array.isArray(result)) {
            if (result.length > 0) {
                state.headers = Object.keys(result[0]);
                state.currentData = result.map(row => state.headers.map(h => {
                    const val = row[h];
                    return val !== null && val !== undefined ? val.toString() : '';
                }));
            } else {
                state.headers = currentSheet.headers.map((h, i) => this.getColumnName(h, i));
                state.currentData = [];
            }
            
            state.sqlResultData = {
                headers: [...state.headers],
                data: state.currentData.map(row => [...row]),
                query: query,
                rowCount: result.length
            };
            state.hasSqlResults = true;
            
            this.showResultBar(result.length, state.headers.length, execTime);
            
            Spreadsheet.renderTable(state.currentData, true);
            this.showStatus(`✓ Query executed in ${execTime}ms — ${result.length} rows returned`, 'success');
            DOM.rowCount.textContent = `${result.length} rows × ${state.headers.length} columns (SQL result)`;
        }
    },

    // Handle UPDATE, DELETE, INSERT query result
    handleModifyResult(result, query, execTime, queryType) {
        const state = AppState;
        
        if (!state.hasUnsavedChanges) {
            Spreadsheet.storeOriginalData();
        }
        
        const tableName = this.extractTableName(query, queryType);
        const sheetIndex = this.findSheetIndexByTableName(tableName);
        
        if (sheetIndex === -1) {
            this.showStatus(`✗ Error: Table "${tableName}" not found`, 'error');
            return;
        }
        
        const modifiedData = alasql.tables[tableName]?.data || [];
        
        const sheet = state.allSheets[sheetIndex];
        const newData = modifiedData.map(row => {
            return sheet.headers.map((h, i) => {
                const cleanHeader = this.getColumnName(h, i);
                const val = row[cleanHeader];
                return val !== null && val !== undefined ? val.toString() : '';
            });
        });
        
        state.allSheets[sheetIndex].data = newData;
        
        if (sheetIndex === state.currentSheetIndex) {
            state.headers = [...sheet.headers];
            state.currentData = newData.map(row => [...row]);
            Spreadsheet.renderTable(state.currentData);
            DOM.rowCount.textContent = `${state.currentData.length} rows × ${state.headers.length} columns`;
        }
        
        Spreadsheet.markAsEdited();
        state.sqlTablesRegistered = false;
        state.hasSqlResults = false;
        this.hideResultBar();
        
        let affectedRows = typeof result === 'number' ? result : (Array.isArray(result) ? result.length : 1);
        const actionText = queryType === 'UPDATE' ? 'updated' : (queryType === 'DELETE' ? 'deleted' : 'inserted');
        this.showStatus(`✓ ${queryType} executed in ${execTime}ms — ${affectedRows} row(s) ${actionText}`, 'success');
        
        Utils.showToast(`${affectedRows} row(s) ${actionText}`, 'success');
    },

    // Extract table name from query
    extractTableName(query, queryType) {
        let tableName = 'data';
        
        try {
            if (queryType === 'UPDATE') {
                const match = query.match(/UPDATE\s+(\w+)/i);
                if (match) tableName = match[1];
            } else if (queryType === 'DELETE') {
                const match = query.match(/DELETE\s+FROM\s+(\w+)/i);
                if (match) tableName = match[1];
            } else if (queryType === 'INSERT') {
                const match = query.match(/INSERT\s+INTO\s+(\w+)/i);
                if (match) tableName = match[1];
            }
        } catch (e) {
            console.error('Error extracting table name:', e);
        }
        
        return tableName;
    },

    // Find sheet index by table name
    findSheetIndexByTableName(tableName) {
        const state = AppState;
        
        if (tableName.toLowerCase() === 'data') {
            return state.currentSheetIndex;
        }
        
        for (let i = 0; i < state.allSheets.length; i++) {
            if (this.getTableName(state.allSheets[i].name).toLowerCase() === tableName.toLowerCase()) {
                return i;
            }
        }
        
        return -1;
    },

    // Show SQL status message
    showStatus(message, type) {
        DOM.sqlStatus.textContent = message;
        DOM.sqlStatus.className = 'sql-status ' + type;
        
        if (type === 'success') {
            setTimeout(() => {
                DOM.sqlStatus.className = 'sql-status';
            }, 5000);
        }
    },

    // Reset to original data
    reset() {
        const state = AppState;
        
        state.headers = [...state.allSheets[state.currentSheetIndex].headers];
        state.currentData = state.allSheets[state.currentSheetIndex].data.map(row => [...row]);
        
        DOM.sqlInput.value = '';
        DOM.sqlStatus.className = 'sql-status';
        DOM.sqlLineCount.textContent = '';
        DOM.searchInput.value = '';
        
        state.hasSqlResults = false;
        state.sqlResultData = null;
        this.hideResultBar();
        
        DOM.sqlInput.style.height = '';
        this.autoResizeEditor();
        
        Spreadsheet.renderTable(state.currentData);
        DOM.rowCount.textContent = `${state.currentData.length} rows × ${state.headers.length} columns`;
    },

    // Show SQL result bar
    showResultBar(rows, cols, execTime) {
        const resultText = document.getElementById('sqlResultText');
        resultText.textContent = `Query returned ${rows} rows × ${cols} columns in ${execTime}ms`;
        DOM.sqlResultBar.style.display = 'flex';
    },

    // Hide SQL result bar
    hideResultBar() {
        if (DOM.sqlResultBar) {
            DOM.sqlResultBar.style.display = 'none';
        }
    },

    // Auto-resize SQL editor
    autoResizeEditor() {
        if (DOM.sqlEditorWrapper.classList.contains('expanded')) return;
        
        DOM.sqlInput.style.height = 'auto';
        
        const minHeight = 60;
        const maxHeight = 300;
        const newHeight = Math.min(Math.max(DOM.sqlInput.scrollHeight, minHeight), maxHeight);
        
        DOM.sqlInput.style.height = newHeight + 'px';
    },

    // Update line count
    updateLineCount() {
        const text = DOM.sqlInput.value;
        if (!text) {
            DOM.sqlLineCount.textContent = '';
            return;
        }
        
        const lines = text.split('\n').length;
        const chars = text.length;
        DOM.sqlLineCount.textContent = `${lines} line${lines !== 1 ? 's' : ''} • ${chars} char${chars !== 1 ? 's' : ''}`;
    },

    // Toggle expanded editor
    toggleExpand() {
        const isExpanded = DOM.sqlEditorWrapper.classList.toggle('expanded');
        
        if (isExpanded) {
            DOM.sqlInput.dataset.originalHeight = DOM.sqlInput.style.height;
            DOM.sqlInput.style.height = '';
            document.body.style.overflow = 'hidden';
        } else {
            DOM.sqlInput.style.height = DOM.sqlInput.dataset.originalHeight || '';
            document.body.style.overflow = '';
            this.autoResizeEditor();
        }
        
        DOM.sqlInput.focus();
    },

    // Toggle examples panel
    toggleExamplesPanel() {
        const isVisible = DOM.sqlExamplesPanel.style.display !== 'none';
        DOM.sqlExamplesPanel.style.display = isVisible ? 'none' : 'block';
        
        if (!isVisible) {
            this.updateExamplesPanel();
        }
    },

    // Update examples panel
    updateExamplesPanel() {
        const state = AppState;
        const tablesListPanel = document.getElementById('tablesListPanel');
        const basicExamples = document.getElementById('basicExamples');
        const joinExamples = document.getElementById('joinExamples');
        const joinExamplesSection = document.getElementById('joinExamplesSection');
        
        // Build tables list
        tablesListPanel.innerHTML = state.allSheets.map((sheet, index) => {
            const tableName = this.getTableName(sheet.name);
            const isActive = index === state.currentSheetIndex;
            return `
                <div class="table-badge ${isActive ? 'active' : ''}" data-table="${tableName}">
                    <code>${tableName}</code>
                    <span class="table-info">${sheet.data.length} rows • ${sheet.headers.length} cols</span>
                </div>
            `;
        }).join('');
        
        const currentTable = this.getTableName(state.allSheets[state.currentSheetIndex].name);
        const currentCols = state.allSheets[state.currentSheetIndex].headers;
        const col1 = currentCols[0] ? this.getColumnName(currentCols[0], 0) : 'column1';
        const col2 = currentCols[1] ? this.getColumnName(currentCols[1], 1) : 'column2';
        
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
        
        if (state.allSheets.length > 1) {
            joinExamplesSection.style.display = 'block';
            
            const table1 = this.getTableName(state.allSheets[0].name);
            const table2 = this.getTableName(state.allSheets[1].name);
            const t1Col = state.allSheets[0].headers[0] ? this.getColumnName(state.allSheets[0].headers[0], 0) : 'id';
            const t2Col = state.allSheets[1].headers[0] ? this.getColumnName(state.allSheets[1].headers[0], 0) : 'id';
            
            joinExamples.innerHTML = `
                <div class="example-item" data-query="SELECT * FROM ${table1} a JOIN ${table2} b ON a.${t1Col} = b.${t2Col}">
                    <strong>Inner Join</strong>
                    <code>SELECT * FROM ${table1} a JOIN ${table2} b ON a.${t1Col} = b.${t2Col}</code>
                </div>
                <div class="example-item" data-query="SELECT * FROM ${table1} a LEFT JOIN ${table2} b ON a.${t1Col} = b.${t2Col}">
                    <strong>Left Join</strong>
                    <code>SELECT * FROM ${table1} a LEFT JOIN ${table2} b ON a.${t1Col} = b.${t2Col}</code>
                </div>
            `;
        } else {
            joinExamplesSection.style.display = 'none';
        }
    }
};

// Export for use in other modules
window.SQL = SQL;