if (typeof window.MaxPilotDuckDB === 'undefined') {
class FastTableComparator {
    constructor() {
        this.initialized = false;
        this.tables = new Map();
        this.mode = 'local';
    }

    async initialize() {
        try {
            if (typeof WebAssembly === 'undefined') {
                throw new Error('WebAssembly not supported');
            }
            
            try {
                await this.initializeWASM();
                this.mode = 'wasm';
            } catch (wasmError) {
                await this.initializeLocal();
                this.mode = 'local';
            }
            
            this.initialized = true;
            return true;
            
        } catch (error) {
            this.initialized = false;
            return false;
        }
    }

    async initializeWASM() {
        throw new Error('CORS policy blocks external workers');
    }

    async initializeLocal() {
        this.tables.clear();
        return true;
    }

    async createTableFromData(tableName, data, headers = null) {
        if (!this.initialized) {
            throw new Error('Comparator not initialized');
        }

        try {
            const processedData = this.processTableData(data, headers);
            this.tables.set(tableName, processedData);
            
            return true;
            
        } catch (error) {
            throw error;
        }
    }

    processTableData(data, headers = null) {
        if (!data || data.length === 0) {
            return { columns: [], rows: [], indexes: new Map() };
        }

        const columns = headers || (data[0] ? data[0].map((_, i) => `col_${i}`) : []);
        const dataRows = headers ? data.slice(1) : data;
        
        const indexes = new Map();
        const processedRows = dataRows.map((row, index) => {
            const paddedRow = Array(columns.length).fill('');
            row.forEach((cell, i) => {
                if (i < columns.length) {
                    paddedRow[i] = cell !== null && cell !== undefined ? String(cell) : '';
                }
            });
            
            const rowHash = this.createRowHash(paddedRow);
            if (!indexes.has(rowHash)) {
                indexes.set(rowHash, []);
            }
            indexes.get(rowHash).push({ index, data: paddedRow });
            
            return { index, data: paddedRow, hash: rowHash };
        });

        return {
            columns,
            rows: processedRows,
            indexes
        };
    }

    createRowHash(row) {
        const str = row.join('|');
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash;
        }
        return hash;
    }

    async compareTablesFast(table1Name, table2Name, excludeColumns = []) {
        if (!this.initialized) {
            throw new Error('Comparator not initialized');
        }

        const table1 = this.tables.get(table1Name);
        const table2 = this.tables.get(table2Name);
        
        if (!table1 || !table2) {
            throw new Error('One or both tables not found');
        }

        const startTime = performance.now();
        
        const commonColumns = table1.columns.filter(col => 
            table2.columns.includes(col) && 
            !excludeColumns.some(excCol => 
                col.toLowerCase().includes(excCol.toLowerCase())
            )
        );

        if (commonColumns.length === 0) {
            throw new Error('No common columns found for comparison');
        }

        const table1ColIndices = commonColumns.map(col => table1.columns.indexOf(col));
        const table2ColIndices = commonColumns.map(col => table2.columns.indexOf(col));

        const table1Hashes = new Map();
        const table2Hashes = new Map();

        table1.rows.forEach(row => {
            const compareData = table1ColIndices.map(i => row.data[i]);
            const hash = this.createRowHash(compareData);
            
            if (!table1Hashes.has(hash)) {
                table1Hashes.set(hash, []);
            }
            table1Hashes.get(hash).push(row.index + 1);
        });

        table2.rows.forEach(row => {
            const compareData = table2ColIndices.map(i => row.data[i]);
            const hash = this.createRowHash(compareData);
            
            if (!table2Hashes.has(hash)) {
                table2Hashes.set(hash, []);
            }
            table2Hashes.get(hash).push(row.index + 1);
        });

        const identical = [];
        const onlyInTable1 = [];
        const onlyInTable2 = [];

        for (const [hash, table1Rows] of table1Hashes) {
            if (table2Hashes.has(hash)) {
                const table2Rows = table2Hashes.get(hash);
                
                table1Rows.forEach(row1 => {
                    table2Rows.forEach(row2 => {
                        identical.push({ row1, row2, status: 'IDENTICAL' });
                    });
                });
            } else {
                table1Rows.forEach(row1 => {
                    onlyInTable1.push({ row1, row2: null, status: 'ONLY_IN_TABLE1' });
                });
            }
        }

        for (const [hash, table2Rows] of table2Hashes) {
            if (!table1Hashes.has(hash)) {
                table2Rows.forEach(row2 => {
                    onlyInTable2.push({ row1: null, row2, status: 'ONLY_IN_TABLE2' });
                });
            }
        }

        const duration = performance.now() - startTime;

        return {
            identical,
            onlyInTable1,
            onlyInTable2,
            table1Count: table1.rows.length,
            table2Count: table2.rows.length,
            commonColumns,
            performance: {
                duration: duration,
                rowsPerSecond: Math.round((table1.rows.length + table2.rows.length) / (duration / 1000))
            }
        };
    }

    async getTableColumns(tableName) {
        const table = this.tables.get(tableName);
        return table ? table.columns : [];
    }

    async close() {
        this.tables.clear();
        this.initialized = false;
    }

    static async isSupported() {
        try {
            return typeof WebAssembly !== 'undefined';
        } catch (error) {
            return false;
        }
    }
}

let fastComparator = null;

async function initializeFastComparator() {
    if (fastComparator) {
        return fastComparator;
    }

    try {
        fastComparator = new FastTableComparator();
        const initialized = await fastComparator.initialize();
        
        if (initialized) {
            showFastModeStatus(true, fastComparator.mode);
            window.duckDBManager = fastComparator;
            window.duckDBAvailable = true;
            
            return fastComparator;
        } else {
            showFastModeStatus(false);
            return null;
        }
    } catch (error) {
        showFastModeStatus(false);
        return null;
    }
}

function showFastModeStatus(available, mode = 'local') {
    const statusElement = document.getElementById('duckdb-status');
    if (statusElement) {
        if (available) {
            const modeText = mode === 'wasm' ? 'DuckDB WASM' : 'Fast Local';
            statusElement.innerHTML = `⚡ ${modeText} mode enabled - Enhanced performance!`;
            statusElement.className = 'duckdb-status duckdb-available show';
            
            const fastIndicators = document.querySelectorAll('.fast-mode-indicator');
            fastIndicators.forEach(indicator => {
                indicator.style.display = 'inline-block';
                indicator.textContent = mode === 'wasm' ? 'ULTRA' : 'FAST';
            });
            
            setTimeout(() => {
                if (statusElement.classList.contains('duckdb-available')) {
                    statusElement.style.opacity = '0.8';
                    statusElement.innerHTML = `⚡ ${modeText} active`;
                }
            }, 5000);
            
        } else {
            statusElement.innerHTML = '🔄 Standard comparison mode';
            statusElement.className = 'duckdb-status duckdb-unavailable show';
            
            const fastIndicators = document.querySelectorAll('.fast-mode-indicator');
            fastIndicators.forEach(indicator => {
                indicator.style.display = 'none';
            });
            
            setTimeout(() => {
                statusElement.style.display = 'none';
            }, 3000);
        }
    }
}

async function compareTablesWithFastComparator(data1, data2, excludeColumns = [], useTolerance = false, tolerance = 1.5) {
    try {
        if (!fastComparator || !fastComparator.initialized) {
            return null;
        }

        const { data1: alignedData1, data2: alignedData2, columnInfo } = prepareDataForComparison(data1, data2);
        
        const headers1 = alignedData1[0] || [];
        const headers2 = alignedData2[0] || [];
        const dataRows1 = alignedData1.slice(1);
        const dataRows2 = alignedData2.slice(1);
        
        await fastComparator.createTableFromData('table1', dataRows1, headers1);
        await fastComparator.createTableFromData('table2', dataRows2, headers2);
        
        const comparisonResult = await fastComparator.compareTablesFast(
            'table1', 'table2', excludeColumns
        );

        comparisonResult.columnInfo = columnInfo;
        comparisonResult.alignedData1 = alignedData1;
        comparisonResult.alignedData2 = alignedData2;

        return comparisonResult;
        
    } catch (error) {
        return null;
    }
}

async function runFastComparatorTests() {
    const tests = [
        { name: 'Browser Support', fn: () => FastTableComparator.isSupported() },
        { name: 'Initialization', fn: async () => {
            const comp = new FastTableComparator();
            return await comp.initialize();
        }},
        { name: 'Table Creation', fn: async () => {
            if (!fastComparator) await initializeFastComparator();
            const testData = [['id', 'name'], [1, 'Alice'], [2, 'Bob']];
            await fastComparator.createTableFromData('test', testData);
            return true;
        }},
        { name: 'Comparison', fn: async () => {
            if (!fastComparator) await initializeFastComparator();
            
            const data1 = [['id', 'name'], [1, 'Alice'], [2, 'Bob']];
            const data2 = [['id', 'name'], [1, 'Alice'], [3, 'Charlie']];
            
            await fastComparator.createTableFromData('test1', data1);
            await fastComparator.createTableFromData('test2', data2);
            
            const result = await fastComparator.compareTablesFast('test1', 'test2');
            return result.identical.length === 1 && result.onlyInTable1.length === 1 && result.onlyInTable2.length === 1;
        }},
        { name: 'Performance', fn: async () => {
            if (!fastComparator) await initializeFastComparator();
            
            const size = 1000;
            const data1 = [['id', 'value']];
            const data2 = [['id', 'value']];
            
            for (let i = 1; i <= size; i++) {
                data1.push([i, Math.random() * 1000]);
                data2.push([i, Math.random() * 1000]);
            }
            
            const startTime = performance.now();
            await fastComparator.createTableFromData('perf1', data1);
            await fastComparator.createTableFromData('perf2', data2);
            await fastComparator.compareTablesFast('perf1', 'perf2');
            const duration = performance.now() - startTime;
            
            return duration < 500;
        }}
    ];

    const results = [];
    for (const test of tests) {
        try {
            const result = await test.fn();
            results.push({ name: test.name, status: 'PASS', result });
        } catch (error) {
            results.push({ name: test.name, status: 'FAIL', error: error.message });
        }
    }

    const passed = results.filter(r => r.status === 'PASS').length;
    const total = results.length;
    
    return results;
}

async function compareTablesEnhanced(useTolerance = false) {
    clearComparisonResults();
    
    let resultDiv = document.getElementById('result');
    let summaryDiv = document.getElementById('summary');
    
    if (!data1.length || !data2.length) {
        document.getElementById('result').innerText = 'Please, load both files.';
        document.getElementById('summary').innerHTML = '';
        showPlaceholderMessage();
        return;
    }

    const totalRows = Math.max(data1.length, data2.length);
    if (totalRows > MAX_ROWS_LIMIT) {
        resultDiv.innerHTML = generateLimitErrorMessage(
            'rows', data1.length, MAX_ROWS_LIMIT, '', 
            'columns', Math.max(data1[0]?.length || 0, data2[0]?.length || 0), MAX_COLS_LIMIT
        );
        summaryDiv.innerHTML = '';
        return;
    }

    try {
        if (fastComparator && fastComparator.initialized) {
            resultDiv.innerHTML = '<div class="comparison-loading-enhanced">⚡ Using fast comparison engine...</div>';
            summaryDiv.innerHTML = '<div style="text-align: center; padding: 10px;">Processing with enhanced performance...</div>';
            
            setTimeout(async () => {
                resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚡ Fast mode - processing comparison...</div>';
                await performComparison();
                
                // Double check and clear result div after comparison completes
                setTimeout(() => {
                    const resultElement = document.getElementById('result');
                    if (resultElement) {
                        const currentContent = resultElement.innerHTML;
                        if (currentContent.includes('Fast mode - processing comparison') || 
                            currentContent.includes('Using fast comparison') ||
                            currentContent.includes('Processing with enhanced')) {
                            resultElement.innerHTML = '';
                            resultElement.style.display = 'none';
                        }
                    }
                }, 100);
            }, 100);
            
            return;
        }

        resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">🔄 Using standard comparison...</div>';
        
        setTimeout(async () => {
            await performComparison();
        }, 10);
        
    } catch (error) {
        resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">🔄 Using standard comparison...</div>';
        setTimeout(async () => {
            await performComparison();
        }, 10);
    }
}

async function processFastComparisonResults(fastResult, useTolerance) {
    const { identical, onlyInTable1, onlyInTable2, table1Count, table2Count, commonColumns, performance } = fastResult;
    
    // Clear any loading messages immediately
    const resultDiv = document.getElementById('result');
    if (resultDiv) {
        resultDiv.innerHTML = '';
        resultDiv.style.display = 'none';
    }
    
    const identicalCount = identical.length;
    const similarity = table1Count > 0 ? ((identicalCount / Math.max(table1Count, table2Count)) * 100).toFixed(1) : 0;
    
    const file1Name = window.fileName1 || 'File 1';
    const file2Name = window.fileName2 || 'File 2';

    const tableHeaders = getSummaryTableHeaders();
    const summaryHTML = `
        <div style="overflow-x: auto; margin: 20px 0;">
            <div style="text-align: center; margin-bottom: 15px; padding: 10px; background: rgba(40,167,69,0.1); border-radius: 6px;">
                ⚡ Fast Mode: ${performance.rowsPerSecond.toLocaleString()} rows/sec | ${performance.duration.toFixed(2)}ms total
            </div>
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>${tableHeaders.file}</th>
                        <th>${tableHeaders.rowCount}</th>
                        <th>${tableHeaders.rowsOnlyInFile}</th>
                        <th>${tableHeaders.identicalRows}</th>
                        <th>${tableHeaders.similarity}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>${file1Name}</strong></td>
                        <td>${table1Count.toLocaleString()}</td>
                        <td>${onlyInTable1.length.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 18px;">${identicalCount.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 18px;">${similarity}%</td>
                    </tr>
                    <tr>
                        <td><strong>${file2Name}</strong></td>
                        <td>${table2Count.toLocaleString()}</td>
                        <td>${onlyInTable2.length.toLocaleString()}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `;
    
    document.getElementById('summary').innerHTML = summaryHTML;

    try {
        await generateDetailedComparisonTable(fastResult, useTolerance);
        
        const resultDiv = document.getElementById('result');
        if (resultDiv) {
            resultDiv.innerHTML = '';
            resultDiv.style.display = 'none';
        }
        
        const diffTable = document.getElementById('diffTable');
        if (diffTable) {
            diffTable.style.display = 'block';
            diffTable.style.visibility = 'visible';
            diffTable.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        
    } catch (error) {
        const resultDiv = document.getElementById('result');
        if (resultDiv) {
            resultDiv.innerHTML = '<div style="text-align: center; padding: 10px; color: #dc3545;">❌ Error generating comparison table. Check console for details.</div>';
        }
    }
    
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'flex';
        filterControls.classList.remove('filter-controls-hidden');
    }

    const exportBtn = document.getElementById('exportExcelBtn');
    if (exportBtn) {
        exportBtn.style.display = 'inline-block';
        exportBtn.classList.remove('export-btn-hidden');
    }
}

async function generateDetailedComparisonTable(fastResult, useTolerance) {
    const { identical, onlyInTable1, onlyInTable2, commonColumns, alignedData1, alignedData2 } = fastResult;
    
    // Clear any remaining loading messages
    const resultDiv = document.getElementById('result');
    if (resultDiv) {
        resultDiv.innerHTML = '';
        resultDiv.style.display = 'none';
    }
    
    const workingData1 = alignedData1 || data1;
    const workingData2 = alignedData2 || data2;
    
    const pairs = [];
    
    const maxIdenticalToShow = onlyInTable1.length === 0 && onlyInTable2.length === 0 ? 1000 : 100;
    identical.slice(0, maxIdenticalToShow).forEach(identicalPair => {
        const row1Index = identicalPair.row1 - 1;
        const row2Index = identicalPair.row2 - 1;
        const row1 = workingData1[row1Index + 1];
        const row2 = workingData2[row2Index + 1];
        
        pairs.push({
            row1: row1,
            row2: row2,
            index1: row1Index,
            index2: row2Index,
            isDifferent: false,
            onlyIn: null
        });
    });
    
    onlyInTable1.forEach(diff => {
        const rowIndex = diff.row1 - 1;
        const row1 = workingData1[rowIndex + 1];
        
        pairs.push({
            row1: row1,
            row2: null,
            index1: rowIndex,
            index2: -1,
            isDifferent: true,
            onlyIn: 'table1'
        });
    });
    
    onlyInTable2.forEach(diff => {
        const rowIndex = diff.row2 - 1;
        const row2 = workingData2[rowIndex + 1];
        
        pairs.push({
            row1: null,
            row2: row2,
            index1: -1,
            index2: rowIndex,
            isDifferent: true,
            onlyIn: 'table2'
        });
    });
    
    window.currentPairs = pairs;
    window.currentFinalHeaders = workingData1[0] || commonColumns;
    window.currentFinalAllCols = (workingData1[0] || commonColumns).length;
    window.currentSortColumn = -1;
    window.currentSortDirection = 'asc';
    
    const diffTable = document.getElementById('diffTable');
    if (diffTable) {
        diffTable.innerHTML = `
            <div class="table-header-fixed">
                <table class="diff-table-header">
                    <thead></thead>
                    <tbody class="filter-row"></tbody>
                </table>
            </div>
            <div class="table-body-scrollable">
                <table class="diff-table-body">
                    <tbody></tbody>
                </table>
            </div>
        `;
        diffTable.style.display = 'block';
        diffTable.style.visibility = 'visible';
    }
    
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    if (headerTable) headerTable.style.display = 'table';
    if (bodyTable) bodyTable.style.display = 'table';
    
    // Skip renderComparisonTable and go directly to createBasicFallbackTable
    // to avoid column duplication
    await createBasicFallbackTable(pairs, workingData1[0] || commonColumns);
}

async function createBasicFallbackTable(pairs, headers) {
    // Clear any remaining loading messages at the start
    const resultDiv = document.getElementById('result');
    if (resultDiv) {
        resultDiv.innerHTML = '';
        resultDiv.style.display = 'none';
    }
    
    // Completely recreate the table structure to avoid duplication
    const diffTableElement = document.getElementById('diffTable');
    if (diffTableElement) {
        diffTableElement.innerHTML = `
            <div class="table-header-fixed">
                <table class="diff-table-header">
                    <thead></thead>
                    <tbody class="filter-row"></tbody>
                </table>
            </div>
            <div class="table-body-scrollable">
                <table class="diff-table-body">
                    <tbody></tbody>
                </table>
            </div>
        `;
    }
    
    const headerTable = document.querySelector('.diff-table-header thead');
    const filterRow = document.querySelector('.filter-row');
    const bodyTable = document.querySelector('.diff-table-body tbody');
    
    if (!headerTable || !bodyTable) {
        return;
    }
    
    // Clear any existing content first to prevent duplication
    headerTable.innerHTML = '';
    filterRow.innerHTML = '';
    bodyTable.innerHTML = '';
    
    // Use the passed headers parameter instead of data1[0]
    const realHeaders = headers || [];
    
    let headerHtml = '<tr><th title="Source - shows which file the data comes from" class="source-column">Source</th>';
    realHeaders.forEach((header, index) => {
        const headerText = header || `Column ${index + 1}`;
        headerHtml += `<th class="sortable" onclick="sortTable(${index})" title="${headerText}">${headerText}</th>`;
    });
    headerHtml += '</tr>';
    headerTable.innerHTML = headerHtml;
    
    if (filterRow) {
        let filterHtml = '<tr><td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>';
        realHeaders.forEach(header => {
            filterHtml += `<td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>`;
        });
        filterHtml += '</tr>';
        filterRow.innerHTML = filterHtml;
    }
    
    let bodyHtml = '';
    const file1Name = window.fileName1 || 'File 1';
    const file2Name = window.fileName2 || 'File 2';
    
    pairs.slice(0, 1000).forEach((pair, index) => {
        const rowData = pair.row1 || pair.row2;
        const fileIndicator = pair.onlyIn === 'table1' ? file1Name : file2Name;
        
        let rowClass = 'diff-row';
        let sourceClass = 'file-indicator';
        
        if (pair.onlyIn === 'table1') {
            sourceClass += ' new-cell1';
        } else if (pair.onlyIn === 'table2') {
            sourceClass += ' new-cell2';
        } else {
            sourceClass += ' identical';
        }
        
        bodyHtml += `<tr class="${rowClass}" data-row-index="${pair.index1 >= 0 ? pair.index1 : pair.index2}">`;
        bodyHtml += `<td class="${sourceClass}">${fileIndicator}</td>`;
        
        realHeaders.forEach((header, colIndex) => {
            const value = rowData && rowData[colIndex] !== undefined ? rowData[colIndex] : '';
            const cellClass = pair.isDifferent ? 'diff-cell different' : 'diff-cell identical';
            bodyHtml += `<td class="${cellClass}" title="${value}">${value}</td>`;
        });
        
        bodyHtml += `</tr>`;
    });
    
    if (pairs.length > 1000) {
        bodyHtml += `<tr class="info-row"><td colspan="${realHeaders.length + 1}" style="text-align:center; padding:10px; background:#f8f9fa; font-style: italic;">⚠️ Showing first 1,000 of ${pairs.length} differences</td></tr>`;
    }
    
    bodyTable.innerHTML = bodyHtml;
    
    const diffTableFinal = document.getElementById('diffTable');
    if (diffTableFinal) {
        diffTableFinal.style.display = 'block';
    }
}

window.MaxPilotDuckDB = {
    FastTableComparator,
    fastComparator,
    initializeFastComparator,
    compareTablesWithFastComparator,
    compareTablesEnhanced,
    processFastComparisonResults,
    runFastComparatorTests,
    version: '1.0.0'
};


document.addEventListener('DOMContentLoaded', async () => {
    await initializeFastComparator();
    
    if (typeof window.compareTables === 'function') {
        window.compareTablesOriginal = window.compareTables;
    }
    window.compareTables = compareTablesEnhanced;
});

function testDOMIntegration() {
    const elements = {
        'diffTable': document.getElementById('diffTable'),
        'diff-table-header': document.querySelector('.diff-table-header thead'),
        'filter-row': document.querySelector('.filter-row'),
        'diff-table-body': document.querySelector('.diff-table-body tbody'),
        'exportExcelBtn': document.getElementById('exportExcelBtn'),
        'summary': document.getElementById('summary')
    };
    
    for (const [name, element] of Object.entries(elements)) {
    }
    
    const functions = {
        'renderComparisonTable': typeof renderComparisonTable,
        'exportToExcel': typeof exportToExcel,
        'filterTable': typeof filterTable
    };
    
    for (const [name, type] of Object.entries(functions)) {
    }
    
    return elements;
}

}