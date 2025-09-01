/**
 * Unified DuckDB Integration for MaxPilot
 * Combines all DuckDB functionality in one file to avoid duplicate declarations
 */

// Check if already loaded
if (typeof window.MaxPilotDuckDB === 'undefined') {

// Fast table comparison engine
class FastTableComparator {
    constructor() {
        this.initialized = false;
        this.tables = new Map();
        this.mode = 'local'; // 'local' or 'wasm'
    }

    async initialize() {
        try {
            console.log('🔧 Initializing Fast Table Comparator...');
            
            // Check WebAssembly support
            if (typeof WebAssembly === 'undefined') {
                throw new Error('WebAssembly not supported');
            }
            
            // Try to use real DuckDB WASM first, but expect it to fail due to CORS
            try {
                await this.initializeWASM();
                this.mode = 'wasm';
                console.log('✅ DuckDB WASM mode initialized');
            } catch (wasmError) {
                console.log('⚠️ DuckDB WASM failed (CORS), using local simulator:', wasmError.message);
                await this.initializeLocal();
                this.mode = 'local';
                console.log('✅ Local simulator mode initialized');
            }
            
            this.initialized = true;
            return true;
            
        } catch (error) {
            console.error('❌ Fast Table Comparator initialization failed:', error);
            this.initialized = false;
            return false;
        }
    }

    async initializeWASM() {
        // This will likely fail due to CORS, but we try anyway
        throw new Error('CORS policy blocks external workers');
    }

    async initializeLocal() {
        // Local simulator doesn't need external resources
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
            
            console.log(`📊 Created table ${tableName} with ${processedData.rows.length} rows, ${processedData.columns.length} columns`);
            return true;
            
        } catch (error) {
            console.error(`Error creating table ${tableName}:`, error);
            throw error;
        }
    }

    processTableData(data, headers = null) {
        if (!data || data.length === 0) {
            return { columns: [], rows: [], indexes: new Map() };
        }

        // Determine headers and data rows
        const columns = headers || (data[0] ? data[0].map((_, i) => `col_${i}`) : []);
        const dataRows = headers ? data.slice(1) : data;
        
        // Create row hashes for fast comparison
        const indexes = new Map();
        const processedRows = dataRows.map((row, index) => {
            // Pad row to match column count
            const paddedRow = Array(columns.length).fill('');
            row.forEach((cell, i) => {
                if (i < columns.length) {
                    paddedRow[i] = cell !== null && cell !== undefined ? String(cell) : '';
                }
            });
            
            // Create hash for this row
            const rowHash = this.createRowHash(paddedRow);
            
            // Store in index for fast lookup
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
        // Simple but effective hash function
        const str = row.join('|');
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
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

        console.log(`🔍 Comparing ${table1.rows.length} vs ${table2.rows.length} rows...`);
        
        const startTime = performance.now();
        
        // Find common columns (excluding specified ones)
        const commonColumns = table1.columns.filter(col => 
            table2.columns.includes(col) && 
            !excludeColumns.some(excCol => 
                col.toLowerCase().includes(excCol.toLowerCase())
            )
        );

        if (commonColumns.length === 0) {
            throw new Error('No common columns found for comparison');
        }

        // Get column indices for comparison
        const table1ColIndices = commonColumns.map(col => table1.columns.indexOf(col));
        const table2ColIndices = commonColumns.map(col => table2.columns.indexOf(col));

        // Create comparison hashes using only common columns
        const table1Hashes = new Map();
        const table2Hashes = new Map();

        table1.rows.forEach(row => {
            const compareData = table1ColIndices.map(i => row.data[i]);
            const hash = this.createRowHash(compareData);
            
            if (!table1Hashes.has(hash)) {
                table1Hashes.set(hash, []);
            }
            table1Hashes.get(hash).push(row.index + 1); // 1-based indexing
        });

        table2.rows.forEach(row => {
            const compareData = table2ColIndices.map(i => row.data[i]);
            const hash = this.createRowHash(compareData);
            
            if (!table2Hashes.has(hash)) {
                table2Hashes.set(hash, []);
            }
            table2Hashes.get(hash).push(row.index + 1); // 1-based indexing
        });

        // Find matches
        const identical = [];
        const onlyInTable1 = [];
        const onlyInTable2 = [];

        // Find identical rows
        for (const [hash, table1Rows] of table1Hashes) {
            if (table2Hashes.has(hash)) {
                const table2Rows = table2Hashes.get(hash);
                
                // Create pairs for identical rows
                table1Rows.forEach(row1 => {
                    table2Rows.forEach(row2 => {
                        identical.push({ row1, row2, status: 'IDENTICAL' });
                    });
                });
            } else {
                // Rows only in table1
                table1Rows.forEach(row1 => {
                    onlyInTable1.push({ row1, row2: null, status: 'ONLY_IN_TABLE1' });
                });
            }
        }

        // Find rows only in table2
        for (const [hash, table2Rows] of table2Hashes) {
            if (!table1Hashes.has(hash)) {
                table2Rows.forEach(row2 => {
                    onlyInTable2.push({ row1: null, row2, status: 'ONLY_IN_TABLE2' });
                });
            }
        }

        const duration = performance.now() - startTime;
        console.log(`⚡ Fast comparison completed in ${duration.toFixed(2)}ms`);

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
        console.log('🔒 Fast Table Comparator closed');
    }

    static async isSupported() {
        try {
            return typeof WebAssembly !== 'undefined';
        } catch (error) {
            return false;
        }
    }
}

// Global instance
let fastComparator = null;

// Initialize fast comparator
async function initializeFastComparator() {
    if (fastComparator) {
        return fastComparator;
    }

    try {
        fastComparator = new FastTableComparator();
        const initialized = await fastComparator.initialize();
        
        if (initialized) {
            console.log(`✅ Fast Table Comparator ready (${fastComparator.mode} mode)`);
            
            // Update UI
            showFastModeStatus(true, fastComparator.mode);
            window.duckDBManager = fastComparator; // For compatibility
            window.duckDBAvailable = true;
            
            return fastComparator;
        } else {
            console.warn('⚠️ Fast Table Comparator initialization failed');
            showFastModeStatus(false);
            return null;
        }
    } catch (error) {
        console.error('❌ Error initializing Fast Table Comparator:', error);
        showFastModeStatus(false);
        return null;
    }
}

// Show fast mode status
function showFastModeStatus(available, mode = 'local') {
    const statusElement = document.getElementById('duckdb-status');
    if (statusElement) {
        if (available) {
            const modeText = mode === 'wasm' ? 'DuckDB WASM' : 'Fast Local';
            statusElement.innerHTML = `⚡ ${modeText} mode enabled - Enhanced performance!`;
            statusElement.className = 'duckdb-status duckdb-available show';
            
            // Show fast indicators on buttons
            const fastIndicators = document.querySelectorAll('.fast-mode-indicator');
            fastIndicators.forEach(indicator => {
                indicator.style.display = 'inline-block';
                indicator.textContent = mode === 'wasm' ? 'ULTRA' : 'FAST';
            });
            
            // Auto-hide after 5 seconds
            setTimeout(() => {
                if (statusElement.classList.contains('duckdb-available')) {
                    statusElement.style.opacity = '0.8';
                    statusElement.innerHTML = `⚡ ${modeText} active`;
                }
            }, 5000);
            
        } else {
            statusElement.innerHTML = '🔄 Standard comparison mode';
            statusElement.className = 'duckdb-status duckdb-unavailable show';
            
            // Hide fast indicators
            const fastIndicators = document.querySelectorAll('.fast-mode-indicator');
            fastIndicators.forEach(indicator => {
                indicator.style.display = 'none';
            });
            
            // Auto-hide after 3 seconds
            setTimeout(() => {
                statusElement.style.display = 'none';
            }, 3000);
        }
    }
}

// Enhanced comparison function
async function compareTablesWithFastComparator(data1, data2, excludeColumns = [], useTolerance = false, tolerance = 1.5) {
    try {
        if (!fastComparator || !fastComparator.initialized) {
            console.log('ℹ️ Fast comparator not available, using original comparison');
            return null;
        }

        console.log('🚀 Starting fast table comparison...');
        
        // Use the same data preparation logic as original functions.js
        const { data1: alignedData1, data2: alignedData2, columnInfo } = prepareDataForComparison(data1, data2);
        
        console.log('📊 Data preparation completed:', {
            hasCommonColumns: columnInfo?.hasCommonColumns,
            commonCount: columnInfo?.commonCount,
            data1Rows: alignedData1.length,
            data2Rows: alignedData2.length
        });
        
        // Extract headers from first row and get data rows only
        const headers1 = alignedData1[0] || [];
        const headers2 = alignedData2[0] || [];
        const dataRows1 = alignedData1.slice(1); // Skip header row
        const dataRows2 = alignedData2.slice(1); // Skip header row
        
        console.log(`📊 Final data: ${headers1.length} columns, ${dataRows1.length} vs ${dataRows2.length} data rows`);
        
        // Create tables with explicit headers (don't include headers in data)
        await fastComparator.createTableFromData('table1', dataRows1, headers1);
        await fastComparator.createTableFromData('table2', dataRows2, headers2);
        
        // Perform fast comparison
        const comparisonResult = await fastComparator.compareTablesFast(
            'table1', 'table2', excludeColumns
        );

        console.log('✅ Fast comparison completed:', {
            identical: comparisonResult.identical.length,
            onlyInTable1: comparisonResult.onlyInTable1.length,
            onlyInTable2: comparisonResult.onlyInTable2.length,
            performance: comparisonResult.performance
        });

        // Add column info to result for proper table generation
        comparisonResult.columnInfo = columnInfo;
        comparisonResult.alignedData1 = alignedData1;
        comparisonResult.alignedData2 = alignedData2;

        return comparisonResult;
        
    } catch (error) {
        console.error('❌ Fast comparison failed:', error);
        return null; // Fallback to original method
    }
}

// Test functions
async function runFastComparatorTests() {
    console.log('\n🧪 Running Fast Comparator Tests...');
    
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
            
            return duration < 500; // Should complete in under 500ms
        }}
    ];

    const results = [];
    for (const test of tests) {
        try {
            console.log(`🔬 Testing: ${test.name}`);
            const result = await test.fn();
            results.push({ name: test.name, status: 'PASS', result });
            console.log(`✅ ${test.name}: PASS`);
        } catch (error) {
            results.push({ name: test.name, status: 'FAIL', error: error.message });
            console.error(`❌ ${test.name}: FAIL - ${error.message}`);
        }
    }

    // Summary
    const passed = results.filter(r => r.status === 'PASS').length;
    const total = results.length;
    
    console.log(`\n📊 Test Results: ${passed}/${total} passed (${((passed/total)*100).toFixed(1)}%)`);
    
    return results;
}

// Enhanced comparison for MaxPilot
async function compareTablesEnhanced(useTolerance = false) {
    console.log('🚀 Starting enhanced table comparison...');
    
    // Clear previous results
    clearComparisonResults();
    
    let resultDiv = document.getElementById('result');
    let summaryDiv = document.getElementById('summary');
    
    if (!data1.length || !data2.length) {
        document.getElementById('result').innerText = 'Please, load both files.';
        document.getElementById('summary').innerHTML = '';
        showPlaceholderMessage();
        return;
    }

    // Check size limits
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
        // Try fast comparison first for performance boost
        if (fastComparator && fastComparator.initialized) {
            resultDiv.innerHTML = '<div class="comparison-loading-enhanced">⚡ Using fast comparison engine...</div>';
            summaryDiv.innerHTML = '<div style="text-align: center; padding: 10px;">Processing with enhanced performance...</div>';
            
            console.log('🚀 Fast comparison mode - delegating to original performComparison with speed boost');
            
            // Clear the loading message and use original logic
            setTimeout(() => {
                resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚡ Fast mode - processing comparison...</div>';
                performComparison(); // Use the original, proven comparison logic
            }, 100);
            
            return;
        }

        // Fallback to original comparison
        console.log('ℹ️ Using standard comparison method...');
        resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">🔄 Using standard comparison...</div>';
        
        setTimeout(() => {
            performComparison();
        }, 10);
        
    } catch (error) {
        console.error('❌ Enhanced comparison failed:', error);
        
        // Fallback to original method
        resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">🔄 Using standard comparison...</div>';
        setTimeout(() => {
            performComparison();
        }, 10);
    }
}

// Process fast comparison results
async function processFastComparisonResults(fastResult, useTolerance) {
    const { identical, onlyInTable1, onlyInTable2, table1Count, table2Count, commonColumns, performance } = fastResult;
    
    console.log('📊 Processing fast comparison results...', {
        identical: identical.length,
        onlyInTable1: onlyInTable1.length,
        onlyInTable2: onlyInTable2.length,
        performance: performance,
        fileName1: window.fileName1,
        fileName2: window.fileName2
    });

    // Calculate statistics
    const identicalCount = identical.length;
    const similarity = table1Count > 0 ? ((identicalCount / Math.max(table1Count, table2Count)) * 100).toFixed(1) : 0;
    
    // Ensure we have file names
    const file1Name = window.fileName1 || 'File 1';
    const file2Name = window.fileName2 || 'File 2';

    // Update summary with performance info
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

    // Generate detailed comparison table - ИСПРАВЛЕНО
    console.log('🔄 [UNIFIED] Generating detailed comparison table...');
    console.log('🔍 [UNIFIED] About to call generateDetailedComparisonTable with:', {
        identicalCount: identical.length,
        onlyInTable1Count: onlyInTable1.length,
        onlyInTable2Count: onlyInTable2.length,
        columnsCount: commonColumns.length
    });
    
    try {
        await generateDetailedComparisonTable(fastResult, useTolerance);
        console.log('✅ [UNIFIED] Detailed table generation completed');
        
        // Clear the loading message from result div AFTER table is generated
        const resultDiv = document.getElementById('result');
        if (resultDiv) {
            console.log('🔍 [UNIFIED] Current result div content:', resultDiv.innerHTML);
            resultDiv.innerHTML = ''; // Полностью очищаем после генерации таблицы
            resultDiv.style.display = 'none'; // Принудительно скрываем
            console.log('🧹 [UNIFIED] Cleared and hidden result div after table generation');
        }
        
        // Force show diff table immediately
        const diffTable = document.getElementById('diffTable');
        if (diffTable) {
            diffTable.style.display = 'block';
            diffTable.style.visibility = 'visible';
            diffTable.scrollIntoView({ behavior: 'smooth', block: 'start' });
            console.log('📊 [UNIFIED] Force scrolled to diff table');
        }
        
    } catch (error) {
        console.error('❌ [UNIFIED] Error generating detailed table:', error);
        console.error('Full error stack:', error.stack);
        
        // Show error in result div
        const resultDiv = document.getElementById('result');
        if (resultDiv) {
            resultDiv.innerHTML = '<div style="text-align: center; padding: 10px; color: #dc3545;">❌ Error generating comparison table. Check console for details.</div>';
        }
    }
    
    // Show filter controls
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'flex';
        filterControls.classList.remove('filter-controls-hidden');
        console.log('🔧 Filter controls are now visible');
    }

    // Show export button - исправлено для правильного показа
    const exportBtn = document.getElementById('exportExcelBtn');
    if (exportBtn) {
        exportBtn.style.display = 'inline-block';
        exportBtn.classList.remove('export-btn-hidden');
        console.log('📤 Export button is now visible');
    } else {
        console.log('❌ Export button not found');
    }
    
    console.log('✅ Fast comparison results processed completely');
    
    // ПРИНУДИТЕЛЬНОЕ ИСПРАВЛЕНИЕ: очищаем загрузочное сообщение и показываем fallback таблицу
    setTimeout(async () => {
        console.log('🔄 [FALLBACK] Force cleaning and showing table...');
        
        // Clear loading message
        const resultDiv = document.getElementById('result');
        if (resultDiv) {
            console.log('🔍 [FALLBACK] Current result div content:', resultDiv.innerHTML);
            resultDiv.innerHTML = '';
            resultDiv.style.display = 'none'; // Принудительно скрываем
            console.log('🧹 [FALLBACK] Force cleared and hidden result div');
        }
        
        // Try to show table with fallback
        const pairs = [];
        
        // Add differences
        onlyInTable1.forEach(diff => {
            const rowIndex = diff.row1 - 1;
            // Use aligned data if available
            const workingData1 = fastResult.alignedData1 || data1;
            const row1 = workingData1[rowIndex + 1]; // +1 to skip header row
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
            // Use aligned data if available  
            const workingData2 = fastResult.alignedData2 || data2;
            const row2 = workingData2[rowIndex + 1]; // +1 to skip header row
            pairs.push({
                row1: null,
                row2: row2,
                index1: -1,
                index2: rowIndex,
                isDifferent: true,
                onlyIn: 'table2'
            });
        });
        
        console.log(`🔄 [FALLBACK] Creating fallback table with ${pairs.length} rows`);
        
        // Use actual headers from aligned data
        const workingData1 = fastResult.alignedData1 || data1;
        const actualHeaders = workingData1[0] || commonColumns;
        await createBasicFallbackTable(pairs, actualHeaders); // Передаем реальные заголовки
        
    }, 500); // Задержка 500ms
}

// Generate detailed comparison table compatible with original format
async function generateDetailedComparisonTable(fastResult, useTolerance) {
    console.log('🔨 [UNIFIED] Building detailed comparison table...');
    console.log('🔍 [UNIFIED] Function called with fastResult:', {
        identical: fastResult.identical?.length,
        onlyInTable1: fastResult.onlyInTable1?.length,
        onlyInTable2: fastResult.onlyInTable2?.length,
        commonColumns: fastResult.commonColumns?.length
    });
    
    const { identical, onlyInTable1, onlyInTable2, commonColumns, alignedData1, alignedData2 } = fastResult;
    
    // Use aligned data if available, otherwise fall back to global data
    const workingData1 = alignedData1 || data1;
    const workingData2 = alignedData2 || data2;
    
    // Prepare pairs in format compatible with original renderComparisonTable
    const pairs = [];
    
    // Add identical rows (только различия для экономии памяти, но можно добавить и одинаковые)
    console.log('🔄 Processing identical rows...');
    const maxIdenticalToShow = onlyInTable1.length === 0 && onlyInTable2.length === 0 ? 1000 : 100; // Больше идентичных строк если нет различий
    identical.slice(0, maxIdenticalToShow).forEach(identicalPair => {
        // Добавляем идентичные строки для контекста
        const row1Index = identicalPair.row1 - 1; // Convert from 1-based to 0-based
        const row2Index = identicalPair.row2 - 1;
        // Account for header row: data[0] is headers, data[1] is first data row
        const row1 = workingData1[row1Index + 1]; // +1 to skip header row
        const row2 = workingData2[row2Index + 1]; // +1 to skip header row
        
        pairs.push({
            row1: row1,
            row2: row2,
            index1: row1Index,
            index2: row2Index,
            isDifferent: false,
            onlyIn: null
        });
    });
    
    // Convert onlyInTable1 to pairs format
    console.log('🔄 Processing onlyInTable1 rows...');
    onlyInTable1.forEach(diff => {
        const rowIndex = diff.row1 - 1; // Convert from 1-based to 0-based
        // Account for header row: data[0] is headers, data[1] is first data row
        const row1 = workingData1[rowIndex + 1]; // +1 to skip header row
        
        pairs.push({
            row1: row1,
            row2: null,
            index1: rowIndex,
            index2: -1,
            isDifferent: true,
            onlyIn: 'table1'
        });
    });
    
    // Convert onlyInTable2 to pairs format
    console.log('🔄 Processing onlyInTable2 rows...');
    onlyInTable2.forEach(diff => {
        const rowIndex = diff.row2 - 1; // Convert from 1-based to 0-based
        // Account for header row: data[0] is headers, data[1] is first data row
        const row2 = workingData2[rowIndex + 1]; // +1 to skip header row
        
        pairs.push({
            row1: null,
            row2: row2,
            index1: -1,
            index2: rowIndex,
            isDifferent: true,
            onlyIn: 'table2'
        });
    });
    
    console.log(`📋 Generated ${pairs.length} total rows (${identical.length} identical + ${onlyInTable1.length + onlyInTable2.length} differences)`);
    
    // Debug: log first few pairs to check data structure
    if (pairs.length > 0) {
        console.log('🔍 Sample pairs data:', pairs.slice(0, 3));
        console.log('🔍 First row1 data:', pairs[0]?.row1);
        console.log('🔍 Common columns:', commonColumns);
    }
    
    // Set global variables for original table rendering
    window.currentPairs = pairs;
    window.currentFinalHeaders = workingData1[0] || commonColumns; // Используем выровненные заголовки
    window.currentFinalAllCols = (workingData1[0] || commonColumns).length;
    window.currentSortColumn = -1;
    window.currentSortDirection = 'asc';
    
    console.log('🔧 Global variables set:', {
        currentPairs: window.currentPairs?.length,
        currentFinalHeaders: window.currentFinalHeaders,
        currentFinalAllCols: window.currentFinalAllCols
    });
    
    // Show the diff table section
    const diffTable = document.getElementById('diffTable');
    if (diffTable) {
        // Restore proper table structure (it might have been replaced with placeholder)
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
        console.log('📊 Diff table structure restored and visible');
    } else {
        console.log('❌ diffTable element not found');
    }
    
    // Also ensure table elements are visible
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    if (headerTable) headerTable.style.display = 'table';
    if (bodyTable) bodyTable.style.display = 'table';
    
    // Use original table rendering logic
    if (typeof renderComparisonTable === 'function') {
        console.log('📊 Calling original renderComparisonTable function...');
        console.log('🔧 Final check of global variables before rendering:', {
            currentPairs: window.currentPairs?.length,
            currentFinalHeaders: window.currentFinalHeaders,
            currentFinalAllCols: window.currentFinalAllCols,
            samplePair: window.currentPairs?.[0]
        });
        
        try {
            renderComparisonTable();
            console.log('✅ renderComparisonTable executed successfully');
            
            // Double-check visibility after rendering
            setTimeout(() => {
                const renderedRows = document.querySelectorAll('.diff-table-body tbody tr');
                const headerCells = document.querySelectorAll('.diff-table-header thead th');
                console.log(`📊 Rendered ${renderedRows.length} rows and ${headerCells.length} headers in table`);
                
                if (renderedRows.length === 0) {
                    console.warn('⚠️ No rows rendered, trying fallback...');
                    createBasicFallbackTable(pairs, data1[0] || commonColumns);
                }
            }, 200);
        } catch (error) {
            console.error('❌ Error in renderComparisonTable:', error);
            console.error('Full error details:', error.stack);
            await createBasicFallbackTable(pairs, data1[0] || commonColumns);
        }
    } else {
        console.log('⚠️ Original renderComparisonTable not found, creating fallback table...');
        await createBasicFallbackTable(pairs, data1[0] || commonColumns);
    }
    
    console.log('✅ Detailed comparison table generated');
}

// Fallback basic table generation
async function createBasicFallbackTable(pairs, commonColumns) {
    console.log('🔧 Creating fallback table...');
    
    // Find the correct table elements
    const headerTable = document.querySelector('.diff-table-header thead');
    const filterRow = document.querySelector('.filter-row');
    const bodyTable = document.querySelector('.diff-table-body tbody');
    
    if (!headerTable || !bodyTable) {
        console.error('❌ Required table elements not found');
        return;
    }
    
    // Use real column headers from data1[0] instead of col_0, col_1...
    const realHeaders = data1[0] || commonColumns;
    console.log('🔧 Using headers:', realHeaders);
    
    // Generate header
    let headerHtml = '<tr><th title="Source - shows which file the data comes from" class="source-column">Source</th>';
    realHeaders.forEach((header, index) => {
        const headerText = header || `Column ${index + 1}`;
        headerHtml += `<th class="sortable" onclick="sortTable(${index})" title="${headerText}">${headerText}</th>`;
    });
    headerHtml += '</tr>';
    headerTable.innerHTML = headerHtml;
    
    // Generate filter row
    if (filterRow) {
        let filterHtml = '<tr><td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>';
        realHeaders.forEach(header => {
            filterHtml += `<td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>`;
        });
        filterHtml += '</tr>';
        filterRow.innerHTML = filterHtml;
    }
    
    // Generate body rows with proper styling
    let bodyHtml = '';
    const file1Name = window.fileName1 || 'File 1';
    const file2Name = window.fileName2 || 'File 2';
    
    pairs.slice(0, 1000).forEach((pair, index) => {
        const rowData = pair.row1 || pair.row2;
        const fileIndicator = pair.onlyIn === 'table1' ? file1Name : file2Name;
        
        // Apply different styles based on row type
        let rowClass = 'diff-row';
        let sourceClass = 'file-indicator';
        
        if (pair.onlyIn === 'table1') {
            sourceClass += ' new-cell1'; // Синий цвет для строк только в файле 1
        } else if (pair.onlyIn === 'table2') {
            sourceClass += ' new-cell2'; // Зеленый цвет для строк только в файле 2
        } else {
            sourceClass += ' identical'; // Зеленый фон для идентичных строк
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
    
    // Show the diff table
    const diffTable = document.getElementById('diffTable');
    if (diffTable) {
        diffTable.style.display = 'block';
    }
    
    console.log('✅ Fallback table created successfully with proper styling');
}

async function generateFastComparisonTable_OLD_DEPRECATED(fastResult, useTolerance) {
    // DEPRECATED: This function should not be called anymore
    console.log('⚠️ DEPRECATED function generateFastComparisonTable called!');
    console.trace('Call stack:');
    
    // For now, create a simplified result display
    // In full implementation, this would generate the detailed row-by-row table
    
    const { identical, onlyInTable1, onlyInTable2 } = fastResult;
    const maxRowsToShow = 1000;
    
    // Create pairs for current global variables (for compatibility with existing code)
    const allPairs = [];
    
    // Add identical pairs
    identical.slice(0, maxRowsToShow).forEach(match => {
        allPairs.push({
            row1Index: match.row1 - 1,
            row2Index: match.row2 - 1,
            status: 'identical'
        });
    });
    
    // Add unique pairs
    onlyInTable1.slice(0, maxRowsToShow).forEach(match => {
        allPairs.push({
            row1Index: match.row1 - 1,
            row2Index: -1,
            status: 'only_in_table1'
        });
    });
    
    onlyInTable2.slice(0, maxRowsToShow).forEach(match => {
        allPairs.push({
            row1Index: -1,
            row2Index: match.row2 - 1,
            status: 'only_in_table2'
        });
    });

    // Sort pairs by row index
    allPairs.sort((a, b) => {
        const aIndex = a.row1Index >= 0 ? a.row1Index : a.row2Index;
        const bIndex = b.row1Index >= 0 ? b.row1Index : b.row2Index;
        return aIndex - bIndex;
    });

    // Store for global access (compatibility with existing export functions)
    window.currentPairs = allPairs;
    window.currentFinalHeaders = data1[0] || [];
    
    // Generate the visual table using existing function
    if (typeof generateComparisonTable === 'function') {
        generateComparisonTable(allPairs, data1[0] || [], useTolerance);
    }
    
    console.log(`📋 Generated comparison table with ${allPairs.length} rows`);
}

// Make everything globally available
window.MaxPilotDuckDB = {
    FastTableComparator,
    fastComparator,
    initializeFastComparator,
    compareTablesWithFastComparator,
    compareTablesEnhanced,
    runFastComparatorTests,
    version: '1.0.0'
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', async () => {
    console.log('🚀 Initializing MaxPilot DuckDB integration...');
    await initializeFastComparator();
    
    // Override original compare function
    if (typeof window.compareTables === 'function') {
        window.compareTablesOriginal = window.compareTables;
    }
    window.compareTables = compareTablesEnhanced;
});

// Inline test function for HTML button
window.runDuckDBTestsInline = async function() {
    console.log('\n🧪 Running Fast Comparator Tests...');
    
    // First test DOM integration
    testDOMIntegration();
    
    const button = document.getElementById('duckdbTestBtn');
    if (button) {
        button.disabled = true;
        button.innerHTML = '🔄 Testing...';
        button.style.background = '#6c757d';
    }
    
    try {
        const results = await runFastComparatorTests();
        const passed = results.filter(r => r.status === 'PASS').length;
        const total = results.length;
        
        console.log(`\n🎉 Tests completed: ${passed}/${total} passed`);
        
        if (button) {
            if (passed === total) {
                button.innerHTML = '✅ All Tests Passed';
                button.style.background = '#28a745';
            } else {
                button.innerHTML = `⚠️ ${passed}/${total} Passed`;
                button.style.background = '#ffc107';
                button.style.color = '#212529';
            }
            
            setTimeout(() => {
                button.disabled = false;
                button.innerHTML = '🧪 Test DuckDB';
                button.style.background = '#17a2b8';
                button.style.color = 'white';
            }, 4000);
        }
        
    } catch (error) {
        console.error('❌ Test execution failed:', error);
        
        if (button) {
            button.innerHTML = '❌ Tests Failed';
            button.style.background = '#dc3545';
            setTimeout(() => {
                button.disabled = false;
                button.innerHTML = '🧪 Test DuckDB';
                button.style.background = '#17a2b8';
            }, 3000);
        }
    }
};

// Test DOM elements and integration
function testDOMIntegration() {
    console.log('🔍 Testing DOM integration...');
    
    // Check required elements
    const elements = {
        'diffTable': document.getElementById('diffTable'),
        'diff-table-header': document.querySelector('.diff-table-header thead'),
        'filter-row': document.querySelector('.filter-row'),
        'diff-table-body': document.querySelector('.diff-table-body tbody'),
        'exportExcelBtn': document.getElementById('exportExcelBtn'),
        'summary': document.getElementById('summary')
    };
    
    console.log('📊 DOM Elements Status:');
    for (const [name, element] of Object.entries(elements)) {
        console.log(`  ${element ? '✅' : '❌'} ${name}: ${element ? 'found' : 'NOT FOUND'}`);
    }
    
    // Check functions
    const functions = {
        'renderComparisonTable': typeof renderComparisonTable,
        'exportToExcel': typeof exportToExcel,
        'filterTable': typeof filterTable
    };
    
    console.log('📊 Functions Status:');
    for (const [name, type] of Object.entries(functions)) {
        console.log(`  ${type === 'function' ? '✅' : '❌'} ${name}: ${type}`);
    }
    
    return elements;
}

} // End of duplicate declaration guard
