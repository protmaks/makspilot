/**
 * Local DuckDB WASM fallback for MaxPilot
 * Uses in-memory comparison without external workers to avoid CORS issues
 */

// Prevent duplicate declarations
if (typeof window.LocalDuckDBSimulator === 'undefined') {

class LocalDuckDBSimulator {
    constructor() {
        this.initialized = false;
        this.tables = new Map();
    }

    async initialize() {
        try {
            console.log('🔧 Initializing Local DuckDB Simulator...');
            
            // Check basic requirements
            if (typeof WebAssembly === 'undefined') {
                throw new Error('WebAssembly not supported');
            }
            
            // Simulate initialization delay
            await new Promise(resolve => setTimeout(resolve, 100));
            
            this.initialized = true;
            console.log('✅ Local DuckDB Simulator initialized');
            return true;
            
        } catch (error) {
            console.error('❌ Local DuckDB Simulator initialization failed:', error);
            this.initialized = false;
            return false;
        }
    }

    async createTableFromData(tableName, data, headers = null) {
        if (!this.initialized) {
            throw new Error('Simulator not initialized');
        }

        try {
            // Store data in memory
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
        // Simple hash function for row comparison
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
            throw new Error('Simulator not initialized');
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
        console.log(`⚡ Comparison completed in ${duration.toFixed(2)}ms`);

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
        console.log('🔒 Local DuckDB Simulator closed');
    }
}

// Global instance for local simulator
let localDuckDBManager = null;

// Initialize local simulator
async function initializeLocalDuckDB() {
    if (localDuckDBManager) {
        return localDuckDBManager;
    }

    try {
        localDuckDBManager = new LocalDuckDBSimulator();
        const initialized = await localDuckDBManager.initialize();
        
        if (initialized) {
            console.log('✅ Local DuckDB Simulator ready');
            
            // Update global references
            window.duckDBManager = localDuckDBManager;
            window.duckDBAvailable = true;
            
            return localDuckDBManager;
        } else {
            console.warn('⚠️ Local DuckDB Simulator initialization failed');
            return null;
        }
    } catch (error) {
        console.error('❌ Error initializing Local DuckDB Simulator:', error);
        return null;
    }
}

// Enhanced comparison function using local simulator
async function compareTablesWithLocalDuckDB(data1, data2, excludeColumns = [], useTolerance = false, tolerance = 1.5) {
    try {
        if (!localDuckDBManager || !localDuckDBManager.initialized) {
            console.log('Local DuckDB not available, using original comparison');
            return null;
        }

        console.log('🚀 Starting Local DuckDB fast comparison...');
        
        // Create tables
        await localDuckDBManager.createTableFromData('table1', data1);
        await localDuckDBManager.createTableFromData('table2', data2);
        
        // Perform fast comparison
        const comparisonResult = await localDuckDBManager.compareTablesFast(
            'table1', 'table2', excludeColumns
        );

        console.log('✅ Local DuckDB comparison completed:', {
            identical: comparisonResult.identical.length,
            onlyInTable1: comparisonResult.onlyInTable1.length,
            onlyInTable2: comparisonResult.onlyInTable2.length,
            performance: comparisonResult.performance
        });

        return comparisonResult;
        
    } catch (error) {
        console.error('❌ Local DuckDB comparison failed:', error);
        return null; // Fallback to original method
    }
}

// Initialize on page load
if (typeof window !== 'undefined') {
    window.LocalDuckDBSimulator = LocalDuckDBSimulator;
    window.initializeLocalDuckDB = initializeLocalDuckDB;
    window.compareTablesWithLocalDuckDB = compareTablesWithLocalDuckDB;
}

} // End of LocalDuckDBSimulator class declaration guard
