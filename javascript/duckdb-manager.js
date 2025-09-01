/**
 * DuckDB WASM Manager for MaxPilot
 * Handles database operations for fast table comparison
 */

// Prevent duplicate declarations
if (typeof window.DuckDBManager === 'undefined') {

class DuckDBManager {
    constructor() {
        this.db = null;
        this.conn = null;
        this.initialized = false;
        this.isLoading = false;
    }

    async initialize() {
        if (this.initialized || this.isLoading) {
            return this.initialized;
        }

        this.isLoading = true;
        
        try {
            // Load DuckDB WASM - используем другой способ загрузки для совместимости
            let duckdb;
            
            // Попробуем загрузить через dynamic import
            try {
                duckdb = await import('https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@latest/+esm');
            } catch (importError) {
                // Fallback: загрузка через script tag
                console.log('Dynamic import failed, trying script loading...');
                throw new Error('Dynamic import not supported');
            }
            
            const MANUAL_BUNDLES = {
                mvp: {
                    mainModule: 'https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@latest/dist/duckdb-mvp.wasm',
                    mainWorker: 'https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@latest/dist/duckdb-browser-mvp.worker.js',
                },
                eh: {
                    mainModule: 'https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@latest/dist/duckdb-eh.wasm',
                    mainWorker: 'https://cdn.jsdelivr.net/npm/@duckdb/duckdb-wasm@latest/dist/duckdb-browser-eh.worker.js',
                },
            };

            // Select bundle
            const bundle = await duckdb.selectBundle(MANUAL_BUNDLES);
            
            // Instantiate the asynchronous version of DuckDB-wasm
            const worker = new Worker(bundle.mainWorker);
            const logger = new duckdb.ConsoleLogger();
            this.db = new duckdb.AsyncDuckDB(logger, worker);
            await this.db.instantiate(bundle.mainModule);
            
            // Create connection
            this.conn = await this.db.connect();
            
            this.initialized = true;
            console.log('DuckDB WASM initialized successfully');
            return true;
            
        } catch (error) {
            console.error('Failed to initialize DuckDB WASM:', error);
            this.initialized = false;
            return false;
        } finally {
            this.isLoading = false;
        }
    }

    async createTableFromData(tableName, data, headers = null) {
        if (!this.initialized) {
            throw new Error('DuckDB not initialized');
        }

        try {
            // Drop table if exists
            await this.conn.query(`DROP TABLE IF EXISTS ${tableName}`);
            
            // If no data, return
            if (!data || data.length === 0) {
                return;
            }

            // Determine headers
            const actualHeaders = headers || (data[0] ? data[0].map((_, i) => `col_${i}`) : []);
            const dataRows = headers ? data.slice(1) : data;
            
            // Create table schema
            const columns = actualHeaders.map((header, index) => {
                // Sanitize column names
                const cleanHeader = header.toString()
                    .replace(/[^a-zA-Z0-9_]/g, '_')
                    .replace(/^(\d)/, '_$1') // Add underscore if starts with number
                    || `col_${index}`;
                
                return `"${cleanHeader}" VARCHAR`;
            }).join(', ');

            await this.conn.query(`CREATE TABLE ${tableName} (${columns})`);

            // Insert data in batches for better performance
            const batchSize = 1000;
            for (let i = 0; i < dataRows.length; i += batchSize) {
                const batch = dataRows.slice(i, i + batchSize);
                const values = batch.map(row => {
                    const paddedRow = Array(actualHeaders.length).fill('');
                    row.forEach((cell, index) => {
                        if (index < actualHeaders.length) {
                            paddedRow[index] = cell !== null && cell !== undefined ? cell.toString() : '';
                        }
                    });
                    return `(${paddedRow.map(cell => `'${cell.replace(/'/g, "''")}'`).join(', ')})`;
                }).join(', ');

                if (values) {
                    await this.conn.query(`INSERT INTO ${tableName} VALUES ${values}`);
                }
            }

            console.log(`Created table ${tableName} with ${dataRows.length} rows`);
            
        } catch (error) {
            console.error(`Error creating table ${tableName}:`, error);
            throw error;
        }
    }

    async compareTablesFast(table1Name, table2Name, excludeColumns = []) {
        if (!this.initialized) {
            throw new Error('DuckDB not initialized');
        }

        try {
            // Get column info for both tables
            const table1Cols = await this.getTableColumns(table1Name);
            const table2Cols = await this.getTableColumns(table2Name);
            
            // Find common columns (excluding specified ones)
            const commonColumns = table1Cols.filter(col => 
                table2Cols.includes(col) && 
                !excludeColumns.some(excCol => 
                    col.toLowerCase().includes(excCol.toLowerCase())
                )
            );

            if (commonColumns.length === 0) {
                throw new Error('No common columns found for comparison');
            }

            const columnList = commonColumns.map(col => `"${col}"`).join(', ');
            
            // Create a unique identifier for each row (hash of all column values)
            const hashQuery = `
                WITH table1_hashed AS (
                    SELECT *, hash(${columnList}) as row_hash, 
                           ROW_NUMBER() OVER() as row_num
                    FROM ${table1Name}
                ),
                table2_hashed AS (
                    SELECT *, hash(${columnList}) as row_hash,
                           ROW_NUMBER() OVER() as row_num  
                    FROM ${table2Name}
                )
            `;

            // Find identical rows
            const identicalQuery = `
                ${hashQuery}
                SELECT t1.row_num as row1, t2.row_num as row2, 'IDENTICAL' as status
                FROM table1_hashed t1
                INNER JOIN table2_hashed t2 ON t1.row_hash = t2.row_hash
            `;

            // Find rows only in table1
            const onlyInTable1Query = `
                ${hashQuery}
                SELECT t1.row_num as row1, NULL as row2, 'ONLY_IN_TABLE1' as status
                FROM table1_hashed t1
                LEFT JOIN table2_hashed t2 ON t1.row_hash = t2.row_hash
                WHERE t2.row_hash IS NULL
            `;

            // Find rows only in table2  
            const onlyInTable2Query = `
                ${hashQuery}
                SELECT NULL as row1, t2.row_num as row2, 'ONLY_IN_TABLE2' as status
                FROM table2_hashed t2
                LEFT JOIN table1_hashed t1 ON t1.row_hash = t2.row_hash
                WHERE t1.row_hash IS NULL
            `;

            // Execute queries
            const [identical, onlyIn1, onlyIn2] = await Promise.all([
                this.conn.query(identicalQuery),
                this.conn.query(onlyInTable1Query), 
                this.conn.query(onlyInTable2Query)
            ]);

            // Get statistics
            const table1Count = await this.conn.query(`SELECT COUNT(*) as count FROM ${table1Name}`);
            const table2Count = await this.conn.query(`SELECT COUNT(*) as count FROM ${table2Name}`);

            return {
                identical: identical.toArray(),
                onlyInTable1: onlyIn1.toArray(),
                onlyInTable2: onlyIn2.toArray(),
                table1Count: table1Count.toArray()[0].count,
                table2Count: table2Count.toArray()[0].count,
                commonColumns: commonColumns
            };

        } catch (error) {
            console.error('Error in fast comparison:', error);
            throw error;
        }
    }

    async getTableColumns(tableName) {
        if (!this.initialized) {
            throw new Error('DuckDB not initialized');
        }

        try {
            const result = await this.conn.query(`DESCRIBE ${tableName}`);
            return result.toArray().map(row => row.column_name);
        } catch (error) {
            console.error(`Error getting columns for ${tableName}:`, error);
            return [];
        }
    }

    async findRowDifferences(table1Name, table2Name, excludeColumns = [], tolerance = 0) {
        if (!this.initialized) {
            throw new Error('DuckDB not initialized');
        }

        try {
            const table1Cols = await this.getTableColumns(table1Name);
            const table2Cols = await this.getTableColumns(table2Name);
            
            const commonColumns = table1Cols.filter(col => 
                table2Cols.includes(col) && 
                !excludeColumns.some(excCol => 
                    col.toLowerCase().includes(excCol.toLowerCase())
                )
            );

            // For tolerance comparison, we'll need a more complex query
            if (tolerance > 0) {
                return await this.findRowDifferencesWithTolerance(
                    table1Name, table2Name, commonColumns, tolerance
                );
            }

            // Simple comparison without tolerance
            const diffQuery = `
                WITH numbered_t1 AS (
                    SELECT *, ROW_NUMBER() OVER() as row_num FROM ${table1Name}
                ),
                numbered_t2 AS (
                    SELECT *, ROW_NUMBER() OVER() as row_num FROM ${table2Name}
                )
                SELECT 
                    t1.row_num as row1,
                    t2.row_num as row2,
                    ${commonColumns.map(col => `
                        CASE 
                            WHEN t1."${col}" != t2."${col}" THEN '${col}'
                            ELSE NULL 
                        END`).join(', ')} as diff_columns
                FROM numbered_t1 t1
                FULL OUTER JOIN numbered_t2 t2 ON t1.row_num = t2.row_num
                WHERE ${commonColumns.map(col => `t1."${col}" != t2."${col}"`).join(' OR ')}
            `;

            const result = await this.conn.query(diffQuery);
            return result.toArray();

        } catch (error) {
            console.error('Error finding row differences:', error);
            throw error;
        }
    }

    async findRowDifferencesWithTolerance(table1Name, table2Name, columns, tolerance) {
        // Implement tolerance-based comparison
        // This is more complex and requires checking numeric columns with tolerance
        const toleranceQueries = columns.map(col => {
            return `
                CASE 
                    WHEN TRY_CAST(t1."${col}" AS DOUBLE) IS NOT NULL 
                         AND TRY_CAST(t2."${col}" AS DOUBLE) IS NOT NULL THEN
                        CASE 
                            WHEN ABS(CAST(t1."${col}" AS DOUBLE) - CAST(t2."${col}" AS DOUBLE)) > 
                                 (GREATEST(ABS(CAST(t1."${col}" AS DOUBLE)), ABS(CAST(t2."${col}" AS DOUBLE))) * ${tolerance / 100})
                            THEN '${col}'
                            ELSE NULL 
                        END
                    WHEN t1."${col}" != t2."${col}" THEN '${col}'
                    ELSE NULL 
                END
            `;
        });

        const diffQuery = `
            WITH numbered_t1 AS (
                SELECT *, ROW_NUMBER() OVER() as row_num FROM ${table1Name}
            ),
            numbered_t2 AS (
                SELECT *, ROW_NUMBER() OVER() as row_num FROM ${table2Name}
            )
            SELECT 
                t1.row_num as row1,
                t2.row_num as row2,
                ${toleranceQueries.join(', ')} as diff_columns
            FROM numbered_t1 t1
            FULL OUTER JOIN numbered_t2 t2 ON t1.row_num = t2.row_num
            WHERE ${toleranceQueries.map((_, i) => `(${toleranceQueries[i]}) IS NOT NULL`).join(' OR ')}
        `;

        const result = await this.conn.query(diffQuery);
        return result.toArray();
    }

    async getTableData(tableName, limit = null, offset = 0) {
        if (!this.initialized) {
            throw new Error('DuckDB not initialized');
        }

        try {
            let query = `SELECT * FROM ${tableName}`;
            if (limit) {
                query += ` LIMIT ${limit} OFFSET ${offset}`;
            }
            
            const result = await this.conn.query(query);
            return result.toArray();
        } catch (error) {
            console.error(`Error getting data from ${tableName}:`, error);
            throw error;
        }
    }

    async close() {
        if (this.conn) {
            await this.conn.close();
            this.conn = null;
        }
        if (this.db) {
            await this.db.close();
            this.db = null;
        }
        this.initialized = false;
    }

    // Utility method to check if DuckDB is available
    static async isSupported() {
        try {
            // Check if WebAssembly is supported
            if (typeof WebAssembly !== 'object') {
                return false;
            }
            
            // Check if dynamic imports are supported by trying to import
            await import('data:text/javascript,export default 1');

            return true;
        } catch (error) {
            console.warn('Dynamic imports not supported:', error);
            return false;
        }
    }
}

// Global instance
let duckDBManager = null;

// Initialize DuckDB when this script loads
async function initializeDuckDB() {
    if (duckDBManager) {
        return duckDBManager;
    }

    const isSupported = await DuckDBManager.isSupported();
    if (!isSupported) {
        console.warn('DuckDB WASM not supported in this browser');
        return null;
    }

    try {
        duckDBManager = new DuckDBManager();
        const initialized = await duckDBManager.initialize();
        
        if (initialized) {
            console.log('DuckDB WASM ready for use');
            // Show a notification to user that fast mode is available
            showDuckDBStatus(true);
        } else {
            console.warn('DuckDB WASM initialization failed');
            duckDBManager = null;
            showDuckDBStatus(false);
        }
        
        return duckDBManager;
    } catch (error) {
        console.error('Error initializing DuckDB:', error);
        duckDBManager = null;
        showDuckDBStatus(false);
        window.duckDBAvailable = false;
        
        // Try to initialize local simulator as fallback
        console.log('🔄 Trying local DuckDB simulator as fallback...');
        if (typeof initializeLocalDuckDB === 'function') {
            try {
                const localManager = await initializeLocalDuckDB();
                if (localManager) {
                    showDuckDBStatus(true);
                    window.duckDBAvailable = true;
                }
            } catch (localError) {
                console.error('❌ Local simulator also failed:', localError);
            }
        }
        
        return null;
    }
}

function showDuckDBStatus(available) {
    // Add a subtle indicator to the UI
    const statusElement = document.getElementById('duckdb-status');
    if (statusElement) {
        if (available) {
            statusElement.innerHTML = '⚡ Fast comparison mode enabled - Up to 10x faster processing!';
            statusElement.className = 'duckdb-status duckdb-available show';
            
            // Show fast indicators on buttons
            const fastIndicators = document.querySelectorAll('.fast-mode-indicator');
            fastIndicators.forEach(indicator => {
                indicator.style.display = 'inline-block';
            });
            
            // Auto-hide after 5 seconds
            setTimeout(() => {
                if (statusElement.classList.contains('duckdb-available')) {
                    statusElement.style.opacity = '0.7';
                    statusElement.innerHTML = '⚡ Fast mode active';
                }
            }, 5000);
            
        } else {
            statusElement.innerHTML = '🔄 Standard comparison mode - DuckDB not available';
            statusElement.className = 'duckdb-status duckdb-unavailable show';
            
            // Auto-hide after 3 seconds
            setTimeout(() => {
                statusElement.style.display = 'none';
            }, 3000);
        }
    }
}

// Enhanced comparison function using DuckDB
async function compareTablesWithDuckDB(data1, data2, excludeColumns = [], useTolerance = false, tolerance = 1.5) {
    try {
        if (!duckDBManager || !duckDBManager.initialized) {
            // Fallback to original comparison
            console.log('DuckDB not available, using original comparison');
            return null;
        }

        console.log('Starting DuckDB fast comparison...');
        
        // Create tables
        await duckDBManager.createTableFromData('table1', data1);
        await duckDBManager.createTableFromData('table2', data2);
        
        // Perform fast comparison
        const comparisonResult = await duckDBManager.compareTablesFast(
            'table1', 'table2', excludeColumns
        );

        console.log('DuckDB comparison completed:', {
            identical: comparisonResult.identical.length,
            onlyInTable1: comparisonResult.onlyInTable1.length,
            onlyInTable2: comparisonResult.onlyInTable2.length
        });

        return comparisonResult;
        
    } catch (error) {
        console.error('DuckDB comparison failed:', error);
        return null; // Fallback to original method
    }
}

// Initialize DuckDB when the page loads
if (typeof window !== 'undefined') {
    // Make classes available globally for testing
    window.DuckDBManager = DuckDBManager;
    window.initializeDuckDB = initializeDuckDB;
    window.compareTablesWithDuckDB = compareTablesWithDuckDB;
    
    window.addEventListener('DOMContentLoaded', () => {
        console.log('🚀 Starting DuckDB initialization...');
        initializeDuckDB();
    });
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { DuckDBManager, initializeDuckDB, compareTablesWithDuckDB };
}

} // End of DuckDBManager class declaration guard
