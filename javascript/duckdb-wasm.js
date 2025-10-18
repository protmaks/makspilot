if (typeof window.MaxPilotDuckDB === 'undefined') {

// Database logging configuration
const DB_LOGGING_ENABLED = true; // Set to false to disable logging
const DB_LOGGING_LEVELS = {
    QUERY: true,      // Log individual SQL queries
    OPERATION: true,  // Log high-level operations (create table, comparison, etc.)
    TIMING: true,     // Log operation timings
    RESULTS: true,    // Log query results summary
    ZERO_VALUES: false // Log zero value formatting (can create many logs)
};
let dbOperationCounter = 0;

// Zero value formatting tracking
let zeroValueFormattingCount = 0;
const MAX_ZERO_VALUE_LOGS = 5; // Log only first 5 zero value formatting operations

// Database logging function
function logDatabaseOperation(operation, details = {}) {
    if (!DB_LOGGING_ENABLED) return;
    
    dbOperationCounter++;
    const timestamp = new Date().toISOString();
    const operationId = `DB_OP_${dbOperationCounter.toString().padStart(4, '0')}`;
    
    // Determine log level
    const isQuery = operation.toLowerCase().includes('query');
    const isOperation = operation.toLowerCase().includes('started') || operation.toLowerCase().includes('completed') || operation.toLowerCase().includes('failed');
    
    if (isQuery && !DB_LOGGING_LEVELS.QUERY) return;
    if (isOperation && !DB_LOGGING_LEVELS.OPERATION) return;
    
    console.group(`🗃️ [${operationId}] ${operation} - ${timestamp}`);
    
    if (details.query && DB_LOGGING_LEVELS.QUERY) {
        console.log('📝 SQL Query:', details.query);
    }
    if (details.parameters) {
        console.log('🔧 Parameters:', details.parameters);
    }
    if (details.tableName) {
        console.log('📊 Table:', details.tableName);
    }
    if (details.rowCount !== undefined) {
        console.log('📈 Row Count:', details.rowCount);
    }
    if (details.duration !== undefined && DB_LOGGING_LEVELS.TIMING) {
        console.log('⏱️ Duration:', `${details.duration}ms`);
    }
    if (details.error) {
        console.error('❌ Error:', details.error);
    }
    if (details.result && typeof details.result === 'object' && DB_LOGGING_LEVELS.RESULTS) {
        console.log('✅ Result:', details.result);
    }
    if (details.result && typeof details.result === 'string' && DB_LOGGING_LEVELS.RESULTS) {
        console.log('✅ Result:', details.result);
    }
    
    // Add performance metrics if available
    if (details.rowsPerSecond && DB_LOGGING_LEVELS.TIMING) {
        console.log('🚀 Performance:', `${details.rowsPerSecond} rows/sec`);
    }
    if (details.identicalRows !== undefined && DB_LOGGING_LEVELS.RESULTS) {
        console.log('🔄 Identical:', details.identicalRows);
    }
    if (details.onlyInTable1 !== undefined && DB_LOGGING_LEVELS.RESULTS) {
        console.log('📋 Only in Table 1:', details.onlyInTable1);
    }
    if (details.onlyInTable2 !== undefined && DB_LOGGING_LEVELS.RESULTS) {
        console.log('📄 Only in Table 2:', details.onlyInTable2);
    }
    
    console.groupEnd();
}

// Global logging control functions for browser console
window.enableDatabaseLogging = function() {
    DB_LOGGING_ENABLED = true;
    console.log('🗃️ Database logging enabled');
};

window.disableDatabaseLogging = function() {
    DB_LOGGING_ENABLED = false;
    console.log('🗃️ Database logging disabled');
};

window.setDatabaseLoggingLevel = function(level, enabled) {
    if (DB_LOGGING_LEVELS.hasOwnProperty(level)) {
        DB_LOGGING_LEVELS[level] = enabled;
        console.log(`🗃️ Database logging level ${level} ${enabled ? 'enabled' : 'disabled'}`);
    } else {
        console.log('🗃️ Available logging levels:', Object.keys(DB_LOGGING_LEVELS));
    }
};

window.showDatabaseLoggingStatus = function() {
    console.group('🗃️ Database Logging Status');
    console.log('Enabled:', DB_LOGGING_ENABLED);
    console.log('Operation counter:', dbOperationCounter);
    console.log('Zero value formattings:', zeroValueFormattingCount);
    console.log('Levels:', DB_LOGGING_LEVELS);
    console.groupEnd();
};

// Reset zero value formatting counter
window.resetZeroValueFormatting = function() {
    zeroValueFormattingCount = 0;
    console.log('🔄 Zero value formatting counter reset');
};

// Quick toggle for zero value logging
window.toggleZeroValueLogging = function() {
    DB_LOGGING_LEVELS.ZERO_VALUES = !DB_LOGGING_LEVELS.ZERO_VALUES;
    console.log(`🔧 Zero value logging ${DB_LOGGING_LEVELS.ZERO_VALUES ? 'enabled' : 'disabled'}`);
};

// Disable zero value logging immediately (for users experiencing spam)
window.disableZeroValueLogging = function() {
    DB_LOGGING_LEVELS.ZERO_VALUES = false;
    console.log('🚫 Zero value logging disabled to reduce log spam');
};

// Global functions for aggregation table management
window.getAggregationData = async function() {
    if (window.MaxPilotDuckDB && window.MaxPilotDuckDB.initialized) {
        return await window.MaxPilotDuckDB.getAggregationData();
    } else {
        console.warn('⚠️ MaxPilotDuckDB not initialized');
        return [];
    }
};

window.clearAggregationData = async function() {
    if (window.MaxPilotDuckDB && window.MaxPilotDuckDB.initialized) {
        await window.MaxPilotDuckDB.clearAggregationData();
        console.log('✅ Aggregation data cleared');
    } else {
        console.warn('⚠️ MaxPilotDuckDB not initialized');
    }
};

window.removeFileFromAggregation = async function(fileId) {
    if (window.MaxPilotDuckDB && window.MaxPilotDuckDB.initialized) {
        await window.MaxPilotDuckDB.removeFileFromAggregation(fileId);
        console.log('✅ File removed from aggregation:', fileId);
    } else {
        console.warn('⚠️ MaxPilotDuckDB not initialized');
    }
};

window.addFileToAggregation = async function(fileId, fileName, data, headers = null) {
    if (window.MaxPilotDuckDB && window.MaxPilotDuckDB.initialized) {
        return await window.MaxPilotDuckDB.addFileToAggregation(fileId, fileName, data, headers);
    } else {
        console.warn('⚠️ MaxPilotDuckDB not initialized');
        return null;
    }
};

// Progress display functions
let activeProgressIntervals = []; // For tracking active intervals
let progressState = {
    totalRows: 0,
    processedRows: 0,
    currentStage: '',
    stageProgress: 0
};

function createProgressIndicator(type = 'normal') {
    const isLarge = type === 'large';
    return `
        <div style="text-align: center; padding: 20px; font-family: Arial, sans-serif;">
            <div style="font-size: 16px; color: #333; margin-bottom: 15px;">
                ⚡ ${isLarge ? 'Processing large dataset' : 'Fast comparison mode'}
            </div>
            <div style="width: 100%; max-width: 400px; margin: 0 auto; background: #f0f0f0; border-radius: 10px; padding: 3px;">
                <div id="progress-bar" style="width: 0%; height: 20px; background: linear-gradient(90deg, #4CAF50, #45a049); border-radius: 8px; transition: width 0.3s ease;"></div>
            </div>
            <div id="progress-message" style="margin-top: 10px; font-size: 14px; color: #666;">
                Initializing...
            </div>
            <div id="progress-details" style="margin-top: 5px; font-size: 12px; color: #999;">
                ${isLarge ? 'Large datasets may take up to 2 minutes' : 'This should complete in seconds'}
            </div>
        </div>
    `;
}

function updateProgressMessage(message, progress = null) {
    const messageEl = document.getElementById('progress-message');
    const progressBar = document.getElementById('progress-bar');
    
    if (messageEl) {
        messageEl.textContent = message;
    }
    
    if (progress !== null && progressBar) {
        progressBar.style.width = `${Math.min(100, Math.max(0, progress))}%`;
    }
}

function updateProgressWithSteps(currentStep, totalSteps, stepName) {
    const progress = Math.round((currentStep / totalSteps) * 100);
    updateProgressMessage(`Step ${currentStep}/${totalSteps}: ${stepName}`, progress);
}

function simulateProgressDuringLongOperation(startProgress, endProgress, stepName, duration = 30000) {
    // Simulates smooth progress increase during long operations
    const startTime = Date.now();
    const progressRange = endProgress - startProgress;
    
    const progressInterval = setInterval(() => {
        const elapsed = Date.now() - startTime;
        const progressFactor = Math.min(elapsed / duration, 1);
        const currentProgress = startProgress + (progressRange * progressFactor);
        
        if (progressFactor < 1) {
            updateProgressMessage(`${stepName}... ${Math.round(elapsed / 1000)}s`, currentProgress);
        } else {
            clearInterval(progressInterval);
            activeProgressIntervals = activeProgressIntervals.filter(id => id !== progressInterval);
        }
    }, 1000);
    
    activeProgressIntervals.push(progressInterval);
    return progressInterval;
}

function clearAllProgressIntervals() {
    activeProgressIntervals.forEach(interval => clearInterval(interval));
    activeProgressIntervals = [];
}

function initializeProgress(totalRows, currentStage = 'Starting...') {
    progressState.totalRows = totalRows;
    progressState.processedRows = 0;
    progressState.currentStage = currentStage;
    progressState.stageProgress = 0;
    updateProgressFromState();
}

function updateRowProgress(processedRows, stageName = null) {
    progressState.processedRows = Math.min(processedRows, progressState.totalRows);
    if (stageName) {
        progressState.currentStage = stageName;
    }
    updateProgressFromState();
}

function updateStageProgress(stageName, stageProgress = 0) {
    progressState.currentStage = stageName;
    progressState.stageProgress = stageProgress;
    updateProgressFromState();
}

function updateProgressFromState() {
    const rowProgress = progressState.totalRows > 0 ? 
        (progressState.processedRows / progressState.totalRows) * 80 : 0; // 80% for row processing
    const stageProgress = progressState.stageProgress * 0.2; // 20% for stages
    const totalProgress = Math.min(100, rowProgress + stageProgress);
    
    const message = progressState.totalRows > 0 ? 
        `${progressState.currentStage} (${progressState.processedRows.toLocaleString()}/${progressState.totalRows.toLocaleString()} rows)` :
        progressState.currentStage;
    
    updateProgressMessage(message, totalProgress);
}

class FastTableComparator {
    constructor() {
        this.initialized = false;
        this.tables = new Map();
        this.mode = 'local';
    }

    async initialize() {
        try {
            
            if (typeof WebAssembly !== 'undefined') {
                try {
                    await this.initializeWASM();
                    this.mode = 'wasm';
                    this.initialized = true;
                    return true;
                }
                catch (wasmError) { console.log('📝 DuckDB WASM not available, falling back to optimized local mode'); }
            }
            
            // Fallback to local mode
            await this.initializeLocal();
            this.mode = 'local';
            this.initialized = true;
            
            return true;
            
        } catch (error) {
            console.error('❌ Failed to initialize comparison engine:', error);
            this.initialized = false;
            return false;
        }
    }

    async initializeWASM() {
        try {
            const script = document.createElement('script');
            script.src = '/javascript/duckdb/duckdb-real.js';
            script.type = 'text/javascript';
            
            await new Promise((resolve, reject) => {
                script.onload = resolve;
                script.onerror = reject;
                document.head.appendChild(script);
            });
            
            await new Promise(resolve => setTimeout(resolve, 100));
            
            if (!window.duckdbLoader) {
                throw new Error('DuckDB loader not available');
            }
            
            this.db = await window.duckdbLoader.initialize();
            
            // Create logged wrapper for database queries
            this.originalQuery = window.duckdbLoader.query.bind(window.duckdbLoader);
            window.duckdbLoader.query = this.createLoggedQuery();
            
            // Initialize aggregation table for file statistics
            await this.initializeAggregationTable();
            
            logDatabaseOperation('DuckDB WASM Initialized', {
                result: 'Database successfully initialized with aggregation table'
            });
            
            return true;
            
        } catch (error) {
            console.error('❌ Real DuckDB WASM initialization failed:', error);
            throw new Error('Real DuckDB WASM not available');
        }
    }

    createLoggedQuery() {
        return async (query, parameters = null) => {
            const startTime = Date.now();
            
            logDatabaseOperation('Query Started', {
                query: query,
                parameters: parameters
            });
            
            try {
                const result = await this.originalQuery(query, parameters);
                const duration = Date.now() - startTime;
                
                logDatabaseOperation('Query Completed', {
                    duration: duration,
                    result: result ? `${result.length || 0} rows returned` : 'No result data'
                });
                
                return result;
                
            } catch (error) {
                const duration = Date.now() - startTime;
                
                logDatabaseOperation('Query Failed', {
                    query: query,
                    duration: duration,
                    error: error.message || error
                });
                
                throw error;
            }
        };
    }

    async initializeLocal() {
        this.tables.clear();
        return true;
    }

    async initializeAggregationTable() {
        try {
            const createAggregationTableSQL = `
                CREATE OR REPLACE TABLE file_aggregation_stats (
                    file_id VARCHAR PRIMARY KEY,
                    file_name VARCHAR NOT NULL,
                    row_count INTEGER NOT NULL DEFAULT 0,
                    column_count INTEGER NOT NULL DEFAULT 0,
                    identical_rows INTEGER NOT NULL DEFAULT 0,
                    rows_only_in_file INTEGER NOT NULL DEFAULT 0,
                    similar_rows INTEGER NOT NULL DEFAULT 0,
                    diff_columns INTEGER NOT NULL DEFAULT 0,
                    similarity_percent DOUBLE NOT NULL DEFAULT 0.0,
                    is_compared BOOLEAN NOT NULL DEFAULT FALSE,
                    loaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    compared_at TIMESTAMP,
                    file_size_bytes INTEGER DEFAULT 0
                )
            `;
            
            await window.duckdbLoader.query(createAggregationTableSQL);
            
            logDatabaseOperation('Aggregation Table Created', {
                tableName: 'file_aggregation_stats',
                purpose: 'Store file statistics and comparison results'
            });
            
            return true;
            
        } catch (error) {
            console.error('❌ Failed to create aggregation table:', error);
            throw error;
        }
    }

    async addFileToAggregation(fileId, fileName, data, headers = null) {
        try {
            const rowCount = data ? data.length - 1 : 0; // Exclude header row
            const columnCount = headers ? headers.length : (data[0] ? data[0].length : 0);
            
            // Calculate approximate file size in bytes (rough estimation)
            const fileSize = this.estimateDataSizeInBytes(data);
            
            const insertSQL = `
                INSERT OR REPLACE INTO file_aggregation_stats 
                (file_id, file_name, row_count, column_count, file_size_bytes)
                VALUES ('${fileId}', '${fileName.replace(/'/g, "''")}', ${rowCount}, ${columnCount}, ${fileSize})
            `;
            
            await window.duckdbLoader.query(insertSQL);
            
            logDatabaseOperation('File Added to Aggregation', {
                fileId: fileId,
                fileName: fileName,
                rowCount: rowCount,
                columnCount: columnCount,
                fileSize: fileSize
            });
            
            // Update UI with current aggregation data
            await this.updateAggregationUI();
            
            return {
                fileId,
                fileName,
                rowCount,
                columnCount,
                fileSize
            };
            
        } catch (error) {
            console.error('❌ Failed to add file to aggregation:', error);
            throw error;
        }
    }

    estimateDataSizeInBytes(data) {
        if (!data || data.length === 0) return 0;
        
        let totalSize = 0;
        const sampleSize = Math.min(100, data.length); // Sample first 100 rows for estimation
        
        for (let i = 0; i < sampleSize; i++) {
            if (data[i] && Array.isArray(data[i])) {
                for (let j = 0; j < data[i].length; j++) {
                    const cellValue = data[i][j];
                    if (cellValue !== null && cellValue !== undefined) {
                        totalSize += cellValue.toString().length * 2; // Estimate 2 bytes per character
                    }
                }
            }
        }
        
        // Extrapolate to full dataset
        const avgRowSize = totalSize / sampleSize;
        return Math.round(avgRowSize * data.length);
    }

    async updateComparisonResults(file1Id, file2Id, comparisonResults) {
        try {
            const { identical, similar, onlyInTable1, onlyInTable2, commonColumns } = comparisonResults;
            
            // Calculate statistics for file 1
            const file1Stats = {
                identicalRows: identical.length,
                rowsOnlyInFile: onlyInTable1.length,
                similarRows: similar ? similar.length : 0,
                diffColumns: this.calculateDifferentColumns(comparisonResults, 1),
                similarityPercent: this.calculateSimilarityPercent(comparisonResults, 1)
            };
            
            // Calculate statistics for file 2
            const file2Stats = {
                identicalRows: identical.length,
                rowsOnlyInFile: onlyInTable2.length,
                similarRows: similar ? similar.length : 0,
                diffColumns: this.calculateDifferentColumns(comparisonResults, 2),
                similarityPercent: this.calculateSimilarityPercent(comparisonResults, 2)
            };
            
            // Update file 1 statistics
            await window.duckdbLoader.query(`
                UPDATE file_aggregation_stats 
                SET identical_rows = ${file1Stats.identicalRows}, 
                    rows_only_in_file = ${file1Stats.rowsOnlyInFile}, 
                    similar_rows = ${file1Stats.similarRows},
                    diff_columns = ${file1Stats.diffColumns}, 
                    similarity_percent = ${file1Stats.similarityPercent}, 
                    is_compared = TRUE, 
                    compared_at = CURRENT_TIMESTAMP
                WHERE file_id = '${file1Id}'
            `);
            
            // Update file 2 statistics
            await window.duckdbLoader.query(`
                UPDATE file_aggregation_stats 
                SET identical_rows = ${file2Stats.identicalRows}, 
                    rows_only_in_file = ${file2Stats.rowsOnlyInFile}, 
                    similar_rows = ${file2Stats.similarRows},
                    diff_columns = ${file2Stats.diffColumns}, 
                    similarity_percent = ${file2Stats.similarityPercent}, 
                    is_compared = TRUE, 
                    compared_at = CURRENT_TIMESTAMP
                WHERE file_id = '${file2Id}'
            `);
            
            logDatabaseOperation('Comparison Results Updated', {
                file1Id: file1Id,
                file2Id: file2Id,
                file1Stats: file1Stats,
                file2Stats: file2Stats
            });
            
            // Update UI with new aggregation data
            await this.updateAggregationUI();
            
            return { file1Stats, file2Stats };
            
        } catch (error) {
            console.error('❌ Failed to update comparison results:', error);
            throw error;
        }
    }

    calculateDifferentColumns(comparisonResults, fileNumber) {
        // Calculate number of columns that have differences
        const { commonColumns, comparisonColumns } = comparisonResults;
        
        if (!comparisonColumns || comparisonColumns.length === 0) {
            return 0;
        }
        
        // For now, return the number of comparison columns
        // This could be enhanced to actually count columns with differences
        return comparisonColumns.length;
    }

    calculateSimilarityPercent(comparisonResults, fileNumber) {
        const { identical, similar, onlyInTable1, onlyInTable2, table1Count, table2Count } = comparisonResults;
        
        const totalRows = fileNumber === 1 ? table1Count : table2Count;
        if (totalRows === 0) return 0;
        
        const identicalCount = identical ? identical.length : 0;
        const similarCount = similar ? similar.length : 0;
        
        // Calculate similarity as percentage of identical + similar rows
        const matchingRows = identicalCount + (similarCount * 0.5); // Similar rows count as 50%
        const similarityPercent = (matchingRows / totalRows) * 100;
        
        return Math.round(similarityPercent * 100) / 100; // Round to 2 decimal places
    }

    async getAggregationData() {
        try {
            // First check if table exists and has data
            const countResult = await window.duckdbLoader.query(`SELECT COUNT(*) as total FROM file_aggregation_stats`);
            const countData = countResult.toArray();
            const count = Number(countData[0]?.total || 0);
            
            const result = await window.duckdbLoader.query(`
                SELECT 
                    file_id,
                    file_name,
                    row_count,
                    column_count,
                    identical_rows,
                    rows_only_in_file,
                    similar_rows,
                    diff_columns,
                    similarity_percent,
                    is_compared,
                    loaded_at,
                    compared_at,
                    file_size_bytes
                FROM file_aggregation_stats 
                ORDER BY loaded_at DESC
            `);
            
            const data = result.toArray();
            
            // Convert BigInt values to regular numbers for JSON serialization
            const convertedData = data.map(row => {
                const convertedRow = {};
                for (const [key, value] of Object.entries(row)) {
                    if (typeof value === 'bigint') {
                        convertedRow[key] = Number(value);
                    } else {
                        convertedRow[key] = value;
                    }
                }
                return convertedRow;
            });
            
            console.log(`📊 Aggregation table updated: ${convertedData.length} files loaded`);
            
            return convertedData;
            
        } catch (error) {
            console.error('❌ Failed to get aggregation data:', error);
            return [];
        }
    }

    async updateAggregationUI() {
        try {
            const aggregationData = await this.getAggregationData();
            
            // UI update with aggregation data
            
            // Dispatch custom event with aggregation data
            const event = new CustomEvent('aggregationDataUpdated', {
                detail: { data: aggregationData }
            });
            window.dispatchEvent(event);
            
            // Also call global function if it exists
            if (typeof window.updateAggregationTable === 'function') {
                window.updateAggregationTable(aggregationData);
            } else {
                // Create a temporary function for aggregation data
                if (!window.updateAggregationTable) {
                    window.updateAggregationTable = function(data) {
                        console.log('� Aggregation UI updated:', data.length > 0 ? `${data.length} files` : 'no data');
                    };
                }
                window.updateAggregationTable(aggregationData);
            }
            
        } catch (error) {
            console.error('❌ Failed to update aggregation UI:', error);
        }
    }

    async clearAggregationData() {
        try {
            await window.duckdbLoader.query('DELETE FROM file_aggregation_stats');
            
            logDatabaseOperation('Aggregation Data Cleared', {
                reason: 'User requested clear or system reset'
            });
            
            await this.updateAggregationUI();
            
        } catch (error) {
            console.error('❌ Failed to clear aggregation data:', error);
            throw error;
        }
    }

    async removeFileFromAggregation(fileId) {
        try {
            await window.duckdbLoader.query(`DELETE FROM file_aggregation_stats WHERE file_id = '${fileId}'`);
            
            logDatabaseOperation('File Removed from Aggregation', {
                fileId: fileId
            });
            
            await this.updateAggregationUI();
            
        } catch (error) {
            console.error('❌ Failed to remove file from aggregation:', error);
            throw error;
        }
    }

    async createTableFromData(tableName, data, headers = null, fileName = null) {
        if (!this.initialized) {
            throw new Error('Comparator not initialized');
        }

        logDatabaseOperation('Create Table Started', {
            tableName: tableName,
            rowCount: data ? data.length : 0,
            headers: headers,
            fileName: fileName
        });

        try {
            if (this.mode === 'wasm') {
                // For WASM mode, add to aggregation table immediately
                const fileId = tableName; // Use table name as file ID
                const displayName = fileName || tableName;
                
                await this.addFileToAggregation(fileId, displayName, data, headers);
            }
            
            const processedData = this.processTableData(data, headers);
            this.tables.set(tableName, processedData);
            
            logDatabaseOperation('Create Table Completed', {
                tableName: tableName,
                rowCount: processedData.rows.length,
                columnCount: processedData.columns.length,
                fileName: fileName
            });
            
            return true;
            
        } catch (error) {
            logDatabaseOperation('Create Table Failed', {
                tableName: tableName,
                error: error.message || error
            });
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

    async compareTablesFast(data1, data2, excludeColumns = [], useTolerance = false, customKeyColumns = null) {
        if (this.mode === 'wasm') {
            // For WASM mode, use the optimized table comparison if data is already in table format
            if (typeof data1 === 'string' && typeof data2 === 'string') {
                // data1 and data2 are table names, use DuckDB comparison
                logDatabaseOperation('Fast Comparison: Using DuckDB table comparison', {
                    table1: data1,
                    table2: data2,
                    excludeColumns: excludeColumns.length
                });
                return await this.compareTablesWithDuckDB(data1, data2, excludeColumns);
            } else {
                // data1 and data2 are raw data arrays
                // Check if data is large enough to benefit from full DuckDB logic
                if (this.shouldUseFullDuckDBLogic(data1, data2)) {
                    logDatabaseOperation('Fast Comparison: Using full DuckDB logic for large dataset', {
                        data1Rows: data1 ? data1.length : 0,
                        data2Rows: data2 ? data2.length : 0
                    });
                    return await this.compareTablesWithOriginalLogic(data1, data2, excludeColumns, useTolerance, customKeyColumns);
                } else {
                    // For smaller datasets, use local comparison for speed
                    logDatabaseOperation('Fast Comparison: Using local comparison for small dataset', {
                        data1Rows: data1 ? data1.length : 0,
                        data2Rows: data2 ? data2.length : 0,
                        useTolerance,
                        customKeyColumns: customKeyColumns ? customKeyColumns.length : 0
                    });
                    return await this.compareTablesLocal(data1, data2, excludeColumns, useTolerance, customKeyColumns);
                }
            }
        } else {
            logDatabaseOperation('Fast Comparison: Using local mode', {
                data1Type: typeof data1,
                data2Type: typeof data2
            });
            return await this.compareTablesLocal(data1, data2, excludeColumns, useTolerance, customKeyColumns);
        }
    }

    async compareTablesWithDuckDB(table1Name, table2Name, excludeColumns = []) {
        const startTime = performance.now();

        try {
            logDatabaseOperation('DuckDB Table Comparison Started', {
                table1Name,
                table2Name,
                excludeColumns
            });

            const table1 = this.tables.get(table1Name);
            const table2 = this.tables.get(table2Name);
            
            if (!table1 || !table2) {
                throw new Error(`One or both tables not found: table1=${!!table1}, table2=${!!table2}`);
            }

            logDatabaseOperation('Creating temporary DuckDB tables', {
                table1Columns: table1.columns.length,
                table2Columns: table2.columns.length,
                table1Rows: table1.rows.length,
                table2Rows: table2.rows.length
            });

            // Drop any existing temporary tables first to ensure clean state
            try {
                await window.duckdbLoader.query('DROP TABLE IF EXISTS temp_table1');
                await window.duckdbLoader.query('DROP TABLE IF EXISTS temp_table2');
                logDatabaseOperation('Cleaned up existing temporary tables');
            } catch (cleanupError) {
                logDatabaseOperation('Cleanup warning (tables may not exist)', {
                    error: cleanupError.message
                });
            }

            await this.createDuckDBTable('temp_table1', table1);
            await this.createDuckDBTable('temp_table2', table2);

            const commonColumns = table1.columns.filter(col => 
                table2.columns.includes(col) && 
                !excludeColumns.some(excCol => 
                    col.toLowerCase().includes(excCol.toLowerCase())
                )
            );

            if (commonColumns.length === 0) {
                throw new Error('No common columns found for comparison');
            }

            // Verify columns exist in both created tables
            try {
                const table1Desc = await window.duckdbLoader.query('DESCRIBE temp_table1');
                const table2Desc = await window.duckdbLoader.query('DESCRIBE temp_table2');
                
                const table1ActualColumns = table1Desc.toArray().map(row => row.column_name).filter(col => col !== 'rowid');
                const table2ActualColumns = table2Desc.toArray().map(row => row.column_name).filter(col => col !== 'rowid');
                
                // Filter common columns to only include those that actually exist in both DB tables
                const verifiedCommonColumns = commonColumns.filter(col => 
                    table1ActualColumns.includes(col) && table2ActualColumns.includes(col)
                );
                
                if (verifiedCommonColumns.length === 0) {
                    throw new Error('No verified common columns found in created tables');
                }
                
                if (verifiedCommonColumns.length !== commonColumns.length) {
                    logDatabaseOperation('Column verification warning', {
                        expectedColumns: commonColumns,
                        verifiedColumns: verifiedCommonColumns,
                        table1ActualColumns,
                        table2ActualColumns
                    });
                }
                
                // Update commonColumns to use only verified columns
                commonColumns.length = 0;
                commonColumns.push(...verifiedCommonColumns);
                
            } catch (verifyError) {
                logDatabaseOperation('Column verification failed', {
                    error: verifyError.message
                });
                // Continue with original commonColumns if verification fails
            }

            logDatabaseOperation('Common columns identified', {
                commonColumns: commonColumns.length,
                columnList: commonColumns
            });

            // Create SELECT for common columns
            const columnList = commonColumns.map(col => `"${col}"`).join(', ');
            
            // Create WHERE clause for all common columns using simple equality comparison
            const whereConditions = commonColumns.map(col => `t1."${col}" = t2."${col}"`);
            const whereClause = whereConditions.join(' AND ');

            // SQL for finding identical rows
            const identicalQuery = `
                SELECT t1.rowid as row1, t2.rowid as row2, 'IDENTICAL' as status
                FROM temp_table1 t1
                INNER JOIN temp_table2 t2 ON ${whereClause}
            `;

            // SQL for rows only in table 1
            const onlyInTable1Query = `
                SELECT t1.rowid as row1, NULL as row2, 'ONLY_IN_TABLE1' as status
                FROM temp_table1 t1
                LEFT JOIN temp_table2 t2 ON ${whereClause}
                WHERE t2.rowid IS NULL
            `;

            // SQL for rows only in table 2
            const onlyInTable2Query = `
                SELECT NULL as row1, t2.rowid as row2, 'ONLY_IN_TABLE2' as status
                FROM temp_table2 t2
                LEFT JOIN temp_table1 t1 ON ${whereClause}
                WHERE t1.rowid IS NULL
            `;

            logDatabaseOperation('Executing DuckDB comparison queries', {
                query: 'Parallel execution of 3 comparison queries',
                identicalQuery: identicalQuery.replace(/\s+/g, ' ').trim(),
                onlyInTable1Query: onlyInTable1Query.replace(/\s+/g, ' ').trim(),
                onlyInTable2Query: onlyInTable2Query.replace(/\s+/g, ' ').trim()
            });

            // Execute queries in parallel with individual error handling
            let identicalResult, onlyTable1Result, onlyTable2Result;
            
            try {
                [identicalResult, onlyTable1Result, onlyTable2Result] = await Promise.all([
                    window.duckdbLoader.query(identicalQuery).catch(err => {
                        throw new Error(`Identical query failed: ${err.message}`);
                    }),
                    window.duckdbLoader.query(onlyInTable1Query).catch(err => {
                        throw new Error(`OnlyInTable1 query failed: ${err.message}`);
                    }),
                    window.duckdbLoader.query(onlyInTable2Query).catch(err => {
                        throw new Error(`OnlyInTable2 query failed: ${err.message}`);
                    })
                ]);
            } catch (queryError) {
                logDatabaseOperation('DuckDB Query Execution Failed', {
                    error: queryError.message,
                    identicalQuery,
                    onlyInTable1Query,
                    onlyInTable2Query
                });
                throw queryError;
            }

            // Convert results
            const identical = identicalResult.toArray().map(row => ({
                row1: row.row1,
                row2: row.row2,
                status: 'IDENTICAL'
            }));

            const onlyInTable1 = onlyTable1Result.toArray().map(row => ({
                row1: row.row1,
                row2: null,
                status: 'ONLY_IN_TABLE1'
            }));

            const onlyInTable2 = onlyTable2Result.toArray().map(row => ({
                row1: null,
                row2: row.row2,
                status: 'ONLY_IN_TABLE2'
            }));

            const duration = performance.now() - startTime;

            logDatabaseOperation('DuckDB Table Comparison Completed', {
                identicalRows: identical.length,
                onlyInTable1: onlyInTable1.length,
                onlyInTable2: onlyInTable2.length,
                commonColumns: commonColumns.length,
                duration: Math.round(duration)
            });

            return {
                identical,
                onlyInTable1,
                onlyInTable2,
                commonColumns,
                table1Count: table1.rows.length,
                table2Count: table2.rows.length,
                duration,
                mode: 'duckdb-wasm'
            };

        } catch (error) {
            const duration = performance.now() - startTime;
            
            logDatabaseOperation('DuckDB Table Comparison Failed', {
                error: error.message,
                duration: Math.round(duration),
                tableName: table1Name && table2Name ? `${table1Name} vs ${table2Name}` : 'unknown'
            });
            
            console.error('❌ DuckDB comparison failed, falling back to local mode:', error);
            return await this.compareTablesLocal(table1Name, table2Name, excludeColumns);
        }
    }

    // Helper method to determine if full DuckDB logic is needed based on data size
    shouldUseFullDuckDBLogic(data1, data2) {
        const threshold = 10000; // Use full logic for datasets larger than 10k rows
        const rows1 = data1 ? (Array.isArray(data1) ? data1.length - 1 : 0) : 0;
        const rows2 = data2 ? (Array.isArray(data2) ? data2.length - 1 : 0) : 0;
        const totalRows = rows1 + rows2;
        
        logDatabaseOperation('Data size check for comparison method', {
            rows1,
            rows2,
            totalRows,
            threshold,
            useFullLogic: totalRows > threshold
        });
        
        return totalRows > threshold;
    }

    // Full DuckDB comparison with table creation and complex SQL logic
    // Use this for thorough comparison when tables need to be created from scratch
    // For quick comparisons, prefer compareTablesFast() or compareTablesWithDuckDB()
    async compareTablesWithOriginalLogic(data1, data2, excludeColumns = [], useTolerance = false, customKeyColumns = null) {
        const startTime = performance.now();

        // Reset zero value formatting counter for this comparison
        zeroValueFormattingCount = 0;

        logDatabaseOperation('Original Logic Comparison Started', {
            table1Rows: data1 ? data1.length - 1 : 0,
            table2Rows: data2 ? data2.length - 1 : 0,
            excludeColumns: excludeColumns,
            useTolerance: useTolerance,
            customKeyColumns: customKeyColumns
        });

        try {
            const table1Size = data1.length - 1;
            const table2Size = data2.length - 1;
            const totalSize = table1Size + table2Size;
            const columnCount = Math.max(data1[0]?.length || 0, data2[0]?.length || 0);
            
            // Initialize progress tracking
            initializeProgress(totalSize, 'Initializing comparison...');
            
            updateStageProgress('Validating input data', 5);

            if (!Array.isArray(data1) || !Array.isArray(data2) || data1.length === 0 || data2.length === 0) {
                console.error('❌ Data validation failed:', {
                    data1IsArray: Array.isArray(data1),
                    data2IsArray: Array.isArray(data2),
                    data1Length: data1?.length,
                    data2Length: data2?.length
                });
                throw new Error('Invalid input data');
            }

            updateStageProgress('Analyzing column structure', 10);

            const headers1 = data1[0] || [];
            const headers2 = data2[0] || [];
            
            const sanitizeColumnName = (name, index) => {
                if (!name || typeof name !== 'string') {
                    return `col_${index}`;
                }
                let sanitized = name.replace(/[^a-zA-Z0-9а-яА-Я_]/g, '_')
                                   .replace(/\s+/g, '_')
                                   .replace(/_{2,}/g, '_')
                                   .replace(/^_|_$/g, '');
                
                if (!sanitized || /^\d/.test(sanitized)) {
                    sanitized = `col_${index}_${sanitized}`;
                }
                
                return sanitized || `col_${index}`;
            };

            const detectColumnType = (data, columnIndex) => {
                const sampleSize = Math.min(100, data.length - 1);
                let numericCount = 0;
                let integerCount = 0;
                let dateCount = 0;
                let timestampCount = 0;
                let totalNonEmpty = 0;
                let hasDecimals = false;

                for (let i = 1; i <= sampleSize; i++) {
                    if (i >= data.length) break;
                    
                    const value = data[i]?.[columnIndex];
                    if (value && value.toString().trim() !== '') {
                        totalNonEmpty++;
                        const strValue = value.toString().trim();
                        
                        if (!isNaN(strValue) && !isNaN(parseFloat(strValue)) && isFinite(strValue)) {
                            numericCount++;
                            const numValue = parseFloat(strValue);
                            
                            if (Number.isInteger(numValue)) {
                                integerCount++;
                            } else {
                                hasDecimals = true;
                            }
                        }
                        
                        const dateValue = new Date(strValue);
                        if (!isNaN(dateValue.getTime()) && strValue.match(/\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}|\d{2}\.\d{2}\.\d{4}/)) {
                            // Check if the string contains time information
                            if (strValue.match(/\d{2}:\d{2}(:\d{2})?/) || strValue.includes('T')) {
                                timestampCount++;
                            } else {
                                dateCount++;
                            }
                        }
                    }
                }

                if (totalNonEmpty === 0) return 'VARCHAR';
                
                const numericRatio = numericCount / totalNonEmpty;
                const dateRatio = dateCount / totalNonEmpty;
                const timestampRatio = timestampCount / totalNonEmpty;

                if (timestampRatio >= 0.8) return 'TIMESTAMP';
                if (dateRatio >= 0.8) { return 'DATE'; }
                
                if (numericRatio >= 0.9) {
                    if (hasDecimals) {
                        return 'DOUBLE';
                    } else {
                        return 'BIGINT';
                    }
                }
                
                return 'VARCHAR';
            };

            const sanitizedHeaders1 = headers1.map((h, i) => sanitizeColumnName(h, i));
            const sanitizedHeaders2 = headers2.map((h, i) => sanitizeColumnName(h, i));

            updateStageProgress('Detecting column types', 20);

            const columnTypes1 = sanitizedHeaders1.map((_, i) => detectColumnType(data1, i));
            const columnTypes2 = sanitizedHeaders2.map((_, i) => detectColumnType(data2, i));

            const harmonizeColumnTypes = (types1, types2) => {
                const harmonized1 = [...types1];
                const harmonized2 = [...types2];
                
                for (let i = 0; i < Math.min(types1.length, types2.length); i++) {
                    const type1 = types1[i];
                    const type2 = types2[i];
                    
                    if (type1 !== type2) {
                        if ((type1 === 'BIGINT' && type2 === 'DOUBLE') || 
                            (type1 === 'DOUBLE' && type2 === 'BIGINT')) {
                            harmonized1[i] = 'DOUBLE';
                            harmonized2[i] = 'DOUBLE';
                        }
                        else if ((type1 === 'INTEGER' && type2 === 'BIGINT') || 
                                 (type1 === 'BIGINT' && type2 === 'INTEGER')) {
                            harmonized1[i] = 'BIGINT';
                            harmonized2[i] = 'BIGINT';
                        }
                        else if ((type1 === 'INTEGER' && type2 === 'DOUBLE') || 
                                 (type1 === 'DOUBLE' && type2 === 'INTEGER') ||
                                 (type1 === 'FLOAT' && type2 === 'DOUBLE') || 
                                 (type1 === 'DOUBLE' && type2 === 'FLOAT')) {
                            harmonized1[i] = 'DOUBLE';
                            harmonized2[i] = 'DOUBLE';
                        }
                        else if ((type1 === 'DATE' && type2 === 'TIMESTAMP') || 
                                 (type1 === 'TIMESTAMP' && type2 === 'DATE')) {
                            // If one is DATE and other is TIMESTAMP, use TIMESTAMP to preserve time info
                            harmonized1[i] = 'TIMESTAMP';
                            harmonized2[i] = 'TIMESTAMP';
                        }
                        else {
                            harmonized1[i] = 'VARCHAR';
                            harmonized2[i] = 'VARCHAR';
                        }
                    }
                }
                
                return { types1: harmonized1, types2: harmonized2 };
            };

            const { types1: finalColumnTypes1, types2: finalColumnTypes2 } = harmonizeColumnTypes(columnTypes1, columnTypes2);

            updateStageProgress('Creating database tables', 30);

            // Check if DuckDB tables already exist (created from preview)
            let tablesExist = false;
            try {
                const table1CountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM table1');
                const table2CountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM table2');
                
                const table1Count = Number(table1CountResult.toArray()[0]?.count || 0);
                const table2Count = Number(table2CountResult.toArray()[0]?.count || 0);
                
                // Tables exist and have data
                if (table1Count > 0 && table2Count > 0) {
                    tablesExist = true;
                    console.log(`✅ Using existing DuckDB tables from preview (table1: ${table1Count} rows, table2: ${table2Count} rows)`);
                    
                    logDatabaseOperation('Using Existing DuckDB Tables', {
                        table1: 'table1',
                        table2: 'table2',
                        table1Count: table1Count,
                        table2Count: table2Count,
                        reason: 'Tables already exist from preview with data'
                    });
                } else {
                    tablesExist = false;
                    console.log('📋 Tables exist but are empty, will recreate');
                    
                    // Drop empty tables
                    await window.duckdbLoader.query('DROP TABLE IF EXISTS table1');
                    await window.duckdbLoader.query('DROP TABLE IF EXISTS table2');
                }
            } catch (error) {
                // Tables don't exist, will create them
                tablesExist = false;
                console.log('📋 Creating new DuckDB tables for comparison');
            }

            if (!tablesExist) {
                const createTable1SQL = `CREATE TABLE table1 (
                    rowid INTEGER,
                    ${sanitizedHeaders1.map((h, i) => `"${h}" ${finalColumnTypes1[i]}`).join(', ')}
                )`;

                const createTable2SQL = `CREATE TABLE table2 (
                    rowid INTEGER,
                    ${sanitizedHeaders2.map((h, i) => `"${h}" ${finalColumnTypes2[i]}`).join(', ')}
                )`;

                await window.duckdbLoader.query(createTable1SQL);
                await window.duckdbLoader.query(createTable2SQL);

                updateStageProgress('Loading data into tables', 40);


                const insertBatch = async (tableName, data, headers, columnTypes) => {
                    // Подготавливаем все значения одним блоком
                    const allValues = [];
                    for (let i = 1; i < data.length; i++) {
                        const row = data[i];
                        const rowId = i - 1;
                        const cleanRow = headers.map((_, colIdx) => {
                            const val = row[colIdx];
                            return formatValue(val, columnTypes[colIdx]);
                        }).join(', ');
                        allValues.push(`(${rowId}, ${cleanRow})`);
                    }

                    if (allValues.length > 0) {
                        const insertSQL = `INSERT INTO ${tableName} VALUES ${allValues.join(', ')}`;
                        await window.duckdbLoader.query(insertSQL);
                    }
                };

                await insertBatch('table1', data1, headers1, finalColumnTypes1);
                await insertBatch('table2', data2, headers2, finalColumnTypes2);
            } else {
                updateStageProgress('Using existing DuckDB tables', 40);
            }
            
            updateStageProgress('Preparing comparison logic', 50);
            
            const createComparisonCondition = (colIdx, useTolerance = false, isKeyColumn = false) => {
                const col1Type = finalColumnTypes1[colIdx];
                const col2Type = finalColumnTypes2[colIdx];
                const col1Name = sanitizedHeaders1[colIdx];
                const col2Name = sanitizedHeaders2[colIdx];
                
                // Safety check: ensure both column names exist and are within bounds
                if (!col1Name || !col2Name || colIdx >= sanitizedHeaders1.length || colIdx >= sanitizedHeaders2.length) {
                    // Skip this comparison if column is missing or out of bounds
                    return 'FALSE'; 
                }
                
                if (col1Type === 'DOUBLE' || col2Type === 'DOUBLE') {
                    if (useTolerance) {
                        // Apply 1.5% tolerance for DOUBLE values
                        const tolerancePercent = (window.currentTolerance || 1.5) / 100;
                        return `(
                            (t1."${col1Name}" = 0 AND t2."${col2Name}" = 0) OR
                            (t1."${col1Name}" = t2."${col2Name}") OR
                            (t1."${col1Name}" != 0 AND t2."${col2Name}" != 0 AND 
                             ABS(t1."${col1Name}" - t2."${col2Name}") / ((ABS(t1."${col1Name}") + ABS(t2."${col2Name}")) / 2) <= ${tolerancePercent})
                        )`;
                    } else {
                        // Handle double comparison with special case for zeros
                        return `(
                            (t1."${col1Name}" = 0 AND t2."${col2Name}" = 0) OR
                            (t1."${col1Name}" != 0 AND t2."${col2Name}" != 0 AND ROUND(t1."${col1Name}", 2) = ROUND(t2."${col2Name}", 2)) OR
                            (t1."${col1Name}" = t2."${col2Name}")
                        )`;
                    }
                }
                
                if (useTolerance && (col1Type === 'BIGINT' || col1Type === 'INTEGER' || col1Type === 'FLOAT')) {
                    const tolerancePercent = (window.currentTolerance || 1.5) / 100;
                    return `(
                        (t1."${col1Name}" = 0 AND t2."${col2Name}" = 0) OR
                        (t1."${col1Name}" = t2."${col2Name}") OR
                        (t1."${col1Name}" != 0 AND t2."${col2Name}" != 0 AND 
                         ABS(t1."${col1Name}" - t2."${col2Name}") / ((ABS(t1."${col1Name}") + ABS(t2."${col2Name}")) / 2) <= ${tolerancePercent})
                    )`;
                }
                
                // Handle DATE and TIMESTAMP types in tolerance mode
                if (useTolerance && (col1Type === 'DATE' || col2Type === 'DATE' || col1Type === 'TIMESTAMP' || col2Type === 'TIMESTAMP')) {
                    if (isKeyColumn) {
                        // For key columns, compare full datetime even in tolerance mode
                        return `t1."${col1Name}" = t2."${col2Name}"`;
                    } else {
                        // For regular columns, compare only date part
                        return `strftime(t1."${col1Name}", '%Y-%m-%d') = strftime(t2."${col2Name}", '%Y-%m-%d')`;
                    }
                }
                
                if (col1Type === 'BIGINT' || col1Type === 'INTEGER' || col1Type === 'FLOAT') {
                    return `t1."${col1Name}" = t2."${col2Name}"`;
                }
                
                // Handle DATE and TIMESTAMP types - for key columns compare with time, for regular columns compare date only
                if (col1Type === 'DATE' || col2Type === 'DATE' || col1Type === 'TIMESTAMP' || col2Type === 'TIMESTAMP') {
                    if (isKeyColumn) {
                        // For key columns, compare full datetime (including time)
                        return `t1."${col1Name}" = t2."${col2Name}"`;
                    } else {
                        // For regular columns, compare only date part (ignoring time)
                        return `strftime(t1."${col1Name}", '%Y-%m-%d') = strftime(t2."${col2Name}", '%Y-%m-%d')`;
                    }
                }
                
                if (col1Type === 'VARCHAR' && col2Type === 'VARCHAR') {
                    return `UPPER(TRIM(t1."${col1Name}")) = UPPER(TRIM(t2."${col2Name}"))`;
                }
                
                return `t1."${col1Name}" = t2."${col2Name}"`;
            };
            
            const comparisonColumns = [];
            headers1.forEach((header, index) => {
                const shouldExclude = excludeColumns.some(excCol => {
                    if (typeof excCol === 'string') {
                        return header.toLowerCase().includes(excCol.toLowerCase());
                    } else if (typeof excCol === 'number') {
                        return index === excCol;
                    }
                    return false;
                });
                
                if (!shouldExclude) {
                    comparisonColumns.push(index);
                }
            });
            

            if (comparisonColumns.length === 0) {
                throw new Error('All columns are excluded from comparison');
            }
            
            // Use custom key columns if provided, otherwise detect automatically
            let allKeyColumns;
            if (customKeyColumns && customKeyColumns.length > 0) {
                allKeyColumns = customKeyColumns;
            } else {
                allKeyColumns = this.detectKeyColumnsSQL(headers1);
            }
            const keyColumns = allKeyColumns.filter(keyCol => comparisonColumns.includes(keyCol));
            
            // Ensure we have at least one key column
            if (keyColumns.length === 0) {
                keyColumns.push(0);
            }
            
            updateStageProgress('Setting up key columns', 55);
            
            // If no custom key columns were provided, mark the auto-detected ones in UI
            if (!customKeyColumns || customKeyColumns.length === 0) {
                const autoDetectedColumnNames = keyColumns.map(idx => headers1[idx] || `Column ${idx}`);

                // Set the auto-detected columns as selected in the UI
                setTimeout(() => {
                    if (typeof window.setSelectedKeyColumns === 'function') {
                        window.setSelectedKeyColumns(autoDetectedColumnNames);
                    } else {
                        console.error('❌ window.setSelectedKeyColumns is not available:', typeof window.setSelectedKeyColumns);
                    }
                }, 100);
            } else {
                console.log('🚫 Skipping auto-detection UI update because custom key columns are provided');
            }
            
            const doubleColumns = comparisonColumns.filter(idx => 
                finalColumnTypes1[idx] === 'DOUBLE' || finalColumnTypes2[idx] === 'DOUBLE'
            );
            if (doubleColumns.length > 0) { }

            // Get table counts - either from existing tables or newly created tables
            let table1Count, table2Count;
            if (tablesExist) {
                // Use counts from table existence check
                const table1CountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM table1');
                const table2CountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM table2');
                table1Count = Number(table1CountResult.toArray()[0]?.count || 0);
                table2Count = Number(table2CountResult.toArray()[0]?.count || 0);
            } else {
                // For newly created tables, count from data
                table1Count = data1.length - 1;
                table2Count = data2.length - 1;
            }
            
            updateStageProgress('Finding identical records', 65);
            
            const identicalSQL = `
                CREATE OR REPLACE TABLE identical_pairs AS
                SELECT 
                    t1.rowid as row1_id,
                    t2.rowid as row2_id,
                    'IDENTICAL' as match_type
                FROM table1 t1
                INNER JOIN table2 t2 ON (
                    ${comparisonColumns.map(colIdx => createComparisonCondition(colIdx, useTolerance, keyColumns.includes(colIdx))).join(' AND ')}
                )
            `;
            
            await window.duckdbLoader.query(identicalSQL);
            
            const identicalCountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM identical_pairs');
            const identicalCount = Number(identicalCountResult.toArray()[0]?.count || 0);

            updateStageProgress('Finding similar records', 75);

            const keyColumnChecks = keyColumns.map(colIdx => createComparisonCondition(colIdx, useTolerance, true)).join(' AND ');

            const minKeyMatchesRequired = Math.max(1, Math.ceil(keyColumns.length * (useTolerance ? 0.8 : 0.8))); // Minimum 80% of key fields
            const minTotalMatchesRequired = Math.max(2, Math.ceil(comparisonColumns.length * (useTolerance ? 0.6 : 0.7))); // Minimum 60-70% of columns for comparison

            const similarLimit = 9999999999;

            // Log debugging information about columns
            logDatabaseOperation('Preparing similar records SQL', {
                keyColumns: keyColumns,
                comparisonColumns: comparisonColumns,
                sanitizedHeaders1Length: sanitizedHeaders1.length,
                sanitizedHeaders2Length: sanitizedHeaders2.length,
                maxKeyColumnIndex: Math.max(...keyColumns),
                maxComparisonColumnIndex: Math.max(...comparisonColumns)
            });

            // Filter keyColumns and comparisonColumns to avoid out-of-bounds access
            const safeKeyColumns = keyColumns.filter(colIdx => 
                colIdx >= 0 && colIdx < sanitizedHeaders1.length && colIdx < sanitizedHeaders2.length
            );
            const safeComparisonColumns = comparisonColumns.filter(colIdx => 
                colIdx >= 0 && colIdx < sanitizedHeaders1.length && colIdx < sanitizedHeaders2.length
            );

            if (safeKeyColumns.length === 0) {
                logDatabaseOperation('Warning: No safe key columns found, using first column as fallback');
                safeKeyColumns.push(0);
            }

            if (safeComparisonColumns.length === 0) {
                throw new Error('No safe comparison columns found after filtering');
            }

            const similarSQL = `
                -- 1. Создаем индексы на ключевых колонках
                ${safeKeyColumns.map(colIdx => `
                CREATE INDEX IF NOT EXISTS idx_t1_${colIdx} ON table1("${sanitizedHeaders1[colIdx]}");
                CREATE INDEX IF NOT EXISTS idx_t2_${colIdx} ON table2("${sanitizedHeaders2[colIdx]}");
                `).join('')}

                -- 2. Создаем хеш для группировки похожих записей
                CREATE OR REPLACE TABLE table1_hashed AS
                SELECT *, 
                    ${safeKeyColumns.length > 0 
                        ? `hash(${safeKeyColumns.map(idx => `COALESCE(UPPER(CAST("${sanitizedHeaders1[idx]}" AS VARCHAR)), '')`).join(' || ')}) as key_hash`
                        : `hash(${safeComparisonColumns.slice(0, 3).map(idx => `COALESCE(UPPER(CAST("${sanitizedHeaders1[idx]}" AS VARCHAR)), '')`).join(' || ')}) as key_hash`
                    }
                FROM table1;

                CREATE OR REPLACE TABLE table2_hashed AS
                SELECT *, 
                    ${safeKeyColumns.length > 0 
                        ? `hash(${safeKeyColumns.map(idx => `COALESCE(UPPER(CAST("${sanitizedHeaders2[idx]}" AS VARCHAR)), '')`).join(' || ')}) as key_hash`
                        : `hash(${safeComparisonColumns.slice(0, 3).map(idx => `COALESCE(UPPER(CAST("${sanitizedHeaders2[idx]}" AS VARCHAR)), '')`).join(' || ')}) as key_hash`
                    }
                FROM table2;

                -- 3. Сопоставляем только записи с одинаковым хешем
                CREATE OR REPLACE TABLE similar_pairs AS
                WITH candidates AS (
                    SELECT 
                        t1.rowid as row1_id,
                        t2.rowid as row2_id,
                        ${safeComparisonColumns.map(colIdx => 
                            `CASE WHEN ${createComparisonCondition(colIdx, useTolerance, safeKeyColumns.includes(colIdx))} THEN 1 ELSE 0 END`
                        ).join(' + ')} as total_matches,
                        ${safeKeyColumns.length > 0 ? safeKeyColumns.map(colIdx => 
                            `CASE WHEN ${createComparisonCondition(colIdx, useTolerance, true)} THEN 1 ELSE 0 END`
                        ).join(' + ') : '0'} as key_matches
                    FROM table1_hashed t1
                    INNER JOIN table2_hashed t2 ON t1.key_hash = t2.key_hash
                    WHERE NOT EXISTS (
                        SELECT 1 FROM identical_pairs ip 
                        WHERE ip.row1_id = t1.rowid AND ip.row2_id = t2.rowid
                    )
                ),
                ranked_matches AS (
                    SELECT 
                        row1_id, row2_id, total_matches, key_matches,
                        ROW_NUMBER() OVER (PARTITION BY row1_id ORDER BY total_matches DESC, row2_id) as rn
                    FROM candidates
                    WHERE 
                        ${safeKeyColumns.length > 0 
                            ? `key_matches = ${safeKeyColumns.length}` // 100% совпадение ключевых колонок
                            : `total_matches >= ${Math.ceil(safeComparisonColumns.length * 0.5)}` // 50% общих колонок
                        }
                )
                SELECT 
                    row1_id, row2_id, 'SIMILAR' as match_type, total_matches, key_matches
                FROM ranked_matches
                WHERE rn = 1 -- Берем только лучшее совпадение для каждой строки из table1
                    AND total_matches < ${comparisonColumns.length}
                ORDER BY total_matches DESC
                LIMIT ${similarLimit}
            `;
            
            await window.duckdbLoader.query(similarSQL);
            
            const similarCountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM similar_pairs');
            const similarCount = Number(similarCountResult.toArray()[0]?.count || 0);
            
            // Для больших файлов пропускаем подсчет кандидатов (избегаем CROSS JOIN)
            const isLargeDataset = table1Count > 5000 || table2Count > 5000;
            let candidatesCount = 0;
            
            if (!isLargeDataset) {
                try {
                    const candidatesCountResult = await window.duckdbLoader.query(`
                        SELECT COUNT(*) as count FROM (
                            SELECT 
                                t1.rowid as row1_id,
                                t2.rowid as row2_id,
                                ${safeComparisonColumns.map(colIdx => 
                                    `CASE WHEN ${createComparisonCondition(colIdx, useTolerance, safeKeyColumns.includes(colIdx))} THEN 1 ELSE 0 END`
                                ).join(' + ')} as total_matches,
                                ${safeKeyColumns.length > 0 ? safeKeyColumns.map(colIdx => 
                                    `CASE WHEN ${createComparisonCondition(colIdx, useTolerance, true)} THEN 1 ELSE 0 END`
                                ).join(' + ') : '0'} as key_matches
                            FROM table1 t1
                            CROSS JOIN table2 t2
                            WHERE NOT EXISTS (
                                SELECT 1 FROM identical_pairs ip 
                                WHERE ip.row1_id = t1.rowid AND ip.row2_id = t2.rowid
                            )
                            LIMIT 50000  -- Ограничиваем для безопасности
                        ) candidates
                    `);
                    candidatesCount = Number(candidatesCountResult.toArray()[0]?.count || 0);
                } catch (error) {
                    console.warn('⚠️ Candidates count query failed for large dataset, skipping:', error.message);
                    candidatesCount = 0;
                }
            } else {
                console.log('📊 Skipping candidates count for large dataset to avoid CROSS JOIN performance issues');
            }
            
            // Пропускаем детальную статистику для больших файлов (избегаем CROSS JOIN)
            let filterStats = null;
            
            if (!isLargeDataset) {
                try {
                    const filterStatsResult = await window.duckdbLoader.query(`
                        SELECT 
                            COUNT(*) as total_candidates,
                            COUNT(CASE WHEN key_matches = ${safeKeyColumns.length} THEN 1 END) as passed_key_filter,
                            COUNT(CASE WHEN total_matches < ${safeComparisonColumns.length} THEN 1 END) as passed_not_identical_filter,
                            COUNT(CASE WHEN key_matches = ${safeKeyColumns.length} 
                                       AND total_matches < ${safeComparisonColumns.length} THEN 1 END) as passed_all_filters,
                            AVG(total_matches) as avg_total_matches,
                            AVG(key_matches) as avg_key_matches,
                            MIN(total_matches) as min_total_matches,
                            MAX(total_matches) as max_total_matches
                        FROM (
                            SELECT 
                                ${safeComparisonColumns.map(colIdx => 
                                    `CASE WHEN ${createComparisonCondition(colIdx, useTolerance, safeKeyColumns.includes(colIdx))} THEN 1 ELSE 0 END`
                                ).join(' + ')} as total_matches,
                                ${safeKeyColumns.length > 0 ? safeKeyColumns.map(colIdx => 
                                    `CASE WHEN ${createComparisonCondition(colIdx, useTolerance, true)} THEN 1 ELSE 0 END`
                                ).join(' + ') : '0'} as key_matches
                            FROM table1 t1
                            CROSS JOIN table2 t2
                            WHERE NOT EXISTS (
                                SELECT 1 FROM identical_pairs ip 
                                WHERE ip.row1_id = t1.rowid AND ip.row2_id = t2.rowid
                            )
                            LIMIT 100000  -- Ограничиваем для безопасности
                        ) stats
                    `);
                    filterStats = filterStatsResult.toArray()[0];
                } catch (error) {
                    console.warn('⚠️ Filter stats query failed for large dataset, skipping:', error.message);
                    filterStats = null;
                }
            } else {
                console.log('📊 Skipping detailed filter statistics for large dataset to avoid performance issues');
            }
            
            updateStageProgress('Generating final results', 90);
            
            const finalResultsSQL = `
                SELECT 'IDENTICAL' as type, row1_id, row2_id, ${headers1.length} as matches
                FROM identical_pairs
                
                UNION ALL
                
                SELECT 'SIMILAR' as type, row1_id, row2_id, total_matches as matches
                FROM similar_pairs
                
                UNION ALL
                
                SELECT 'ONLY_IN_TABLE1' as type, t1.rowid as row1_id, NULL as row2_id, 0 as matches
                FROM table1 t1
                WHERE NOT EXISTS (SELECT 1 FROM identical_pairs ip WHERE ip.row1_id = t1.rowid)
                  AND NOT EXISTS (SELECT 1 FROM similar_pairs sp WHERE sp.row1_id = t1.rowid)
                
                UNION ALL
                
                SELECT 'ONLY_IN_TABLE2' as type, NULL as row1_id, t2.rowid as row2_id, 0 as matches
                FROM table2 t2
                WHERE NOT EXISTS (SELECT 1 FROM identical_pairs ip WHERE ip.row2_id = t2.rowid)
                  AND NOT EXISTS (SELECT 1 FROM similar_pairs sp WHERE sp.row2_id = t2.rowid)
                
                ORDER BY type, row1_id, row2_id
            `;

            const result = await window.duckdbLoader.query(finalResultsSQL);
            const allResults = result.toArray();

            const identical = allResults.filter(r => r.type === 'IDENTICAL').map(r => ({
                row1: r.row1_id,
                row2: r.row2_id,
                status: 'IDENTICAL',
                matches: r.matches
            }));

            const similar = allResults.filter(r => r.type === 'SIMILAR').map(r => ({
                row1: r.row1_id,
                row2: r.row2_id,
                status: 'SIMILAR',
                matches: r.matches
            }));

            const onlyInTable1 = allResults.filter(r => r.type === 'ONLY_IN_TABLE1').map(r => ({
                row1: r.row1_id,
                row2: null,
                status: 'ONLY_IN_TABLE1'
            }));

            const onlyInTable2 = allResults.filter(r => r.type === 'ONLY_IN_TABLE2').map(r => ({
                row1: null,
                row2: r.row2_id,
                status: 'ONLY_IN_TABLE2'
            }));

            const duration = performance.now() - startTime;

            updateProgressMessage('Comparison completed successfully!', 100);

            const comparisonResults = {
                identical: identical,
                similar: similar, 
                onlyInTable1: onlyInTable1,
                onlyInTable2: onlyInTable2,
                table1Count: data1.length - 1,
                table2Count: data2.length - 1,
                commonColumns: headers1, 
                comparisonColumns: comparisonColumns,
                excludedColumns: excludeColumns, 
                keyColumns: keyColumns,
                performance: {
                    duration: duration,
                    rowsPerSecond: Math.round(((data1.length + data2.length) / duration) * 1000)
                }
            };

            // Update aggregation table with comparison results
            if (this.mode === 'wasm') {
                try {
                    await this.updateComparisonResults('table1', 'table2', comparisonResults);
                } catch (error) {
                    console.error('❌ Failed to update aggregation table:', error);
                }
            }

            logDatabaseOperation('Original Logic Comparison Completed', {
                duration: Math.round(duration),
                identicalRows: identical.length,
                similarRows: similar.length,
                onlyInTable1: onlyInTable1.length,
                onlyInTable2: onlyInTable2.length,
                rowsPerSecond: Math.round(((data1.length + data2.length) / duration) * 1000),
                zeroValuesFormatted: zeroValueFormattingCount
            });

            return comparisonResults;

        } catch (error) {
            clearAllProgressIntervals();
            updateProgressMessage('Comparison failed - switching to fallback mode', 0);
            console.error('❌ Multi-stage DuckDB comparison failed:', error);
            logDatabaseOperation('Original Logic Comparison Failed', {
                error: error.message || error
            });
            throw error;
        }
    }


    detectKeyColumnsSQL(headers, data = null) {
        
        if (!headers || headers.length === 0) {
            return [0]; 
        }
        
        const columnCount = headers.length;
        
        const keyIndicators = {
            high: ['id', 'uid', 'key', 'primary', 'identifier', 'код', 'номер', 'артикул', 'pk', 'primarykey'],
            medium: ['name', 'title', 'label', 'имя', 'название', 'наименование', 'фио', 'customer', 'client', 'клиент'],
            low: ['date', 'time', 'created', 'modified', 'дата', 'время', 'создан', 'изменен'],
            aggregation: ['yyyymm', 'yyyymmdd', 'year_month', 'yearmonth', 'period', 'период', 'reporting_period', 'отчетный_период',
                         'year', 'год', 'month', 'месяц', 'quarter', 'квартал', 'partition', 'раздел']
        };
        
        const columnScores = [];
        
        for (let colIndex = 0; colIndex < columnCount; colIndex++) {
            const header = (headers[colIndex] || '').toString().toLowerCase();
            let score = 0;
            
            let headerScore = 0;
            
            if (keyIndicators.aggregation.some(keyword => header.includes(keyword))) {
                if (header.match(/^y{4}m{2}$|yyyymm|year.*month|месяц.*год|period|период/i)) {
                    headerScore = 12; 
                } else {
                    headerScore = 9; 
                }
            }
            else if (keyIndicators.high.some(keyword => header.includes(keyword))) {
                headerScore = 10;
            }
            else if (keyIndicators.medium.some(keyword => header.includes(keyword))) {
                headerScore = 6;
            }
            else if (keyIndicators.low.some(keyword => header.includes(keyword))) {
                headerScore = 3;
            }
            else {
                headerScore = 1;
            }
            
            score += headerScore * 0.4;
            
            let uniquenessScore = 5;
            score += uniquenessScore * 0.4;
            
            const positionScore = Math.max(1, 10 - colIndex * 2); 
            score += positionScore * 0.2;
            
            columnScores.push({
                index: colIndex,
                header: headers[colIndex],
                score: score,
                headerScore: headerScore,
                uniquenessScore: uniquenessScore,
                positionScore: positionScore,
                isAggregationField: keyIndicators.aggregation.some(keyword => header.includes(keyword))
            });
        }
        
        columnScores.sort((a, b) => b.score - a.score);
        
        let keyColumns = [];
        
        const aggregationFields = columnScores.filter(col => col.isAggregationField);
        if (aggregationFields.length > 0) {
            keyColumns.push(aggregationFields[0].index);
            
            for (let i = 1; i < Math.min(3, aggregationFields.length); i++) {
                const col = aggregationFields[i];
                if (col.score >= 8) {
                    keyColumns.push(col.index);
                }
            }
        }
        
        if (keyColumns.length === 0) {
            if (columnScores.length > 0) {
                keyColumns.push(columnScores[0].index);
            }
        }
        
        for (let i = 0; i < Math.min(3, columnScores.length) && keyColumns.length < 3; i++) {
            const col = columnScores[i];
            if (!keyColumns.includes(col.index)) {
                if (col.score >= 6) {
                    keyColumns.push(col.index);
                }
            }
        }
        
        if (keyColumns.length === 0) {
            keyColumns = [0];
        }
        
        keyColumns.sort((a, b) => a - b);
        
        return keyColumns;
    }

    async createDuckDBTable(tableName, tableData) {
        logDatabaseOperation('DuckDB Table Creation Started', {
            tableName: tableName,
            rowCount: tableData.rows.length,
            columnCount: tableData.columns.length
        });

        const detectColumnType = (data, columnIndex) => {
            const sampleSize = Math.min(100, data.length);
            let numericCount = 0;
            let integerCount = 0;
            let dateCount = 0;
            let timestampCount = 0;
            let totalNonEmpty = 0;
            let hasDecimals = false;

            for (let i = 0; i < sampleSize; i++) {
                if (i >= data.length) break;
                
                const value = data[i]?.data?.[columnIndex];
                if (value && value.toString().trim() !== '') {
                    totalNonEmpty++;
                    const strValue = value.toString().trim();
                    
                    if (!isNaN(strValue) && !isNaN(parseFloat(strValue)) && isFinite(strValue)) {
                        numericCount++;
                        const numValue = parseFloat(strValue);
                        
                        if (Number.isInteger(numValue)) {
                            integerCount++;
                        } else {
                            hasDecimals = true; 
                        }
                    }
                    
                    const dateValue = new Date(strValue);
                    if (!isNaN(dateValue.getTime()) && strValue.match(/\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}|\d{2}\.\d{2}\.\d{4}/)) {
                        // Check if the string contains time information
                        if (strValue.match(/\d{2}:\d{2}(:\d{2})?/) || strValue.includes('T')) {
                            timestampCount++;
                        } else {
                            dateCount++;
                        }
                    }
                }
            }

            if (totalNonEmpty === 0) return 'VARCHAR';
            
            const numericRatio = numericCount / totalNonEmpty;
            const dateRatio = dateCount / totalNonEmpty;
            const timestampRatio = timestampCount / totalNonEmpty;

            if (timestampRatio >= 0.8) return 'TIMESTAMP';
            if (dateRatio >= 0.8) return 'DATE';
            if (numericRatio >= 0.9) {
                // If there's at least one decimal number, the entire column should be DOUBLE
                return hasDecimals ? 'DOUBLE' : 'BIGINT';
            }
            return 'VARCHAR';
        };

        // Function for formatting value depending on data type
        const formatValue = (value, columnType) => {
            if (value === null || value === undefined || value === '') {
                return 'NULL';
            }
            
            const strValue = value.toString().trim();
            
            switch (columnType) {
                case 'BIGINT':
                case 'INTEGER':
                    const intValue = parseInt(strValue);
                    return isNaN(intValue) ? 'NULL' : intValue.toString();
                
                case 'DOUBLE':
                case 'FLOAT':
                    const floatValue = parseFloat(strValue);
                    if (isNaN(floatValue)) {
                        return 'NULL';
                    } else {
                        // Round DOUBLE values to 2 decimal places
                        return columnType === 'DOUBLE' ? 
                            Math.round(floatValue * 100) / 100 : 
                            floatValue.toString();
                    }
                
                case 'DATE':
                    const dateValue = new Date(strValue);
                    if (isNaN(dateValue.getTime())) {
                        return `'${strValue.replace(/'/g, "''")}'`;
                    }
                    return `'${dateValue.toISOString().split('T')[0]}'`;
                
                case 'TIMESTAMP':
                    const timestampValue = new Date(strValue);
                    if (isNaN(timestampValue.getTime())) {
                        return `'${strValue.replace(/'/g, "''")}'`;
                    }
                    return `'${timestampValue.toISOString()}'`;
                
                case 'VARCHAR':
                default:
                    return `'${strValue.replace(/'/g, "''")}'`;
            }
        };

        const columnTypes = tableData.columns.map((_, i) => detectColumnType(tableData.rows, i));
        
        logDatabaseOperation('Creating DuckDB table schema', {
            tableName,
            columns: tableData.columns,
            columnTypes: columnTypes
        });
        
        const columns = tableData.columns.map((col, i) => `"${col}" ${columnTypes[i]}`).join(', ');
        const createTableSQL = `CREATE OR REPLACE TABLE ${tableName} (rowid INTEGER, ${columns})`;
        
        logDatabaseOperation('Executing CREATE TABLE', {
            query: createTableSQL.substring(0, 200) + '...'
        });
        
        await window.duckdbLoader.query(createTableSQL);

        // Verify table was created correctly
        try {
            const verifySQL = `DESCRIBE ${tableName}`;
            const description = await window.duckdbLoader.query(verifySQL);
            logDatabaseOperation('Table structure verified', {
                tableName,
                structure: description.toArray().map(row => ({ column_name: row.column_name, column_type: row.column_type }))
            });
        } catch (verifyError) {
            logDatabaseOperation('Table verification failed', {
                tableName,
                error: verifyError.message
            });
        }

        // Загружаем все данные одним запросом
        const allValues = tableData.rows.map(row => {
            const rowData = row.data.map((cell, colIdx) => formatValue(cell, columnTypes[colIdx])).join(', ');
            return `(${row.index + 1}, ${rowData})`;
        }).join(', ');
        
        if (allValues) {
            const insertSQL = `INSERT INTO ${tableName} VALUES ${allValues}`;
            await window.duckdbLoader.query(insertSQL);
        }

        logDatabaseOperation('DuckDB Table Creation Completed', {
            tableName: tableName,
            rowCount: tableData.rows.length,
            columnCount: tableData.columns.length
        });
    }

    async compareTablesLocal(table1Name, table2Name, excludeColumns = [], useTolerance = false, customKeyColumns = null) {
        if (!this.initialized) {
            throw new Error('Comparator not initialized');
        }

        logDatabaseOperation('Local Table Comparison Started', {
            table1Name: table1Name,
            table2Name: table2Name,
            excludeColumns: excludeColumns,
            useTolerance: useTolerance,
            customKeyColumns: customKeyColumns
        });

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

        logDatabaseOperation('Local Table Comparison Completed', {
            table1Name: table1Name,
            table2Name: table2Name,
            duration: Math.round(duration),
            identicalRows: identical.length,
            onlyInTable1: onlyInTable1.length,
            onlyInTable2: onlyInTable2.length,
            rowsPerSecond: Math.round((table1.rows.length + table2.rows.length) / (duration / 1000))
        });

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
        logDatabaseOperation('Fast Comparator Already Initialized', {
            mode: fastComparator.mode,
            result: 'Using existing instance'
        });
        return fastComparator;
    }

    logDatabaseOperation('Fast Comparator Initialization Started');

    try {
        fastComparator = new FastTableComparator();
        const initialized = await fastComparator.initialize();
        
        if (initialized) {
            logDatabaseOperation('Fast Comparator Initialization Completed', {
                mode: fastComparator.mode,
                result: 'Successfully initialized'
            });
            showFastModeStatus(true, fastComparator.mode);
            window.duckDBManager = fastComparator;
            window.duckDBAvailable = true;
            
            return fastComparator;
        } else {
            logDatabaseOperation('Fast Comparator Initialization Failed', {
                error: 'Initialization returned false'
            });
            showFastModeStatus(false);
            return null;
        }
    } catch (error) {
        logDatabaseOperation('Fast Comparator Initialization Failed', {
            error: error.message || error
        });
        showFastModeStatus(false);
        return null;
    }
}

function showFastModeStatus(available, mode = 'local') {
    const statusElement = document.getElementById('duckdb-status');
    if (statusElement) {
        if (available) {
            const modeText = mode === 'wasm' ? 'DuckDB WASM' : 'Optimized Fast';
            statusElement.innerHTML = `⚡ ${modeText} mode enabled - Enhanced performance!`;
            statusElement.className = 'duckdb-status duckdb-available show';
            
            const fastIndicators = document.querySelectorAll('.fast-mode-indicator');
            fastIndicators.forEach(indicator => {
                indicator.style.display = 'inline-block';
                indicator.textContent = mode === 'wasm' ? 'ULTRA' : 'FAST';
            });
            
            // Update export button if available
            const exportBtn = document.getElementById('exportExcelBtn');
            if (exportBtn) {
                exportBtn.title = `⚡ Fast export enabled - powered by ${modeText} engine`;
            }
            
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
            
            // Update export button
            const exportBtn = document.getElementById('exportExcelBtn');
            if (exportBtn) {
                exportBtn.title = 'Export to Excel - standard mode';
            }
            
            setTimeout(() => {
                statusElement.style.display = 'none';
            }, 3000);
        }
    }
}

// Track ongoing table creation to prevent duplicate calls
const ongoingTableCreation = new Set();

// Create table from preview data - called when file is loaded or sheet is changed
async function createTableFromPreviewData(tableNumber, data, fileName = null) {
    const callId = Math.random().toString(36).substr(2, 9);
    const tableName = `table${tableNumber}`;
    
    // Check if table creation is already in progress
    if (ongoingTableCreation.has(tableName)) {
        console.log(`⏭️ [CALL-${callId}] Table ${tableName} creation already in progress, skipping`);
        return false;
    }
    
    // Mark table creation as in progress
    ongoingTableCreation.add(tableName);
    
    console.log(`🚀 [CALL-${callId}] createTableFromPreviewData started for table${tableNumber}`);
    
    try {
        if (!fastComparator || !fastComparator.initialized) {
            console.log('⚠️ Fast comparator not initialized, skipping table creation');
            return false;
        }
        
        if (!data || data.length === 0) {
            console.log('⚠️ No data provided for table creation');
            return false;
        }
        
        // For WASM mode, create DuckDB tables directly for preview
        // For local mode, create local tables for preview
        const headers = data[0] || [];
        const dataRows = data.slice(1);
        
        if (fastComparator.mode === 'wasm') {
            // Get file name with sheet from global variables or parameter
            let displayName = fileName;
            if (!displayName) {
                // Try to get from global variables (assuming they are available)
                if (typeof window !== 'undefined') {
                    if (tableNumber === 1 && window.fileName1) {
                        // Use getFileDisplayName to include sheet name
                        if (typeof window.getFileDisplayName === 'function') {
                            displayName = window.getFileDisplayName(window.fileName1, window.sheetName1 || '');
                        } else {
                            displayName = window.fileName1;
                            if (window.sheetName1) {
                                displayName += ':' + window.sheetName1;
                            }
                        }
                    } else if (tableNumber === 2 && window.fileName2) {
                        // Use getFileDisplayName to include sheet name
                        if (typeof window.getFileDisplayName === 'function') {
                            displayName = window.getFileDisplayName(window.fileName2, window.sheetName2 || '');
                        } else {
                            displayName = window.fileName2;
                            if (window.sheetName2) {
                                displayName += ':' + window.sheetName2;
                            }
                        }
                    }
                }
                displayName = displayName || `Table ${tableNumber}`;
            }
            
            // Add to aggregation table first
            await fastComparator.addFileToAggregation(tableName, displayName, data, headers);
            
            // Create DuckDB table directly for preview in WASM mode
            
            // Detect column types for DuckDB table
            const detectColumnType = (data, columnIndex) => {
                const sampleSize = Math.min(100, data.length);
                let numericCount = 0;
                let integerCount = 0;
                let dateCount = 0;
                let timestampCount = 0;
                let totalNonEmpty = 0;
                let hasDecimals = false;

                for (let i = 0; i < Math.min(sampleSize, data.length); i++) {
                    const value = data[i]?.[columnIndex];
                    if (value && value.toString().trim() !== '') {
                        totalNonEmpty++;
                        const strValue = value.toString().trim();
                        
                        if (!isNaN(strValue) && !isNaN(parseFloat(strValue)) && isFinite(strValue)) {
                            numericCount++;
                            const numValue = parseFloat(strValue);
                            
                            if (Number.isInteger(numValue)) {
                                integerCount++;
                            } else {
                                hasDecimals = true;
                            }
                        }
                        
                        const dateValue = new Date(strValue);
                        if (!isNaN(dateValue.getTime()) && strValue.match(/\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}|\d{2}\.\d{2}\.\d{4}/)) {
                            if (strValue.match(/\d{2}:\d{2}(:\d{2})?/) || strValue.includes('T')) {
                                timestampCount++;
                            } else {
                                dateCount++;
                            }
                        }
                    }
                }

                if (totalNonEmpty === 0) return 'VARCHAR';
                
                const numericRatio = numericCount / totalNonEmpty;
                const dateRatio = dateCount / totalNonEmpty;
                const timestampRatio = timestampCount / totalNonEmpty;

                if (timestampRatio >= 0.8) return 'TIMESTAMP';
                if (dateRatio >= 0.8) return 'DATE';
                if (numericRatio >= 0.9) {
                    return hasDecimals ? 'DOUBLE' : 'BIGINT';
                }
                return 'VARCHAR';
            };

            // Sanitize column names for DuckDB
            const sanitizeColumnName = (name, index) => {
                if (!name || typeof name !== 'string') {
                    return `col_${index}`;
                }
                let sanitized = name.replace(/[^a-zA-Z0-9а-яА-Я_]/g, '_')
                                   .replace(/\s+/g, '_')
                                   .replace(/_{2,}/g, '_')
                                   .replace(/^_|_$/g, '');
                
                if (!sanitized || /^\d/.test(sanitized)) {
                    sanitized = `col_${index}_${sanitized}`;
                }
                
                return sanitized || `col_${index}`;
            };

            const sanitizedHeaders = headers.map((h, i) => sanitizeColumnName(h, i));
            const columnTypes = headers.map((_, i) => detectColumnType(dataRows, i));

            // Format value function for data insertion
            const formatValue = (value, columnType) => {
                if (value === null || value === undefined || value === '') {
                    return 'NULL';
                }
                
                const strValue = value.toString().trim();
                
                switch (columnType) {
                    case 'BIGINT':
                    case 'INTEGER':
                        const intValue = parseInt(strValue);
                        return isNaN(intValue) ? 'NULL' : intValue.toString();
                    
                    case 'DOUBLE':
                    case 'FLOAT':
                        const floatValue = parseFloat(strValue);
                        if (isNaN(floatValue)) {
                            return 'NULL';
                        } else {
                            return columnType === 'DOUBLE' ? 
                                (Math.round(floatValue * 100) / 100).toString() : 
                                floatValue.toString();
                        }
                    
                    case 'DATE':
                        const dateValue = new Date(strValue);
                        if (isNaN(dateValue.getTime())) {
                            return `'${strValue.replace(/'/g, "''")}'`;
                        }
                        return `'${dateValue.toISOString().split('T')[0]}'`;
                    
                    case 'TIMESTAMP':
                        const timestampValue = new Date(strValue);
                        if (isNaN(timestampValue.getTime())) {
                            return `'${strValue.replace(/'/g, "''")}'`;
                        }
                        return `'${timestampValue.toISOString()}'`;
                    
                    case 'VARCHAR':
                    default:
                        return `'${strValue.replace(/'/g, "''")}'`;
                }
            };

            // Always create or replace the table to ensure fresh data on tab switch
            console.log(`🔄 [CALL-${callId}] Creating/replacing table ${tableName}`);
            
            const createTableSQL = `CREATE OR REPLACE TABLE ${tableName} (
                rowid INTEGER,
                ${sanitizedHeaders.map((h, i) => `"${h}" ${columnTypes[i]}`).join(', ')}
            )`;

            await window.duckdbLoader.query(createTableSQL);

            logDatabaseOperation('Preview Table Creation Started', {
                tableName: tableName,
                rowCount: dataRows.length,
                columnCount: headers.length
            });

            // Загружаем все данные одним запросом
            const allValues = dataRows.map((row, idx) => {
                const rowId = idx;
                const cleanRow = headers.map((_, colIdx) => {
                    const val = row[colIdx];
                    return formatValue(val, columnTypes[colIdx]);
                }).join(', ');
                return `(${rowId}, ${cleanRow})`;
            }).join(', ');

            if (allValues) {
                const insertSQL = `INSERT INTO ${tableName} VALUES ${allValues}`;
                await window.duckdbLoader.query(insertSQL);
            }
            
            logDatabaseOperation('Preview Table Creation Completed', {
                tableName: tableName,
                rowCount: dataRows.length,
                columnCount: headers.length
            });
            
            console.log(`✅ [CALL-${callId}] Table ${tableNumber} created in DuckDB WASM mode with ${dataRows.length} rows`);
        } else {
            // For local mode, create local tables for preview
            await fastComparator.createTableFromData(tableName, dataRows, headers);
            console.log(`✅ [CALL-${callId}] Table ${tableNumber} created in local mode with ${dataRows.length} rows`);
        }
        
        // Ensure UI is updated after table creation
        if (fastComparator.mode === 'wasm') {
            await fastComparator.updateAggregationUI();
        }
        
        console.log(`🏁 [CALL-${callId}] createTableFromPreviewData completed for table${tableNumber}`);
        return true;
        
    } catch (error) {
        console.error(`❌ [CALL-${callId}] Failed to create table ${tableNumber}:`, error);
        return false;
    } finally {
        // Always clear the ongoing creation flag
        ongoingTableCreation.delete(tableName);
    }
}

// Create preview from existing table - called to update UI from table data
async function createPreviewFromTable(tableNumber, elementId) {
    if (!fastComparator || !fastComparator.initialized) {
        return;
    }
    
    try {
        const tableName = `table${tableNumber}`;
        
        if (fastComparator.mode === 'wasm') {
            // For WASM mode, query DuckDB table directly
            try {
                const result = await window.duckdbLoader.query(`
                    SELECT * FROM ${tableName} 
                    ORDER BY rowid 
                    LIMIT 10
                `);
                
                if (result && result.toArray) {
                    const rows = result.toArray();
                    if (rows.length > 0) {
                        // Get column names (excluding rowid)
                        const columns = Object.keys(rows[0]).filter(col => col !== 'rowid');
                        
                        // Reconstruct data format for renderPreview
                        const fullData = [columns]; // Headers
                        rows.forEach(row => {
                            const dataRow = columns.map(col => row[col]);
                            fullData.push(dataRow);
                        });
                        
                        // Use existing renderPreview function but don't call createTableFromPreviewData again
                        if (typeof renderPreview === 'function') {
                            const originalCreateTable = window.MaxPilotDuckDB.createTableFromPreviewData;
                            window.MaxPilotDuckDB.createTableFromPreviewData = () => {}; // Temporarily disable to avoid recursion
                            renderPreview(fullData, elementId);
                            window.MaxPilotDuckDB.createTableFromPreviewData = originalCreateTable; // Restore
                        }
                        
                        console.log(`✅ Preview updated for table ${tableNumber} from DuckDB with ${rows.length} rows`);
                        return;
                    }
                }
            } catch (error) {
                console.log(`Note: Could not query DuckDB table ${tableName}:`, error.message);
            }
        }
        
        // Fallback to local mode or if DuckDB query failed
        const tableData = fastComparator.tables.get(tableName);
        
        if (!tableData) {
            return;
        }
        
        // Reconstruct full data array (headers + rows)
        const fullData = [tableData.columns];
        tableData.rows.slice(0, 10).forEach(row => { // Show first 10 rows in preview
            fullData.push(row.data);
        });
        
        // Use existing renderPreview function but don't call createTableFromPreviewData again
        if (typeof renderPreview === 'function') {
            const originalCreateTable = window.MaxPilotDuckDB.createTableFromPreviewData;
            window.MaxPilotDuckDB.createTableFromPreviewData = () => {}; // Temporarily disable to avoid recursion
            renderPreview(fullData, elementId);
            window.MaxPilotDuckDB.createTableFromPreviewData = originalCreateTable; // Restore
        }
        
        console.log(`✅ Preview updated for table ${tableNumber} from stored data`);
        
    } catch (error) {
        console.error(`❌ Failed to create preview from table ${tableNumber}:`, error);
    }
}

async function compareTablesWithFastComparator(data1, data2, excludeColumns = [], useTolerance = false, tolerance = 1.5, customKeyColumns = null) {
    logDatabaseOperation('Table Comparison Started', {
        table1Rows: data1 ? data1.length : 0,
        table2Rows: data2 ? data2.length : 0,
        excludeColumns: excludeColumns,
        useTolerance: useTolerance,
        tolerance: tolerance,
        customKeyColumns: customKeyColumns
    });

    try {

        if (!fastComparator || !fastComparator.initialized) {
            console.log('❌ Fast comparator not initialized');
            logDatabaseOperation('Table Comparison Failed', {
                error: 'Fast comparator not initialized'
            });
            return null;
        }

        // Initialize progress for comparison
        const totalRows = (data1 ? data1.length - 1 : 0) + (data2 ? data2.length - 1 : 0);
        initializeProgress(totalRows, 'Preparing comparison data...');

        // Store original data for export functionality - extract data rows only (skip headers)
        if (data1 && data1.length > 1) {
            window.originalBody1 = data1.slice(1); // Store body without headers
        }
        if (data2 && data2.length > 1) {
            window.originalBody2 = data2.slice(1); // Store body without headers
        }
        
        // Store headers and column information for export
        if (data1 && data1.length > 0 && data2 && data2.length > 0) {
            const headers1 = data1[0] || [];
            const headers2 = data2[0] || [];
            
            // Create unified header list
            const allHeaders = [...new Set([...headers1, ...headers2])];
            window.currentFinalHeaders = allHeaders;
            window.currentFinalAllCols = allHeaders.length;
        }

        const startTime = performance.now();

        if (fastComparator.mode === 'wasm') {
            
            // Apply exclude columns filtering BEFORE preparing data for comparison
            let filteredData1 = data1;
            let filteredData2 = data2;
            
            if (excludeColumns && excludeColumns.length > 0) {
                filteredData1 = filterExcludeColumns(data1, excludeColumns);
                filteredData2 = filterExcludeColumns(data2, excludeColumns);
            }
            
            // Prepare data for comparison to get column information
            const { data1: alignedData1, data2: alignedData2, columnInfo } = prepareDataForComparison(filteredData1, filteredData2);
            

            const result = await fastComparator.compareTablesFast(alignedData1, alignedData2, [], useTolerance, customKeyColumns);
            
            if (!result) {
                console.log('⚠️ DuckDB WASM returned empty result');
                return null;
            }
            
            const duration = performance.now() - startTime;
            
            // Add column information to the result
            result.columnInfo = columnInfo;
            result.alignedData1 = alignedData1;
            result.alignedData2 = alignedData2;
            
            return result;
            
        } else {
            
            updateStageProgress('Filtering excluded columns', 10);
            
            // Apply exclude columns filtering BEFORE preparing data for comparison
            let filteredData1 = data1;
            let filteredData2 = data2;
            
            if (excludeColumns && excludeColumns.length > 0) {
                filteredData1 = filterExcludeColumns(data1, excludeColumns);
                filteredData2 = filterExcludeColumns(data2, excludeColumns);
            }
            
            updateStageProgress('Preparing data alignment', 20);
            
            // Prepare data for comparison to get column information
            const { data1: alignedData1, data2: alignedData2, columnInfo } = prepareDataForComparison(filteredData1, filteredData2);
            
            updateStageProgress('Creating comparison tables', 30);
            
            const headers1 = alignedData1[0] || [];
            const headers2 = alignedData2[0] || [];
            const dataRows1 = alignedData1.slice(1);
            const dataRows2 = alignedData2.slice(1);
            
            // For WASM mode, pass the data directly to compareTablesFast
            // For local mode, create tables first
            let comparisonResult;
            if (fastComparator.mode === 'wasm') {
                updateStageProgress('Running table comparison', 60);
                
                comparisonResult = await fastComparator.compareTablesFast(
                    alignedData1, alignedData2, excludeColumns, useTolerance, customKeyColumns
                );
            } else {
                await fastComparator.createTableFromData('table1', dataRows1, headers1);
                await fastComparator.createTableFromData('table2', dataRows2, headers2);
                
                updateStageProgress('Running table comparison', 60);
                
                comparisonResult = await fastComparator.compareTablesFast(
                    'table1', 'table2', excludeColumns, useTolerance, customKeyColumns
                );
            }

            updateStageProgress('Finalizing results', 90);

            const totalTime = performance.now() - startTime;

            // Add column information to the result
            comparisonResult.columnInfo = columnInfo;
            comparisonResult.alignedData1 = alignedData1;
            comparisonResult.alignedData2 = alignedData2;

            updateProgressMessage('Comparison completed successfully!', 100);

            logDatabaseOperation('Table Comparison Completed', {
                duration: Math.round(totalTime),
                identicalRows: comparisonResult.identical?.length || 0,
                onlyInTable1: comparisonResult.onlyInTable1?.length || 0,
                onlyInTable2: comparisonResult.onlyInTable2?.length || 0,
                result: 'Comparison successful'
            });

            return comparisonResult;
        }
        
    } catch (error) {
        clearAllProgressIntervals();
        updateProgressMessage('Comparison failed', 0);
        console.error('❌ Fast comparison failed:', error);
        logDatabaseOperation('Table Comparison Failed', {
            error: error.message || error
        });
        throw error;
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
        console.log('❌ No data loaded:', { data1Length: data1?.length, data2Length: data2?.length });
        document.getElementById('result').innerText = 'Please, load both files.';
        document.getElementById('summary').innerHTML = '';
        showPlaceholderMessage();
        return;
    }

    const totalRows = Math.max(data1.length, data2.length);

    try {

        if (fastComparator && fastComparator.initialized) {
            resultDiv.innerHTML = '<div class="comparison-loading-enhanced">⚡ Using fast comparison engine...</div>';
            summaryDiv.innerHTML = '<div style="text-align: center; padding: 10px;">Processing large dataset with enhanced performance...</div>';
            
            setTimeout(async () => {
                try {
                    const totalRows = Math.max(data1.length, data2.length);
                    const columnCount = Math.max(data1[0]?.length || 0, data2[0]?.length || 0);
                    
                    if (totalRows > 20000 || columnCount > 30) {
                        resultDiv.innerHTML = createProgressIndicator('large');
                        updateProgressMessage('Initializing comparison for large dataset...');
                    } else {
                        resultDiv.innerHTML = createProgressIndicator('normal');
                        updateProgressMessage('Processing comparison...');
                    }
                    
                    const excludedColumns = getExcludedColumns ? getExcludedColumns() : [];
                    const tolerance = window.currentTolerance || 1.5;
                    
                    // Get selected key columns from UI
                    const selectedKeyColumns = getSelectedKeyColumns ? getSelectedKeyColumns() : [];
                    let customKeyColumns = null;
                    
                    if (selectedKeyColumns && selectedKeyColumns.length > 0) {
                        // Convert selected column names to indexes
                        const headers = data1.length > 0 ? data1[0] : [];
                        customKeyColumns = selectedKeyColumns.map(columnName => {
                            const index = headers.indexOf(columnName);
                            return index !== -1 ? index : null;
                        }).filter(index => index !== null);
                        
                    } else {  }
                    
                    const fastResult = await compareTablesWithFastComparator(data1, data2, excludedColumns, useTolerance, tolerance, customKeyColumns);
                    
                    if (fastResult) {
                        await processFastComparisonResults(fastResult, useTolerance);
                        
                        if (!customKeyColumns || customKeyColumns.length === 0) {
 
                            // Get the key columns that were actually used from the fast result
                            // For now, we'll get them from current selection (if any) or try to extract from the comparison
                            setTimeout(() => {
                                if (typeof window.setSelectedKeyColumns === 'function') {
                                    // For now, let's try to mark the first column as it's commonly the key
                                    const headers = data1.length > 0 ? data1[0] : [];
                                    const firstColumnName = headers[0];
                                    if (firstColumnName) {
                                        window.setSelectedKeyColumns([firstColumnName]);
                                    }
                                } else {
                                    console.error('❌ setSelectedKeyColumns not available');
                                }
                            }, 500);
                        } else { }
                    } else {
                        await performComparison();
                    }
                } catch (error) {
                    clearAllProgressIntervals(); // Clear all intervals on error
                    console.error('❌ DuckDB WASM comparison failed:', error);
                    
                    // Fallback to standard mode
                    await performComparison();
                }
                
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
    // Show final processing stage
    updateProgressMessage('Processing results...', 95);
    
    const { identical, similar, onlyInTable1, onlyInTable2, table1Count, table2Count, commonColumns, performance, columnInfo } = fastResult;
    
    window.currentFastResult = fastResult;
    
    const perfData = performance || { duration: 0, rowsPerSecond: 0 };

    updateProgressMessage('Displaying results...', 100);
    
    setTimeout(() => {
        const resultDiv = document.getElementById('result');
        if (resultDiv) {
            resultDiv.innerHTML = '';
            resultDiv.style.display = 'none';
        }
    }, 500); // Give time to show 100% before hiding
    
    const identicalCount = identical?.length || 0;
    const similarCount = similar?.length || 0;
    const diffCount = similarCount; // similar contains different/modified rows
    const totalMatches = identicalCount;
    const similarity = table1Count > 0 ? ((identicalCount / Math.max(table1Count, table2Count)) * 100).toFixed(1) : 0;
    
    let percentClass = 'percent-high';
    if (parseFloat(similarity) < 30) percentClass = 'percent-low';
    else if (parseFloat(similarity) < 70) percentClass = 'percent-medium';
    else percentClass = 'percent-high';
    
    const file1Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName1 || 'File 1', window.sheetName1 || '')
        : (window.fileName1 || 'File 1');
    const file2Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName2 || 'File 2', window.sheetName2 || '')
        : (window.fileName2 || 'File 2');

    const tableHeaders = getSummaryTableHeaders();
    
    // Display column alignment info if available (using the same logic as functions.js)
    if (columnInfo && columnInfo.hasCommonColumns && columnInfo.reordered) {
        const infoDiv = document.createElement('div');
        infoDiv.className = 'column-alignment-info';
        infoDiv.style.cssText = 'background: #d1ecf1; border: 1px solid #bee5eb; color: #0c5460; padding: 12px; margin: 15px 0; border-radius: 6px; font-size: 14px;';
        infoDiv.innerHTML = `
            <strong>📊 Column Alignment:</strong> Columns have been automatically aligned by name for accurate comparison. 
            ${columnInfo.commonCount} common columns found.
            ${columnInfo.onlyInFile1.length > 0 ? `<br><strong>Only in File 1:</strong> ${columnInfo.onlyInFile1.join(', ')}` : ''}
            ${columnInfo.onlyInFile2.length > 0 ? `<br><strong>Only in File 2:</strong> ${columnInfo.onlyInFile2.join(', ')}` : ''}
        `;
        
        const summaryEl = document.getElementById('summary');
        if (summaryEl.firstChild) {
            summaryEl.insertBefore(infoDiv, summaryEl.firstChild);
        } else {
            summaryEl.appendChild(infoDiv);
        }
    } else if (columnInfo && !columnInfo.hasCommonColumns) {
        const warningDiv = document.createElement('div');
        warningDiv.className = 'column-warning-info';
        warningDiv.style.cssText = 'background: #fff3cd; border: 1px solid #ffeaa7; color: #856404; padding: 12px; margin: 15px 0; border-radius: 6px; font-size: 14px;';
        warningDiv.innerHTML = `
            <strong>⚠️ Warning:</strong> No common column names found between files. 
            Comparison will be done by column position. For more accurate results, ensure both files have matching column headers.
        `;
        
        const summaryEl = document.getElementById('summary');
        if (summaryEl.firstChild) {
            summaryEl.insertBefore(warningDiv, summaryEl.firstChild);
        } else {
            summaryEl.appendChild(warningDiv);
        }
    }
    
    // Use the same diff columns logic as functions.js
    let onlyInFile1, onlyInFile2;
    
    if (columnInfo && (columnInfo.hasCommonColumns || columnInfo.onlyInFile1 || columnInfo.onlyInFile2)) {
        onlyInFile1 = columnInfo.onlyInFile1 || [];
        onlyInFile2 = columnInfo.onlyInFile2 || [];
    } else {
        onlyInFile1 = [];
        onlyInFile2 = [];
    }
    
    let diffColumns1 = onlyInFile1.length > 0 ? onlyInFile1.join(', ') : '-';
    let diffColumns2 = onlyInFile2.length > 0 ? onlyInFile2.join(', ') : '-';
    
    // Store current diff columns globally for other functions to use
    window.currentDiffColumns1 = diffColumns1;
    window.currentDiffColumns2 = diffColumns2;
    
    // Create HTML for diff columns with red styling if there are differences
    let diffColumns1Html = onlyInFile1.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns1}</span>` : diffColumns1;
    let diffColumns2Html = onlyInFile2.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns2}</span>` : diffColumns2;
    
    const summaryHTML = `
        <div style="overflow-x: auto; margin: 20px 0;">
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>${tableHeaders.file}</th>
                        <th>${tableHeaders.rowCount}</th>
                        <th>${tableHeaders.rowsOnlyInFile}</th>
                        <th>${tableHeaders.diffRows}</th>
                        <th>Identical Rows</th>
                        <th>${tableHeaders.similarity}</th>
                        <th>${tableHeaders.diffColumns}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>${file1Name}</strong></td>
                        <td>${table1Count.toLocaleString()}</td>
                        <td>${onlyInTable1.length.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 16px; color: #dc3545;">${diffCount.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 16px; color: #28a745;">${identicalCount.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 18px;" class="percent-cell ${percentClass}">${similarity}%</td>
                        <td>${diffColumns1Html}</td>
                    </tr>
                    <tr>
                        <td><strong>${file2Name}</strong></td>
                        <td>${table2Count.toLocaleString()}</td>
                        <td>${onlyInTable2.length.toLocaleString()}</td>
                        <td>${diffColumns2Html}</td>
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
    const buttonsContainer = document.querySelector('.buttons-container');
    const exportButtonHalf = exportBtn ? exportBtn.closest('.button-half') : null;
    
    if (exportBtn && buttonsContainer) {
        exportBtn.style.display = 'inline-block';
        exportBtn.classList.remove('export-btn-hidden');
        
        if (exportButtonHalf) {
            exportButtonHalf.classList.remove('export-hidden');
        }
        buttonsContainer.classList.remove('export-hidden');
        
        } else {
        console.log('❌ Export button elements not found:', { exportBtn: !!exportBtn, buttonsContainer: !!buttonsContainer });
    }
}

async function generateDetailedComparisonTable(fastResult, useTolerance) {
    const { identical, similar, onlyInTable1, onlyInTable2, commonColumns, alignedData1, alignedData2 } = fastResult;
    
    // Clear any remaining loading messages
    const resultDiv = document.getElementById('result');
    if (resultDiv) {
        resultDiv.innerHTML = '';
        resultDiv.style.display = 'none';
    }
    
    const workingData1 = alignedData1 || data1;
    const workingData2 = alignedData2 || data2;
    
    const pairs = [];
    
    const maxIdenticalToShow = (onlyInTable1?.length || 0) === 0 && (onlyInTable2?.length || 0) === 0 ? 10000 : 1000;
    (identical || []).slice(0, maxIdenticalToShow).forEach(identicalPair => {
        const row1Index = identicalPair.row1;
        const row2Index = identicalPair.row2;
        
        const row1 = (row1Index >= 0 && row1Index + 1 < workingData1.length) ? workingData1[row1Index + 1] : null;
        const row2 = (row2Index >= 0 && row2Index + 1 < workingData2.length) ? workingData2[row2Index + 1] : null;
        
        if (!row1 || !row2 || row1 === workingData1[0] || row2 === workingData2[0]) {
            console.warn('⚠️ Skipping invalid IDENTICAL pair:', { row1Index, row2Index, row1, row2 });
            return;
        }
        
        pairs.push({
            row1: row1,
            row2: row2,
            index1: row1Index,
            index2: row2Index,
            isDifferent: false,
            onlyIn: null,
            matchType: 'IDENTICAL'
        });
    });
    
    const maxSimilarToShow = 100;
    (similar || []).slice(0, maxSimilarToShow).forEach(similarPair => {
        const row1Index = similarPair.row1;
        const row2Index = similarPair.row2;
        
        const row1 = (row1Index >= 0 && row1Index + 1 < workingData1.length) ? workingData1[row1Index + 1] : null;
        const row2 = (row2Index >= 0 && row2Index + 1 < workingData2.length) ? workingData2[row2Index + 1] : null;
        
        if (!row1 || !row2 || row1 === workingData1[0] || row2 === workingData2[0]) {
            console.warn('⚠️ Skipping invalid SIMILAR pair:', { row1Index, row2Index, row1, row2 });
            return;
        }
        
        pairs.push({
            row1: row1,
            row2: row2,
            index1: row1Index,
            index2: row2Index,
            isDifferent: true,
            onlyIn: null,
            matchType: 'SIMILAR',
            matches: similarPair.matches || 0
        });
    });
    
    (onlyInTable1 || []).forEach(diff => {
        const rowIndex = diff.row1;
        
        const row1 = (rowIndex >= 0 && rowIndex + 1 < workingData1.length) ? workingData1[rowIndex + 1] : null;
        
        if (!row1 || row1 === workingData1[0]) {
            console.warn('⚠️ Skipping invalid ONLY_IN_TABLE1:', { rowIndex, row1 });
            return;
        }
        
        pairs.push({
            row1: row1,
            row2: null,
            index1: rowIndex,
            index2: -1,
            isDifferent: true,
            onlyIn: 'table1',
            matchType: 'ONLY_IN_TABLE1'
        });
    });
    
    (onlyInTable2 || []).forEach(diff => {
        const rowIndex = diff.row2;
        
        const row2 = (rowIndex >= 0 && rowIndex + 1 < workingData2.length) ? workingData2[rowIndex + 1] : null;
        
        if (!row2 || row2 === workingData2[0]) {
            console.warn('⚠️ Skipping invalid ONLY_IN_TABLE2:', { rowIndex, row2 });
            return;
        }
        
        pairs.push({
            row1: null,
            row2: row2,
            index1: -1,
            index2: rowIndex,
            isDifferent: true,
            onlyIn: 'table2',
            matchType: 'ONLY_IN_TABLE2'
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
    
    // Add horizontal scroll synchronization
    syncTableScroll();
}

// Helper functions for improved value comparison
function isNumericStringLocal(str) {
    if (!str || typeof str !== 'string') return false;
    const cleanStr = str.replace(/['",$\s]/g, '').trim();
    return !isNaN(cleanStr) && !isNaN(parseFloat(cleanStr)) && isFinite(cleanStr);
}

function parseNumberLocal(str) {
    if (!str) return 0;
    const cleanStr = str.toString().replace(/['",$\s]/g, '').trim();
    return parseFloat(cleanStr) || 0;
}

function isWithinToleranceLocal(val1, val2, tolerance = 0.015) {
    const num1 = parseNumberLocal(val1);
    const num2 = parseNumberLocal(val2);
    
    if (num1 === 0 && num2 === 0) return true;
    if (num1 === 0 || num2 === 0) return false;
    
    const diff = Math.abs(num1 - num2);
    const avg = (Math.abs(num1) + Math.abs(num2)) / 2;
    
    return (diff / avg) <= tolerance;
}

function compareValuesImproved(v1, v2, useTolerance = false) {
    if (!v1 && !v2) return 'identical';
    if (!v1 || !v2) return 'different';
    
    const str1 = v1.toString().trim();
    const str2 = v2.toString().trim();
    
    // Basic string comparison
    if (str1.toUpperCase() === str2.toUpperCase()) {
        return 'identical';
    }
    
    // Numeric comparison with special handling for zeros
    if (isNumericStringLocal(str1) && isNumericStringLocal(str2)) {
        const num1 = parseNumberLocal(str1);
        const num2 = parseNumberLocal(str2);
        
        // Special case for zeros
        if (num1 === 0 && num2 === 0) {
            return 'identical';
        }
        
        // Check tolerance if enabled
        if (useTolerance && isWithinToleranceLocal(str1, str2, 0.015)) {
            return 'tolerance';
        }
        
        // Exact numeric comparison
        if (num1 === num2) {
            return 'identical';
        }
    }
    
    return 'different';
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
    
    // Get column info to identify diff columns (use global variables from the summary)
    const onlyInFile1 = [];
    const onlyInFile2 = [];
    
    // Parse diff columns from global variables if they exist
    if (window.currentDiffColumns1 && window.currentDiffColumns1 !== '-') {
        onlyInFile1.push(...window.currentDiffColumns1.split(', ').filter(col => col.trim() !== ''));
    }
    if (window.currentDiffColumns2 && window.currentDiffColumns2 !== '-') {
        onlyInFile2.push(...window.currentDiffColumns2.split(', ').filter(col => col.trim() !== ''));
    }
    
    let headerHtml = '<tr><th title="Source - shows which file the data comes from" class="source-column">Source</th>';
    realHeaders.forEach((header, index) => {
        const headerText = header || `Column ${index + 1}`;
        
        // Check if this column is a diff column
        let isDiffColumn = false;
        let diffColumnType = '';
        if (onlyInFile1.includes(headerText)) {
            isDiffColumn = true;
            diffColumnType = 'only-in-file1';
        } else if (onlyInFile2.includes(headerText)) {
            isDiffColumn = true;
            diffColumnType = 'only-in-file2';
        }
        
        // Add column type icon for test_wasm page
        let columnTypeIcon = '';
        const isTestWasmPage = window.location.pathname.includes('/test_wasm');
        
        if (headerText && typeof getColumnType === 'function' && isTestWasmPage) {
            // Get column data for type detection from pairs parameter
            let columnValues = [];
            
            // Find column index in headers
            const columnIndex = headers.indexOf(headerText);
            
            if (columnIndex !== -1 && pairs && pairs.length > 0) {
                
                // Extract values from pairs data
                pairs.forEach((pair, pairIndex) => {                    
                    // Each pair has row1 and row2 properties (not data1/data2)
                    if (pair.row1 && pair.row1[columnIndex] !== undefined) {
                        columnValues.push(pair.row1[columnIndex]);
                    }
                    if (pair.row2 && pair.row2[columnIndex] !== undefined) {
                        columnValues.push(pair.row2[columnIndex]);
                    }
                });
            }
            
            const columnType = getColumnType(columnValues, headerText);
            columnTypeIcon = `<span class="column-type-icon column-type-${columnType}" title="Column type: ${columnType}"></span>`;
        }
        
        // Add diff column indicator
        let diffColumnIcon = '';
        let headerClass = 'sortable';
        let headerTitle = headerText;
        
        if (isDiffColumn) {
            headerClass += ` diff-column ${diffColumnType}`;
            if (diffColumnType === 'only-in-file1') {
                diffColumnIcon = '<span class="diff-column-indicator file1-only" title="Column only in File 1">📍</span>';
                headerTitle += ' (Only in File 1)';
            } else if (diffColumnType === 'only-in-file2') {
                diffColumnIcon = '<span class="diff-column-indicator file2-only" title="Column only in File 2">📍</span>';
                headerTitle += ' (Only in File 2)';
            }
        }
        
        headerHtml += `<th class="${headerClass}" onclick="sortTable(${index})" title="${headerTitle}">${columnTypeIcon}${diffColumnIcon}${headerText}</th>`;
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
    
    // Get checkbox filter states
    const hideSameEl = document.getElementById('hideSameRows');
    const hideDiffEl = document.getElementById('hideDiffColumns');
    const hideNewRows1El = document.getElementById('hideNewRows1');
    const hideNewRows2El = document.getElementById('hideNewRows2');
    
    const hideSame = hideSameEl ? hideSameEl.checked : false;
    const hideDiffRows = hideDiffEl ? hideDiffEl.checked : false;
    const hideNewRows1 = hideNewRows1El ? hideNewRows1El.checked : false;
    const hideNewRows2 = hideNewRows2El ? hideNewRows2El.checked : false;
    
    // Use getFileDisplayName to display filename with sheet in Source column
    const file1Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName1 || 'File 1', window.sheetName1 || '')
        : (window.fileName1 || 'File 1');
    const file2Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName2 || 'File 2', window.sheetName2 || '')
        : (window.fileName2 || 'File 2');
    
    pairs.slice(0, 10000).forEach((pair, index) => {
        // Check checkbox filters
        const row1 = pair.row1;
        const row2 = pair.row2;
        
        // Determine row characteristics for filtering
        let allSame = true;
        let hasWarn = false;
        let isEmpty = true;
        
        if (row1 && row2) {
            // Compare rows to determine allSame and hasWarn
            for (let c = 0; c < realHeaders.length; c++) {
                const v1 = row1[c] !== undefined ? row1[c] : '';
                const v2 = row2[c] !== undefined ? row2[c] : '';
                
                if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                    isEmpty = false;
                }
                
                // Use improved comparison function
                const compResult = compareValuesImproved(v1, v2, false);
                if (compResult !== 'identical') {
                    allSame = false;
                    hasWarn = true;
                }
            }
        } else {
            allSame = false;
            hasWarn = false;
            
            // Check for emptiness
            const existingRow = row1 || row2;
            if (existingRow) {
                for (let c = 0; c < realHeaders.length; c++) {
                    const v = existingRow[c] !== undefined ? existingRow[c] : '';
                    if (v && v.toString().trim() !== '') {
                        isEmpty = false;
                        break;
                    }
                }
            }
        }
        
        // Apply filters
        if (isEmpty) return;
        if (hideSame && row1 && row2 && allSame) return;
        if (hideNewRows1 && row1 && !row2) return;
        if (hideNewRows2 && !row1 && row2) return;
        if (hideDiffRows && row1 && row2 && hasWarn) return;
        
        // For SIMILAR pairs create grouped rows (like in functions.js)
        if (pair.matchType === 'SIMILAR' && row1 && row2) {
            // Analyze columns for differences
            const columnComparisons = [];
            let hasAnyDifference = false;
            
            realHeaders.forEach((header, colIndex) => {
                const v1 = row1[colIndex] !== undefined ? row1[colIndex] : '';
                const v2 = row2[colIndex] !== undefined ? row2[colIndex] : '';
                
                // Use improved comparison function
                const compResult = compareValuesImproved(v1, v2, false);
                if (compResult === 'identical') {
                    columnComparisons[colIndex] = 'identical';
                } else {
                    columnComparisons[colIndex] = 'different';
                    hasAnyDifference = true;
                }
            });
            
            if (hasAnyDifference) {
                // First row (File 1)
                bodyHtml += `<tr class="warn-row warn-row-group-start" data-row-index="${pair.index1}">`;
                bodyHtml += `<td class="warn-cell">${file1Name}</td>`;
                
                realHeaders.forEach((header, colIndex) => {
                    const v1 = row1[colIndex] !== undefined ? row1[colIndex] : '';
                    const compResult = columnComparisons[colIndex];
                    
                    // Add diff column styling
                    let cellClass = compResult === 'identical' ? 'identical' : 'warn-cell';
                    if (onlyInFile1.includes(header)) {
                        cellClass += ' diff-column-cell only-in-file1';
                        if (!v1 || v1 === '') {
                            cellClass += ' empty-diff-column';
                        }
                    } else if (onlyInFile2.includes(header)) {
                        cellClass += ' diff-column-cell only-in-file2';
                        if (!v1 || v1 === '') {
                            cellClass += ' empty-diff-column';
                        }
                    }
                    
                    if (compResult === 'identical') {
                        // Merge identical columns
                        bodyHtml += `<td class="${cellClass}" rowspan="2" style="vertical-align: middle; text-align: center;">${v1}</td>`;
                    } else {
                        // Show different columns separately
                        bodyHtml += `<td class="${cellClass}">${v1}</td>`;
                    }
                });
                bodyHtml += '</tr>';
                
                // Second row (File 2)
                bodyHtml += `<tr class="warn-row warn-row-group-end" data-row-index="${pair.index2}">`;
                bodyHtml += `<td class="warn-cell">${file2Name}</td>`;
                
                realHeaders.forEach((header, colIndex) => {
                    const v2 = row2[colIndex] !== undefined ? row2[colIndex] : '';
                    const compResult = columnComparisons[colIndex];
                    
                    if (compResult === 'different') {
                        // Add diff column styling for file 2 rows
                        let cellClass = 'warn-cell';
                        if (onlyInFile1.includes(header)) {
                            cellClass += ' diff-column-cell only-in-file1';
                            if (!v2 || v2 === '') {
                                cellClass += ' empty-diff-column';
                            }
                        } else if (onlyInFile2.includes(header)) {
                            cellClass += ' diff-column-cell only-in-file2';
                            if (!v2 || v2 === '') {
                                cellClass += ' empty-diff-column';
                            }
                        }
                        
                        // Show only different columns (identical already shown with rowspan)
                        bodyHtml += `<td class="${cellClass}">${v2}</td>`;
                    }
                    // Skip identical columns as they are already displayed with rowspan="2"
                });
                bodyHtml += '</tr>';
            } else {
                // If no differences, show as identical
                bodyHtml += `<tr class="identical-row" data-row-index="${pair.index1}">`;
                bodyHtml += `<td class="file-both">Both files</td>`;
                
                realHeaders.forEach((header, colIndex) => {
                    const value = row1[colIndex] !== undefined ? row1[colIndex] : '';
                    
                    // Add diff column styling
                    let cellClass = 'identical';
                    if (onlyInFile1.includes(header)) {
                        cellClass += ' diff-column-cell only-in-file1';
                        if (!value || value === '') {
                            cellClass += ' empty-diff-column';
                        }
                    } else if (onlyInFile2.includes(header)) {
                        cellClass += ' diff-column-cell only-in-file2';
                        if (!value || value === '') {
                            cellClass += ' empty-diff-column';
                        }
                    }
                    
                    bodyHtml += `<td class="${cellClass}" title="${value}">${value}</td>`;
                });
                bodyHtml += `</tr>`;
            }
        } else {
            // For other types (IDENTICAL, ONLY_IN_TABLE1, ONLY_IN_TABLE2) use old logic
            const rowData = row1 || row2;
            let fileIndicator = '';
            
            // Determine source indicator depending on match type
            if (pair.matchType === 'IDENTICAL') {
                fileIndicator = 'Both files';
            } else if (pair.onlyIn === 'table1') {
                fileIndicator = file1Name;
            } else if (pair.onlyIn === 'table2') {
                fileIndicator = file2Name;
            }
            
            let rowClass = 'diff-row';
            let sourceClass = 'file-indicator';
            
            // Set styles depending on match type
            if (pair.matchType === 'IDENTICAL') {
                rowClass += ' identical-row';
                sourceClass += ' file-both';
            } else if (pair.onlyIn === 'table1') {
                rowClass += ' different-row';
                sourceClass += ' new-cell1';
            } else if (pair.onlyIn === 'table2') {
                rowClass += ' different-row';
                sourceClass += ' new-cell2';
            } else {
                // Fallback for other cases
                rowClass += ' identical-row';
                sourceClass += ' file-both';
            }
            
            bodyHtml += `<tr class="${rowClass}" data-row-index="${pair.index1 >= 0 ? pair.index1 : pair.index2}">`;
            bodyHtml += `<td class="${sourceClass}">${fileIndicator}</td>`;
            
            realHeaders.forEach((header, colIndex) => {
                const value = rowData && rowData[colIndex] !== undefined ? rowData[colIndex] : '';
                
                // Enhanced logic for determining cell style
                let cellClass = 'diff-cell';
                if (pair.matchType === 'IDENTICAL') {
                    cellClass += ' identical';
                } else {
                    cellClass += ' different';
                }
                
                // Add diff column styling
                if (onlyInFile1.includes(header)) {
                    cellClass += ' diff-column-cell only-in-file1';
                    if (!value || value === '') {
                        cellClass += ' empty-diff-column';
                    }
                } else if (onlyInFile2.includes(header)) {
                    cellClass += ' diff-column-cell only-in-file2';
                    if (!value || value === '') {
                        cellClass += ' empty-diff-column';
                    }
                }
                
                bodyHtml += `<td class="${cellClass}" title="${value}">${value}</td>`;
            });
            
            bodyHtml += `</tr>`;
        }
    });
    
    bodyTable.innerHTML = bodyHtml;
    
    const diffTableFinal = document.getElementById('diffTable');
    if (diffTableFinal) {
        diffTableFinal.style.display = 'block';
        diffTableFinal.style.visibility = 'visible';
    }

    // Add horizontal scroll synchronization
    syncTableScroll();
    
    // Add refreshTableLayout call for proper filter functionality
    if (typeof refreshTableLayout === 'function') {
        refreshTableLayout();
    }
    
    // Add syncColumnWidths call for column width synchronization
    if (typeof syncColumnWidths === 'function') {
        syncColumnWidths();
    }
}

window.MaxPilotDuckDB = {
    FastTableComparator,
    fastComparator,
    initializeFastComparator,
    createTableFromPreviewData,
    createPreviewFromTable,
    compareTablesWithFastComparator,
    compareTablesEnhanced,
    processFastComparisonResults,
    prepareDataForExportFast,
    benchmarkExportPerformance,
    runFastComparatorTests,
    testExportPerformance: () => window.testExportPerformance(),
    // Aggregation table methods
    getAggregationData: async () => {
        if (fastComparator && fastComparator.initialized) {
            return await fastComparator.getAggregationData();
        }
        return [];
    },
    clearAggregationData: async () => {
        if (fastComparator && fastComparator.initialized) {
            return await fastComparator.clearAggregationData();
        }
    },
    removeFileFromAggregation: async (fileId) => {
        if (fastComparator && fastComparator.initialized) {
            return await fastComparator.removeFileFromAggregation(fileId);
        }
    },
    addFileToAggregation: async (fileId, fileName, data, headers = null) => {
        if (fastComparator && fastComparator.initialized) {
            return await fastComparator.addFileToAggregation(fileId, fileName, data, headers);
        }
        return null;
    },
    get initialized() {
        return fastComparator && fastComparator.initialized;
    },
    version: '1.3.0'
};


document.addEventListener('DOMContentLoaded', async () => {
    await initializeFastComparator();
    
    if (typeof window.compareTables === 'function') {
        window.compareTablesOriginal = window.compareTables;
    }
    // Re-enable DuckDB WASM with enhanced logic
    window.compareTables = compareTablesEnhanced;
    
    // Add event handlers for checkbox filters
    const hideSameRowsEl = document.getElementById('hideSameRows');
    if (hideSameRowsEl) {
        hideSameRowsEl.addEventListener('change', function() {
            if (window.currentPairs && window.currentPairs.length > 0 && window.currentFinalHeaders) {
                createBasicFallbackTable(window.currentPairs, window.currentFinalHeaders);
            }
        });
    }
    
    const hideDiffColumnsEl = document.getElementById('hideDiffColumns');
    if (hideDiffColumnsEl) {
        hideDiffColumnsEl.addEventListener('change', function() {
            if (window.currentPairs && window.currentPairs.length > 0 && window.currentFinalHeaders) {
                createBasicFallbackTable(window.currentPairs, window.currentFinalHeaders);
            }
        });
    }
    
    const hideNewRows1El = document.getElementById('hideNewRows1');
    if (hideNewRows1El) {
        hideNewRows1El.addEventListener('change', function() {
            if (window.currentPairs && window.currentPairs.length > 0 && window.currentFinalHeaders) {
                createBasicFallbackTable(window.currentPairs, window.currentFinalHeaders);
            }
        });
    }
    
    const hideNewRows2El = document.getElementById('hideNewRows2');
    if (hideNewRows2El) {
        hideNewRows2El.addEventListener('change', function() {
            if (window.currentPairs && window.currentPairs.length > 0 && window.currentFinalHeaders) {
                createBasicFallbackTable(window.currentPairs, window.currentFinalHeaders);
            }
        });
    }
});

// Horizontal scroll synchronization functions
function syncTableScroll() {
    const headerTable = document.querySelector('.table-header-fixed');
    const bodyTable = document.querySelector('.table-body-scrollable');
    
    if (headerTable && bodyTable) {
        bodyTable.removeEventListener('scroll', syncScrollHandler);
        bodyTable.addEventListener('scroll', syncScrollHandler);
    }
}

function syncScrollHandler() {
    const headerTable = document.querySelector('.table-header-fixed');
    const bodyTable = document.querySelector('.table-body-scrollable');
    
    if (headerTable && bodyTable) {
        headerTable.scrollLeft = bodyTable.scrollLeft;
    }
}

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

async function benchmarkExportPerformance() {
    if (!window.currentFastResult) {
        return null;
    }
    
    const { table1Count, table2Count, identical, onlyInTable1, onlyInTable2 } = window.currentFastResult;
    const totalDataRows = table1Count + table2Count;
    const exportRows = identical.length + onlyInTable1.length + onlyInTable2.length;
    
    // Estimate expected performance improvement
    const expectedSpeedup = fastComparator && fastComparator.initialized ? 
        (exportRows > 10000 ? '3-5x faster' : 
         exportRows > 1000 ? '2-3x faster' : 
         '1.5-2x faster') : 
        'Standard speed';
    
    return {
        totalDataRows,
        exportRows,
        fastModeAvailable: !!(fastComparator && fastComparator.initialized),
        mode: fastComparator?.mode || 'standard',
        expectedSpeedup
    };
}

// Utility function to test export performance
window.testExportPerformance = async function() {
    
    if (!window.currentFastResult) {
        console.log('❌ No comparison result available. Please run a comparison first.');
        return;
    }
    
    const benchmark = await benchmarkExportPerformance();
    if (!benchmark) return;
    
    const startTime = performance.now();
    const testData = await prepareDataForExportFast(window.currentFastResult, false);
    const endTime = performance.now();
    
    const duration = endTime - startTime;
    const rowsPerSecond = Math.round(benchmark.exportRows / (duration / 1000));
    
    return {
        duration,
        rowsPerSecond,
        exportRows: benchmark.exportRows,
        efficiency: rowsPerSecond > 10000 ? 'Excellent' : 
                   rowsPerSecond > 5000 ? 'Good' : 
                   rowsPerSecond > 1000 ? 'Fair' : 'Needs improvement'
    };
};

async function prepareDataForExportFast(fastResult, useTolerance = false) {
    if (!fastResult || !fastComparator || !fastComparator.initialized) {
        return null;
    }

    const startTime = performance.now();
    const { identical, similar, onlyInTable1, onlyInTable2, commonColumns, alignedData1, alignedData2 } = fastResult;
    
    const workingData1 = alignedData1 || data1;
    const workingData2 = alignedData2 || data2;
    
    const headers = ['Source'];
    const realHeaders = workingData1[0] || commonColumns;
    realHeaders.forEach(header => headers.push(String(header || '')));
    
    const data = [headers];
    const formatting = {};
    const colWidths = [];
    
    // Set column widths efficiently
    colWidths.push({ wch: 20 }); // Source column
    for (let i = 0; i < realHeaders.length; i++) {
        colWidths.push({ wch: 15 });
    }
    
    // Pre-create header formatting
    const headerFormatting = {
        fill: { fgColor: { rgb: 'f8f9fa' } },
        font: { bold: true, color: { rgb: '212529' } },
        border: {
            top: { style: 'thin', color: { rgb: 'D4D4D4' } },
            bottom: { style: 'thin', color: { rgb: 'D4D4D4' } },
            left: { style: 'thin', color: { rgb: 'D4D4D4' } },
            right: { style: 'thin', color: { rgb: 'D4D4D4' } }
        }
    };
    
    // Apply header formatting efficiently
    for (let col = 0; col < headers.length; col++) {
        formatting[XLSX.utils.encode_cell({ r: 0, c: col })] = { ...headerFormatting };
    }
    
    let rowIndex = 1;
    // Use getFileDisplayName to display filename with sheet
    const file1Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName1 || 'File 1', window.sheetName1 || '')
        : (window.fileName1 || 'File 1');
    const file2Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName2 || 'File 2', window.sheetName2 || '')
        : (window.fileName2 || 'File 2');
    
    // Pre-create formatting templates for different row types
    const identicalFormatting = {
        fill: { fgColor: { rgb: 'd4edda' } },
        font: { color: { rgb: '212529' } },
        border: {
            top: { style: 'thin', color: { rgb: 'D4D4D4' } },
            bottom: { style: 'thin', color: { rgb: 'D4D4D4' } },
            left: { style: 'thin', color: { rgb: 'D4D4D4' } },
            right: { style: 'thin', color: { rgb: 'D4D4D4' } }
        }
    };
    
    const sourceIdenticalFormatting = {
        ...identicalFormatting,
        font: { color: { rgb: '212529' }, bold: true }
    };
    
    const table1Formatting = {
        fill: { fgColor: { rgb: '65add7' } },
        font: { color: { rgb: '212529' } },
        border: {
            top: { style: 'thin', color: { rgb: 'D4D4D4' } },
            bottom: { style: 'thin', color: { rgb: 'D4D4D4' } },
            left: { style: 'thin', color: { rgb: 'D4D4D4' } },
            right: { style: 'thin', color: { rgb: 'D4D4D4' } }
        }
    };
    
    const sourceTable1Formatting = {
        ...table1Formatting,
        font: { color: { rgb: '212529' }, bold: true }
    };
    
    const table2Formatting = {
        fill: { fgColor: { rgb: '63cfbf' } },
        font: { color: { rgb: '212529' } },
        border: {
            top: { style: 'thin', color: { rgb: 'D4D4D4' } },
            bottom: { style: 'thin', color: { rgb: 'D4D4D4' } },
            left: { style: 'thin', color: { rgb: 'D4D4D4' } },
            right: { style: 'thin', color: { rgb: 'D4D4D4' } }
        }
    };
    
    const sourceTable2Formatting = {
        ...table2Formatting,
        font: { color: { rgb: '212529' }, bold: true }
    };
    
    // Batch processing for better performance
    const BATCH_SIZE = 1000;
    
    // Process ALL identical rows (remove artificial limit)
    const maxIdentical = identical.length;
    
    for (let batchStart = 0; batchStart < maxIdentical; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, maxIdentical);
        const batch = identical.slice(batchStart, batchEnd);
        
        batch.forEach(identicalPair => {
            // Fix index calculation - DuckDB returns data-relative indices, add 1 to skip header
            const row1Index = identicalPair.row1 + 1; // Add 1 to skip header row
            const row1 = workingData1[row1Index];
            
            if (!row1) {
                console.log('⚠️ Missing row1 for identical pair:', { row1Index, identicalPair, workingData1Length: workingData1.length });
                return;
            }
            
            const dataRow = ['Both files'];
            for (let c = 0; c < realHeaders.length; c++) {
                const value = row1[c] !== undefined ? String(row1[c]) : '';
                dataRow.push(value);
                
                formatting[XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 })] = { ...identicalFormatting };
            }
            
            data.push(dataRow);
            formatting[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })] = { ...sourceIdenticalFormatting };
            rowIndex++;
        });
        
        // Allow UI breathing room for large datasets
        if (batchEnd < maxIdentical && (batchEnd % (BATCH_SIZE * 5)) === 0) {
            await new Promise(resolve => setTimeout(resolve, 1));
        }
    }
    
    // Process similar/different rows (show both versions)
    if (similar && similar.length > 0) {
        
        // Pre-create formatting templates for different row types
        const differentFormatting = {
            // White background for normal cells (no fill specified)
            font: { color: { rgb: '212529' } },
            border: {
                top: { style: 'thick', color: { rgb: '000000' } }, // 2pt top border for first row of pair
                bottom: { style: 'thin', color: { rgb: 'D4D4D4' } },
                left: { style: 'thin', color: { rgb: 'D4D4D4' } },
                right: { style: 'thin', color: { rgb: 'D4D4D4' } }
            }
        };
        
        const differentFormattingSecond = {
            // White background for normal cells (no fill specified)
            font: { color: { rgb: '212529' } },
            border: {
                top: { style: 'thin', color: { rgb: 'D4D4D4' } },
                bottom: { style: 'thick', color: { rgb: '000000' } }, // 2pt bottom border for second row of pair
                left: { style: 'thin', color: { rgb: 'D4D4D4' } },
                right: { style: 'thin', color: { rgb: 'D4D4D4' } }
            }
        };
        
        const sourceDifferentFormatting = {
            fill: { fgColor: { rgb: 'f8d7da' } }, // Light red for source column
            font: { color: { rgb: '212529' }, bold: true },
            border: {
                top: { style: 'thick', color: { rgb: '000000' } },
                bottom: { style: 'thin', color: { rgb: 'D4D4D4' } },
                left: { style: 'thin', color: { rgb: 'D4D4D4' } },
                right: { style: 'thin', color: { rgb: 'D4D4D4' } }
            }
        };
        
        const sourceDifferentFormattingSecond = {
            fill: { fgColor: { rgb: 'f8d7da' } }, // Light red for source column
            font: { color: { rgb: '212529' }, bold: true },
            border: {
                top: { style: 'thin', color: { rgb: 'D4D4D4' } },
                bottom: { style: 'thick', color: { rgb: '000000' } },
                left: { style: 'thin', color: { rgb: 'D4D4D4' } },
                right: { style: 'thin', color: { rgb: 'D4D4D4' } }
            }
        };
        
        for (let batchStart = 0; batchStart < similar.length; batchStart += BATCH_SIZE) {
            const batchEnd = Math.min(batchStart + BATCH_SIZE, similar.length);
            const batch = similar.slice(batchStart, batchEnd);
            
            batch.forEach(similarPair => {
                // Fix index calculation - DuckDB returns data-relative indices, add 1 to skip header
                const row1Index = similarPair.row1 + 1; // Add 1 to skip header row
                const row1 = workingData1[row1Index];
                
                const row2Index = similarPair.row2 + 1; // Add 1 to skip header row  
                const row2 = workingData2[row2Index];
                
                if (row1 && row2) {
                    // Add row from file 1 (first row of pair - thick top border)
                    const dataRow1 = [file1Name];
                    for (let c = 0; c < realHeaders.length; c++) {
                        const value1 = row1[c] !== undefined ? String(row1[c]) : '';
                        const value2 = row2[c] !== undefined ? String(row2[c]) : '';
                        
                        dataRow1.push(value1);
                        
                        // Check if this cell differs from corresponding cell in file 2
                        const isDifferent = value1.toLowerCase().trim() !== value2.toLowerCase().trim();
                        
                        if (isDifferent) {
                            // Red formatting for different cells (with thick top border)
                            formatting[XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 })] = {
                                fill: { fgColor: { rgb: 'f8d7da' } }, // Light red background
                                font: { color: { rgb: '212529' } },
                                border: {
                                    top: { style: 'thick', color: { rgb: '000000' } },
                                    bottom: { style: 'thin', color: { rgb: 'D4D4D4' } },
                                    left: { style: 'thin', color: { rgb: 'D4D4D4' } },
                                    right: { style: 'thin', color: { rgb: 'D4D4D4' } }
                                }
                            };
                        } else {
                            // White background for same cells (with thick top border)
                            formatting[XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 })] = { ...differentFormatting };
                        }
                    }
                    
                    data.push(dataRow1);
                    formatting[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })] = { ...sourceDifferentFormatting };
                    rowIndex++;
                    
                    // Add row from file 2 (second row of pair - thick bottom border)
                    const dataRow2 = [file2Name];
                    for (let c = 0; c < realHeaders.length; c++) {
                        const value1 = row1[c] !== undefined ? String(row1[c]) : '';
                        const value2 = row2[c] !== undefined ? String(row2[c]) : '';
                        
                        dataRow2.push(value2);
                        
                        // Check if this cell differs from corresponding cell in file 1
                        const isDifferent = value1.toLowerCase().trim() !== value2.toLowerCase().trim();
                        
                        if (isDifferent) {
                            // Red formatting for different cells (with thick bottom border)
                            formatting[XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 })] = {
                                fill: { fgColor: { rgb: 'f8d7da' } }, // Light red background
                                font: { color: { rgb: '212529' } },
                                border: {
                                    top: { style: 'thin', color: { rgb: 'D4D4D4' } },
                                    bottom: { style: 'thick', color: { rgb: '000000' } },
                                    left: { style: 'thin', color: { rgb: 'D4D4D4' } },
                                    right: { style: 'thin', color: { rgb: 'D4D4D4' } }
                                }
                            };
                        } else {
                            // White background for same cells (with thick bottom border)
                            formatting[XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 })] = { ...differentFormattingSecond };
                        }
                    }
                    
                    data.push(dataRow2);
                    formatting[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })] = { ...sourceDifferentFormattingSecond };
                    rowIndex++;
                } else {
                    if (!row1) {
                        console.log('⚠️ Missing row1 for similar pair:', { row1Index, similarPair, workingData1Length: workingData1.length });
                    }
                    if (!row2) {
                        console.log('⚠️ Missing row2 for similar pair:', { row2Index, similarPair, workingData2Length: workingData2.length });
                    }
                }
            });
            
            // Allow UI breathing room for large datasets
            if (batchEnd < similar.length && (batchEnd % (BATCH_SIZE * 5)) === 0) {
                await new Promise(resolve => setTimeout(resolve, 1));
            }
        }
    }
    
    // Process table1-only rows efficiently
    for (let batchStart = 0; batchStart < onlyInTable1.length; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, onlyInTable1.length);
        const batch = onlyInTable1.slice(batchStart, batchEnd);
        
        batch.forEach(diff => {
            const row1Index = diff.row1 + 1; // Add 1 to skip header row
            const row1 = workingData1[row1Index];
            
            if (!row1) {
                console.log('⚠️ Missing row1 for table1-only:', { row1Index, diff, workingData1Length: workingData1.length });
                return;
            }
            
            const dataRow = [file1Name];
            for (let c = 0; c < realHeaders.length; c++) {
                const value = row1[c] !== undefined ? String(row1[c]) : '';
                dataRow.push(value);
                
                formatting[XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 })] = { ...table1Formatting };
            }
            
            data.push(dataRow);
            formatting[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })] = { ...sourceTable1Formatting };
            rowIndex++;
        });
        
        if (batchEnd < onlyInTable1.length && (batchEnd % (BATCH_SIZE * 5)) === 0) {
            await new Promise(resolve => setTimeout(resolve, 1));
        }
    }
    
    // Process table2-only rows efficiently
    for (let batchStart = 0; batchStart < onlyInTable2.length; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, onlyInTable2.length);
        const batch = onlyInTable2.slice(batchStart, batchEnd);
        
        batch.forEach(diff => {
            const row2Index = diff.row2 + 1; // Add 1 to skip header row
            const row2 = workingData2[row2Index];
            
            if (!row2) {
                console.log('⚠️ Missing row2 for table2-only:', { row2Index, diff, workingData2Length: workingData2.length });
                return;
            }
            
            const dataRow = [file2Name];
            for (let c = 0; c < realHeaders.length; c++) {
                const value = row2[c] !== undefined ? String(row2[c]) : '';
                dataRow.push(value);
                
                formatting[XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 })] = { ...table2Formatting };
            }
            
            data.push(dataRow);
            formatting[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })] = { ...sourceTable2Formatting };
            rowIndex++;
        });
        
        if (batchEnd < onlyInTable2.length && (batchEnd % (BATCH_SIZE * 5)) === 0) {
            await new Promise(resolve => setTimeout(resolve, 1));
        }
    }
    
    const duration = performance.now() - startTime;
    console.log(`⚡ Fast export completed in ${duration.toFixed(2)}ms - prepared ${data.length - 1} rows for export (1 header + ${rowIndex - 1} data rows)`);
    return { data, formatting, colWidths };
}

}