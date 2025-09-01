/**
 * Enhanced comparison functions using DuckDB WASM
 * Extends the original functions.js with high-performance database operations
 */

// Enhanced comparison function that tries DuckDB first, fallback to original
async function compareTablesEnhanced(useTolerance = false) {
    console.log('Starting enhanced comparison...');
    
    // Clear previous results
    clearComparisonResults();
    
    let resultDiv = document.getElementById('result');
    let summaryDiv = document.getElementById('summary');
    
    resultDiv.innerHTML = '<div id="comparison-loading" style="text-align: center; padding: 20px; font-size: 16px;">⚡ Starting fast comparison... Please wait</div>';
    summaryDiv.innerHTML = '<div style="text-align: center; padding: 10px;">Initializing DuckDB...</div>';
    
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
        // Try Local DuckDB Simulator first (no CORS issues)
        if (window.localDuckDBManager && window.localDuckDBManager.initialized) {
            resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚡ Using Local DuckDB fast comparison...</div>';
            
            const excludedColumns = getExcludedColumns();
            const tolerance = useTolerance ? 1.5 : 0;
            
            const localResult = await compareTablesWithLocalDuckDB(
                data1, data2, excludedColumns, useTolerance, tolerance
            );

            if (localResult) {
                console.log('✅ Local DuckDB comparison successful, processing results...');
                await processDuckDBResults(localResult, useTolerance);
                return;
            }
        }

        // Try full DuckDB WASM if available
        if (duckDBManager && duckDBManager.initialized) {
            resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚡ Using DuckDB WASM fast comparison...</div>';
            
            const excludedColumns = getExcludedColumns();
            const tolerance = useTolerance ? 1.5 : 0;
            
            const duckResult = await compareTablesWithDuckDB(
                data1, data2, excludedColumns, useTolerance, tolerance
            );

            if (duckResult) {
                console.log('✅ DuckDB WASM comparison successful, processing results...');
                await processDuckDBResults(duckResult, useTolerance);
                return;
            }
        }

        // Fallback to original comparison
        console.log('ℹ️ Falling back to original comparison method...');
        resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">🔄 Using standard comparison...</div>';
        
        setTimeout(() => {
            performComparison();
        }, 10);
        
    } catch (error) {
        console.error('Enhanced comparison failed:', error);
        
        // Fallback to original method
        resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">🔄 Using standard comparison...</div>';
        setTimeout(() => {
            performComparison();
        }, 10);
    }
}

async function processDuckDBResults(duckResult, useTolerance) {
    const { identical, onlyInTable1, onlyInTable2, table1Count, table2Count, commonColumns } = duckResult;
    
    console.log('Processing DuckDB results...', {
        identical: identical.length,
        onlyInTable1: onlyInTable1.length, 
        onlyInTable2: onlyInTable2.length,
        table1Count,
        table2Count
    });

    // Calculate statistics
    const identicalCount = identical.length;
    const diffCount = table1Count + table2Count - identicalCount - onlyInTable1.length - onlyInTable2.length;
    const similarity = table1Count > 0 ? ((identicalCount / Math.max(table1Count, table2Count)) * 100).toFixed(1) : 0;

    // Update summary
    updateSummaryFromDuckDB({
        file1Count: table1Count,
        file2Count: table2Count, 
        identicalCount,
        onlyInFile1: onlyInTable1.length,
        onlyInFile2: onlyInTable2.length,
        similarity,
        diffColumnsCount1: 0, // Will be calculated if needed
        diffColumnsCount2: 0  // Will be calculated if needed
    });

    // Generate comparison table
    await generateComparisonTableFromDuckDB(duckResult, useTolerance);
    
    // Show filter controls
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'flex';
    }

    // Show export button
    const exportBtn = document.getElementById('exportExcelBtn');
    if (exportBtn) {
        exportBtn.style.display = 'block';
    }
}

function updateSummaryFromDuckDB(stats) {
    const tableHeaders = getSummaryTableHeaders();
    
    const summaryHTML = `
        <div style="overflow-x: auto; margin: 20px 0;">
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>${tableHeaders.file}</th>
                        <th>${tableHeaders.rowCount}</th>
                        <th>${tableHeaders.rowsOnlyInFile}</th>
                        <th>${tableHeaders.identicalRows}</th>
                        <th>${tableHeaders.similarity}</th>
                        <th>${tableHeaders.diffColumns}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>${fileName1}</strong></td>
                        <td>${stats.file1Count.toLocaleString()}</td>
                        <td>${stats.onlyInFile1.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 18px;">${stats.identicalCount.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 18px;">${stats.similarity}%</td>
                        <td>${stats.diffColumnsCount1}</td>
                    </tr>
                    <tr>
                        <td><strong>${fileName2}</strong></td>
                        <td>${stats.file2Count.toLocaleString()}</td>
                        <td>${stats.onlyInFile2.toLocaleString()}</td>
                        <td>${stats.diffColumnsCount2}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    `;
    
    document.getElementById('summary').innerHTML = summaryHTML;
}

async function generateComparisonTableFromDuckDB(duckResult, useTolerance) {
    const { identical, onlyInTable1, onlyInTable2 } = duckResult;
    
    // For now, we'll create a simplified table showing the comparison results
    // In a full implementation, you'd want to merge this with the detailed row-by-row comparison
    
    const maxRowsToShow = Math.min(1000, Math.max(data1.length, data2.length));
    const headers = data1[0] || [];
    
    // Create comparison pairs
    const allPairs = [];
    
    // Add identical pairs
    identical.forEach(match => {
        if (match.row1 <= data1.length && match.row2 <= data2.length) {
            allPairs.push({
                row1Index: match.row1 - 1,
                row2Index: match.row2 - 1,
                status: 'identical'
            });
        }
    });
    
    // Add only in table1 pairs
    onlyInTable1.forEach(match => {
        if (match.row1 <= data1.length) {
            allPairs.push({
                row1Index: match.row1 - 1,
                row2Index: -1,
                status: 'only_in_table1'
            });
        }
    });
    
    // Add only in table2 pairs  
    onlyInTable2.forEach(match => {
        if (match.row2 <= data2.length) {
            allPairs.push({
                row1Index: -1,
                row2Index: match.row2 - 1,
                status: 'only_in_table2'
            });
        }
    });

    // Sort pairs by row index
    allPairs.sort((a, b) => {
        const aIndex = a.row1Index >= 0 ? a.row1Index : a.row2Index;
        const bIndex = b.row1Index >= 0 ? b.row1Index : b.row2Index;
        return aIndex - bIndex;
    });

    // Store for global access
    currentPairs = allPairs.slice(0, maxRowsToShow);
    currentFinalHeaders = headers;
    
    // Generate the visual table
    generateComparisonTable(currentPairs, headers, useTolerance);
    
    console.log(`Generated comparison table with ${currentPairs.length} rows`);
}

// Enhanced file processing with DuckDB preparation
async function processFileWithDuckDB(file, fileNum) {
    console.log(`Processing file ${fileNum} with DuckDB enhancement:`, file.name);
    
    // First, process the file normally
    processFile(file, fileNum);
    
    // Wait a bit for the normal processing to complete
    await new Promise(resolve => setTimeout(resolve, 100));
    
    // If DuckDB is available and we have data, prepare it in the database
    if (duckDBManager && duckDBManager.initialized) {
        try {
            const data = fileNum === 1 ? data1 : data2;
            if (data && data.length > 0) {
                const tableName = `file${fileNum}_data`;
                await duckDBManager.createTableFromData(tableName, data);
                console.log(`File ${fileNum} data loaded into DuckDB table: ${tableName}`);
            }
        } catch (error) {
            console.error(`Failed to load file ${fileNum} into DuckDB:`, error);
        }
    }
}

// Utility function to get excluded columns (from original code)
function getExcludedColumns() {
    const excludeInput = document.getElementById('excludeColumns');
    if (!excludeInput || !excludeInput.value.trim()) {
        return [];
    }
    
    return excludeInput.value
        .split(',')
        .map(col => col.trim())
        .filter(col => col.length > 0);
}

// Override the original compare function to use enhanced version
if (typeof window !== 'undefined') {
    window.compareTablesOriginal = window.compareTables || function() {};
    window.compareTables = compareTablesEnhanced;
    
    // Initialize local DuckDB when this script loads
    window.addEventListener('DOMContentLoaded', async () => {
        console.log('🔧 Initializing DuckDB Enhanced mode...');
        
        // Try to initialize local simulator as fallback
        if (typeof initializeLocalDuckDB === 'function') {
            await initializeLocalDuckDB();
        }
        
        // Update UI indicators
        setTimeout(() => {
            updateFastModeIndicators();
        }, 1000);
    });
}

// Update fast mode indicators in UI
function updateFastModeIndicators() {
    const indicators = document.querySelectorAll('.fast-mode-indicator');
    const isAvailable = (window.duckDBManager && window.duckDBManager.initialized) || 
                       (window.localDuckDBManager && window.localDuckDBManager.initialized);
    
    indicators.forEach(indicator => {
        if (isAvailable) {
            indicator.style.display = 'inline-block';
            indicator.textContent = window.localDuckDBManager ? 'FAST' : 'ULTRA';
        } else {
            indicator.style.display = 'none';
        }
    });
    
    // Update DuckDB status
    const statusEl = document.getElementById('duckdb-status');
    if (statusEl && isAvailable) {
        const mode = window.localDuckDBManager ? 'Local Fast Mode' : 'DuckDB WASM Mode';
        statusEl.innerHTML = `⚡ ${mode} enabled - Enhanced performance!`;
        statusEl.className = 'duckdb-status duckdb-available show';
    }
}
