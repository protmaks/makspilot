async function exportToExcel() {
    // Check if we have data to export - either from standard comparison or fast comparison
    const hasStandardData = currentPairs && currentPairs.length > 0;
    const hasFastData = window.currentFastResult && (
        (window.currentFastResult.identical && window.currentFastResult.identical.length > 0) ||
        (window.currentFastResult.similar && window.currentFastResult.similar.length > 0) ||
        (window.currentFastResult.onlyInTable1 && window.currentFastResult.onlyInTable1.length > 0) ||
        (window.currentFastResult.onlyInTable2 && window.currentFastResult.onlyInTable2.length > 0)
    );
    
    if (!hasStandardData && !hasFastData) {
        alert('No data to export. Please compare files first.');
        return;
    }
    
    // Check if we need to perform full comparison first
    if (window.isQuickMode) {
        console.log('🔄 Quick mode detected - performing full comparison for export');
        
        const exportBtn = document.getElementById('exportExcelBtn');
        const originalText = exportBtn ? exportBtn.innerHTML : '';
        
        if (exportBtn) {
            exportBtn.innerHTML = '⚡ Performing full comparison...';
            exportBtn.disabled = true;
        }
        
        try {
            // Perform full comparison
            await performFullComparisonForExport();
        } catch (error) {
            console.error('Full comparison failed:', error);
            alert('Error performing full comparison: ' + error.message);
            if (exportBtn) {
                exportBtn.innerHTML = originalText;
                exportBtn.disabled = false;
            }
            return;
        }
        
        // Reset quick mode flag
        window.isQuickMode = false;
        
        if (exportBtn) {
            exportBtn.innerHTML = '📊 Exporting to Excel...';
        }
    }
    
    // Run performance benchmark
    if (window.MaxPilotDuckDB && window.MaxPilotDuckDB.benchmarkExportPerformance) {
        await window.MaxPilotDuckDB.benchmarkExportPerformance();
    }
    
    try {
        // Try fast export first if available
        if (window.currentFastResult && window.MaxPilotDuckDB && window.MaxPilotDuckDB.prepareDataForExportFast) {
            console.log('⚡ Using fast export engine');
            await exportWithFastEngine();
        } else {
            console.log('🔄 Using standard export method - will use full data if available');
            // Fallback to old method - will use full comparison data if available
            const exportData = prepareDataForExport();
            if (!exportData) {
                throw new Error('No data available for export from standard method');
            }
            const htmlContent = await createStyledHTMLTable(exportData);
            downloadExcelFromHTML(htmlContent);
        }
    } catch (error) {
        console.error('Export error:', error);
        alert('Error exporting to Excel: ' + error.message);
    }
}

async function exportWithFastEngine() {
    try {
        const exportBtn = document.getElementById('exportExcelBtn');
        const originalText = exportBtn ? exportBtn.innerHTML : '';
        const startTime = performance.now();
        
        if (exportBtn) {
            exportBtn.innerHTML = '⚡ Fast export processing...';
            exportBtn.disabled = true;
        }
        
        // Show progress for large files
        const totalRows = window.currentFastResult ? 
            (window.currentFastResult.table1Count + window.currentFastResult.table2Count) : 0;
        
        console.log(`⚡ Fast export starting for ${totalRows.toLocaleString()} total rows`);
        
        if (totalRows > 10000 && exportBtn) {
            exportBtn.innerHTML = `⚡ Fast export - processing ${totalRows.toLocaleString()} rows...`;
        }
        
        const fastExportData = await window.MaxPilotDuckDB.prepareDataForExportFast(window.currentFastResult, window.toleranceMode || false);
        
        if (fastExportData) {
            const dataProcessingTime = performance.now() - startTime;
            console.log(`⚡ Data processing completed in ${dataProcessingTime.toFixed(2)}ms`);
            
            if (exportBtn) {
                exportBtn.innerHTML = '⚡ Generating Excel file...';
            }
            
            const htmlContent = await createStyledHTMLTable(fastExportData);
            downloadExcelFromHTML(htmlContent);
            
            const totalTime = performance.now() - startTime;
            console.log(`⚡ Total fast export time: ${totalTime.toFixed(2)}ms`);
            
            // Show success message
            showFastExportSuccess(fastExportData.data.length - 1, totalTime); // -1 for header
            
        } else {
            throw new Error('Fast export failed, falling back to standard method');
        }
        
        if (exportBtn) {
            exportBtn.innerHTML = originalText;
            exportBtn.disabled = false;
        }
        
    } catch (error) {
        console.error('Fast export failed:', error);
        
        const exportBtn = document.getElementById('exportExcelBtn');
        if (exportBtn) {
            exportBtn.innerHTML = '🔄 Using standard export...';
        }
        
        // Fallback to old method
        setTimeout(async () => {
            const exportData = prepareDataForExport();
            const htmlContent = await createStyledHTMLTable(exportData);
            downloadExcelFromHTML(htmlContent);
            
            if (exportBtn) {
                exportBtn.innerHTML = 'Export to Excel';
                exportBtn.disabled = false;
            }
        }, 100);
    }
}

function showFastExportSuccess(rowCount, totalTime = null) {
    const successDiv = document.createElement('div');
    successDiv.style.cssText = 'position:fixed;top:20px;right:20px;background:#d4edda;color:#155724;padding:15px 25px;border:1px solid #c3e6cb;border-radius:8px;box-shadow:0 4px 15px rgba(0,0,0,0.15);z-index:10000;font-size:14px;font-weight:500;max-width:350px;line-height:1.4;';
    
    let message = `⚡ <strong>Fast Excel export completed!</strong><br>Processed ${rowCount.toLocaleString()} rows with enhanced performance.`;
    
    if (totalTime) {
        const timeText = totalTime < 1000 ? `${totalTime.toFixed(0)}ms` : `${(totalTime/1000).toFixed(1)}s`;
        message += `<br>Total time: ${timeText}`;
    }
    
    successDiv.innerHTML = message;
    document.body.appendChild(successDiv);
    setTimeout(() => { successDiv.parentNode && successDiv.parentNode.removeChild(successDiv); }, 6000);
}

async function createStyledHTMLTable(exportData) {
    const startTime = performance.now();
    console.log('⚡ Generating HTML for Excel export...', { rows: exportData.data.length, columns: exportData.data[0]?.length || 0 });
    
    // Pre-build HTML parts for better performance
    const htmlParts = [];
    
    htmlParts.push(`
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
        <meta charset="utf-8">
        <style>
            .thick-top { border-top: 2pt solid black !important; }
            .thick-bottom { border-bottom: 2pt solid black !important; }
            .thin-border { border: 0.5pt solid #D4D4D4; }
            .text-format { mso-number-format:"@"; mso-data-type:string; white-space: pre; }
            td { mso-number-format:"@"; white-space: pre; mso-data-type: string; }
            table td { mso-number-format: "@"; mso-data-type: string; }
        </style>
        <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Comparison Results</x:Name><x:WorksheetOptions><x:DisplayGridlines/><x:Selected/><x:DefaultRowHeight>255</x:DefaultRowHeight><x:Panes><x:Pane><x:Number>0</x:Number><x:ActiveRow>0</x:ActiveRow><x:ActiveCol>0</x:ActiveCol></x:Pane></x:Panes></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
    </head>
    <body>
        <table border="0" style="border-collapse: collapse;" x:publishsource="Excel">
            <!-- XML Schema directive to force all columns as text -->
            <colgroup>
                <col style="mso-number-format:'@';" />
            </colgroup>
    `);
    
    // Process rows in batches for better performance
    const BATCH_SIZE = 500;
    const totalRows = exportData.data.length;
    
    for (let batchStart = 0; batchStart < totalRows; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, totalRows);
        const batchRows = [];
        
        for (let rowIndex = batchStart; rowIndex < batchEnd; rowIndex++) {
            const row = exportData.data[rowIndex];
            const isHeader = rowIndex === 0;
            
            let rowHtml = isHeader ? '<tr style="font-weight: bold;">' : '<tr>';
            
            row.forEach((cellValue, colIndex) => {
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
                const formatting = exportData.formatting[cellAddress];
                
                let style = 'border: 0.5pt solid #D4D4D4; padding: 4px; vertical-align: top;';
                
                if (formatting) {
                    if (formatting.border && formatting.border.top && formatting.border.top.style === 'thick') {
                        style = 'border-top: 2pt solid #000; border-bottom: 0.5pt solid #D4D4D4; border-left: 0.5pt solid #D4D4D4; border-right: 0.5pt solid #D4D4D4; padding:4px; vertical-align:top;';
                    } else if (formatting.border && formatting.border.bottom && formatting.border.bottom.style === 'thick') {
                        style = 'border-top: 0.5pt solid #D4D4D4; border-bottom: 2pt solid #000; border-left: 0.5pt solid #D4D4D4; border-right: 0.5pt solid #D4D4D4; padding:4px; vertical-align:top;';
                    }
                    
                    if (formatting.fill && formatting.fill.fgColor) {
                        let bgColor = formatting.fill.fgColor.rgb;
                        switch (bgColor) {
                            case 'FF6B6B': bgColor = 'f8d7da'; break;
                            case '4CAF50': bgColor = 'd4edda'; break;
                            case 'ffeaa7': bgColor = 'ffeaa7'; break;
                            case '65add7': bgColor = '65add7'; break;
                            case '63cfbf': bgColor = '63cfbf'; break;
                            case 'd4edda': bgColor = 'd4edda'; break;
                            case 'f8f9fa': bgColor = 'f8f9fa'; break;
                        }
                        style += ` background-color:#${bgColor};`;
                    }
                    
                    if (formatting.font) {
                        if (formatting.font.bold) style += ' font-weight:bold;';
                        style += ' color:#212529;';
                    }
                }
                
                const safeValue = String(cellValue || '').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
                
                if (isHeader) {
                    rowHtml += `<td style="${style}" class="thin-border" x:autofilter="all">${safeValue}</td>`;
                } else {
                    const textValue = ' \t' + safeValue;
                    rowHtml += `<td style="${style}mso-number-format:'@';mso-data-type:string;white-space:pre;" class="thin-border text-format" x:str="${safeValue}" x:format="@">${textValue}</td>`;
                }
            });
            
            rowHtml += '</tr>';
            batchRows.push(rowHtml);
        }
        
        htmlParts.push(batchRows.join(''));
        
        // Allow UI breathing room for very large files
        if (batchEnd < totalRows && (batchEnd % (BATCH_SIZE * 4)) === 0) {
            await new Promise(resolve => setTimeout(resolve, 1));
        }
    }
    
    htmlParts.push(`
        </table>
    </body>
    </html>
    `);
    
    const html = htmlParts.join('');
    const duration = performance.now() - startTime;
    console.log(`⚡ HTML generation completed in ${duration.toFixed(2)}ms for ${totalRows} rows`);
    
    return html;
}

function downloadExcelFromHTML(htmlContent) {
    const now = new Date();
    const date = now.toISOString().slice(0, 10);
    
    let firstFileName = 'file1';
    if (typeof fileName1 !== 'undefined' && fileName1) {
        firstFileName = fileName1.replace(/\.[^/.]+$/, '');
        firstFileName = firstFileName.replace(/[^a-zA-Z0-9_-]/g, '_');
    }
    
    const filename = `compare_${firstFileName}_${date}.xls`;
    
    const blob = new Blob([htmlContent], {
        type: 'application/vnd.ms-excel;charset=utf-8;'
    });
    
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    showExportSuccess(filename);
}

function prepareDataForExport() {
    return prepareDataFromRawData();
}

function prepareDataFromRenderedTable() {
    const tableBody = document.querySelector('.diff-table-body tbody');
    const tableHeader = document.querySelector('.diff-table-header thead');
    
    if (!tableBody || !tableHeader) {
        return null;
    }
    
    const data = [];
    const formatting = {};
    const colWidths = [];
    
    const headerRow = tableHeader.querySelector('tr');
    if (headerRow) {
        const headers = Array.from(headerRow.querySelectorAll('th')).map(th => String(th.textContent || '')); // Ensure string format
        data.push(headers);
        
        headers.forEach(() => colWidths.push({ wch: 15 }));
        
        headers.forEach((header, col) => {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            formatting[cellAddress] = {
                fill: { fgColor: { rgb: "f8f9fa" } },
                font: { bold: true, color: { rgb: "212529" } },
                border: {
                    top: { style: "thin", color: { rgb: "D4D4D4" } },
                    bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                    left: { style: "thin", color: { rgb: "D4D4D4" } },
                    right: { style: "thin", color: { rgb: "D4D4D4" } }
                }
            };
        });
    }
    
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    
    rows.forEach((row, rowIndex) => {
        const cells = Array.from(row.querySelectorAll('td'));
        const rowData = cells.map(cell => {
            const cellText = cell.textContent || '';
            return String(cellText);
        });
        data.push(rowData);
        
        cells.forEach((cell, colIndex) => {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: colIndex });
            let bgColor = "ffffff";
            
            if (cell.classList.contains('warn-cell')) {
                bgColor = "f8d7da";
            } else if (cell.classList.contains('tolerance-cell')) {
                bgColor = "ffeaa7";
            } else if (cell.classList.contains('identical')) {
                bgColor = "d4edda";
            } else if (cell.classList.contains('new-cell1')) {
                bgColor = "65add7";
            } else if (cell.classList.contains('new-cell2')) {
                bgColor = "63cfbf";
            } else if (cell.classList.contains('new-cell')) {
                bgColor = "65add7";
            } else if (cell.classList.contains('unique-cell')) {
                if (cell.classList.contains('only-in-file1') || cell.classList.contains('new-cell1')) {
                    bgColor = "65add7";
                } else if (cell.classList.contains('only-in-file2') || cell.classList.contains('new-cell2')) {
                    bgColor = "63cfbf";
                } else {
                    bgColor = "cdedff";
                }
            } else if (cell.classList.contains('only-in-file1')) {
                bgColor = "65add7";
            } else if (cell.classList.contains('only-in-file2')) {
                bgColor = "63cfbf";
            }
            
            if (bgColor === "ffffff") {
                const computedStyle = window.getComputedStyle(cell);
                const bgColorStyle = computedStyle.backgroundColor;
                
                if (bgColorStyle.includes('255, 234, 167') || bgColorStyle.includes('ffeaa7')) {
                    bgColor = "ffeaa7";
                } else if (bgColorStyle.includes('248, 215, 218') || bgColorStyle.includes('f8d7da')) {
                    bgColor = "f8d7da"; 
                } else if (bgColorStyle.includes('212, 237, 218') || bgColorStyle.includes('d4edda')) {
                    bgColor = "d4edda";
                } else if (bgColorStyle.includes('99, 207, 191') || bgColorStyle.includes('63cfbf')) {
                    bgColor = "63cfbf";
                } else if (bgColorStyle.includes('101, 173, 215') || bgColorStyle.includes('65add7')) {
                    bgColor = "65add7";
                }
            }
            
            formatting[cellAddress] = {
                fill: { fgColor: { rgb: bgColor } },
                font: { color: { rgb: "212529" } },
                border: {
                    top: { style: "thin", color: { rgb: "D4D4D4" } },
                    bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                    left: { style: "thin", color: { rgb: "D4D4D4" } },
                    right: { style: "thin", color: { rgb: "D4D4D4" } }
                }
            };
        });
    });
    
    return { data, formatting, colWidths };
}

function prepareDataFromRawData() {
    // Use full comparison data for export if available, otherwise use currentPairs
    const pairsForExport = window.fullComparisonPairs || currentPairs;
    
    // If no data available, return null
    if (!pairsForExport || pairsForExport.length === 0) {
        console.log('No pairs data available for export');
        return null;
    }
    
    const data = []; const formatting = []; const colWidths = [];
    const hideSame = document.getElementById('hideSameRows')?.checked || false;
    const hideDiffRows = document.getElementById('hideDiffColumns')?.checked || false;
    const hideNewRows1 = document.getElementById('hideNewRows1')?.checked || false;
    const hideNewRows2 = document.getElementById('hideNewRows2')?.checked || false;
    const headers = ['Source'];
    for (let c = 0; c < currentFinalAllCols; c++) headers.push(String(currentFinalHeaders[c] || ''));
    data.push(headers); colWidths.push({ wch: 20 }); for (let c = 0; c < currentFinalAllCols; c++) colWidths.push({ wch: 15 });
    headers.forEach((_, col) => { const addr = XLSX.utils.encode_cell({ r:0, c:col }); formatting[addr] = { fill:{ fgColor:{ rgb:'f8f9fa'}}, font:{ bold:true, color:{rgb:'212529'}}, border:{ top:{style:'thin',color:{rgb:'D4D4D4'}}, bottom:{style:'thin',color:{rgb:'D4D4D4'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } }; });
    let rowIndex = 1;
    
    pairsForExport.forEach(pair => {
        const row1 = pair.row1, row2 = pair.row2;
        const row1Upper = row1 ? row1.map(v => v !== undefined ? v.toString().toUpperCase() : '') : null;
        const row2Upper = row2 ? row2.map(v => v !== undefined ? v.toString().toUpperCase() : '') : null;
        let isEmpty = true, hasDiff = false, allSame = true, hasTolerance = false;
        
        let columnComparisons = [];
        
        for (let c = 0; c < currentFinalAllCols; c++) {
            const v1 = row1 ? (row1[c] ?? '') : ''; 
            const v2 = row2 ? (row2[c] ?? '') : '';
            if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) isEmpty = false;
            
            if (row1 && row2) {
                if (toleranceMode && typeof compareValuesWithTolerance === 'function') {
                    const comparisonResult = compareValuesWithTolerance(v1, v2);
                    columnComparisons[c] = comparisonResult;
                    
                    if (comparisonResult === 'different') {
                        hasDiff = true;
                        allSame = false;
                    } else if (comparisonResult === 'tolerance') {
                        hasTolerance = true;
                        allSame = false;
                    }
                } else {
                    if (row1Upper[c] !== row2Upper[c]) {
                        hasDiff = true;
                        allSame = false;
                        columnComparisons[c] = 'different';
                    } else {
                        columnComparisons[c] = 'identical';
                    }
                }
            } else { 
                allSame = false; 
                columnComparisons[c] = 'different';
            }
        }
        if (isEmpty) return;
        if (hideSame && row1 && row2 && allSame) return;
        if (hideNewRows1 && row1 && !row2) return;
        if (hideNewRows2 && !row1 && row2) return;
        if (hideDiffRows && row1 && row2 && (hasDiff || hasTolerance)) return;
        
        if (row1 && row2 && (hasDiff || hasTolerance)) {
            const dataRow1 = [getFileDisplayName(fileName1, sheetName1) || 'File 1'];
            for (let c = 0; c < currentFinalAllCols; c++) {
                const v1 = row1[c] ?? ''; const v2 = row2[c] ?? ''; dataRow1.push(String(v1));
                const addr = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                const compResult = columnComparisons[c];
                
                if (compResult === 'different') {
                    formatting[addr] = { fill:{ fgColor:{ rgb:'f8d7da'}}, font:{ color:{rgb:'212529'}}, border:{ top:{style:'thick',color:{rgb:'000000'}}, bottom:{style:'thin',color:{rgb:'D4D4D4'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
                } else if (compResult === 'tolerance') {
                    formatting[addr] = { fill:{ fgColor:{ rgb:'ffeaa7'}}, font:{ color:{rgb:'212529'}}, border:{ top:{style:'thick',color:{rgb:'000000'}}, bottom:{style:'thin',color:{rgb:'D4D4D4'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
                } else {
                    formatting[addr] = { border:{ top:{style:'thick',color:{rgb:'000000'}}, bottom:{style:'thin',color:{rgb:'D4D4D4'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
                }
            }
            data.push(dataRow1);
            const source1 = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
            formatting[source1] = { fill:{ fgColor:{ rgb:'f8d7da'}}, font:{ color:{rgb:'212529'}, bold:true}, border:{ top:{style:'thick',color:{rgb:'000000'}}, bottom:{style:'thin',color:{rgb:'D4D4D4'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
            rowIndex++;
            const dataRow2 = [getFileDisplayName(fileName2, sheetName2) || 'File 2'];
            for (let c = 0; c < currentFinalAllCols; c++) {
                const v2 = row2[c] ?? ''; dataRow2.push(String(v2));
                const addr = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                const compResult = columnComparisons[c];
                
                if (compResult === 'different') {
                    formatting[addr] = { fill:{ fgColor:{ rgb:'f8d7da'}}, font:{ color:{rgb:'212529'}}, border:{ top:{style:'thin',color:{rgb:'D4D4D4'}}, bottom:{style:'thick',color:{rgb:'000000'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
                } else if (compResult === 'tolerance') {
                    formatting[addr] = { fill:{ fgColor:{ rgb:'ffeaa7'}}, font:{ color:{rgb:'212529'}}, border:{ top:{style:'thin',color:{rgb:'D4D4D4'}}, bottom:{style:'thick',color:{rgb:'000000'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
                } else {
                    formatting[addr] = { border:{ top:{style:'thin',color:{rgb:'D4D4D4'}}, bottom:{style:'thick',color:{rgb:'000000'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
                }
            }
            data.push(dataRow2);
            const source2 = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
            formatting[source2] = { fill:{ fgColor:{ rgb:'f8d7da'}}, font:{ color:{rgb:'212529'}, bold:true}, border:{ top:{style:'thin',color:{rgb:'D4D4D4'}}, bottom:{style:'thick',color:{rgb:'000000'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
            rowIndex++;
        } else {
            let source = '', fillColor = '';
            if (row1 && row2 && allSame) { source = 'Both files'; fillColor = 'd4edda'; }
            else if (row1 && !row2) { source = getFileDisplayName(fileName1, sheetName1) || 'File 1'; fillColor = '65add7'; }
            else if (!row1 && row2) { source = getFileDisplayName(fileName2, sheetName2) || 'File 2'; fillColor = '63cfbf'; }
            const dataRow = [source];
            for (let c = 0; c < currentFinalAllCols; c++) {
                let v = ''; if (row1 && !row2) v = row1[c] ?? ''; else if (!row1 && row2) v = row2[c] ?? ''; else if (row1 && row2) v = row1[c] ?? '';
                dataRow.push(String(v));
                const addr = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                formatting[addr] = { fill:{ fgColor:{ rgb:fillColor } }, font:{ color:{rgb:'212529'}}, border:{ top:{style:'thin',color:{rgb:'D4D4D4'}}, bottom:{style:'thin',color:{rgb:'D4D4D4'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
            }
            data.push(dataRow);
            const sourceAddr = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
            formatting[sourceAddr] = { fill:{ fgColor:{ rgb:fillColor } }, font:{ color:{rgb:'212529'}, bold:true}, border:{ top:{style:'thin',color:{rgb:'D4D4D4'}}, bottom:{style:'thin',color:{rgb:'D4D4D4'}}, left:{style:'thin',color:{rgb:'D4D4D4'}}, right:{style:'thin',color:{rgb:'D4D4D4'}} } };
            rowIndex++;
        }
    });
    return { data, formatting, colWidths };
}

function showExportSuccess() {
    const successDiv = document.createElement('div');
    successDiv.style.cssText = 'position:fixed;top:20px;right:20px;background:#d4edda;color:#155724;padding:15px 25px;border:1px solid #c3e6cb;border-radius:8px;box-shadow:0 4px 15px rgba(0,0,0,0.15);z-index:10000;font-size:14px;font-weight:500;max-width:350px;line-height:1.4;';
    successDiv.innerHTML = '✅ <strong>Excel exported!</strong>';
    document.body.appendChild(successDiv);
    setTimeout(() => { successDiv.parentNode && successDiv.parentNode.removeChild(successDiv); }, 5000);
}

function setColumnWidths(ws, colWidths) { ws['!cols'] = colWidths; }

async function performFullComparisonForExport() {
    return new Promise((resolve, reject) => {
        try {
            // Get the original data that was used for quick comparison
            const body1 = window.originalBody1 || [];
            const body2 = window.originalBody2 || [];
            const finalHeaders = window.currentFinalHeaders || [];
            const finalAllCols = window.currentFinalAllCols || 0;
            
            if (!body1.length || !body2.length) {
                reject(new Error('Original comparison data not available'));
                return;
            }
            
            // Clear quick mode flag during processing
            const quickModeInfoDiv = document.getElementById('quick-mode-info');
            if (quickModeInfoDiv) {
                quickModeInfoDiv.innerHTML = `
                    <strong>⚡ Full Comparison in Progress:</strong> Processing all <strong>${Math.max(body1.length, body2.length).toLocaleString()}</strong> rows for complete analysis...
                    <br><div style="margin-top: 5px; background-color: #e9ecef; height: 4px; border-radius: 2px;"><div id="full-comparison-progress" style="background-color: #28a745; height: 100%; width: 0%; border-radius: 2px; transition: width 0.3s;"></div></div>
                `;
            }
            
            // Perform full comparison without quickModeOnly flag
            performFuzzyMatchingForExportInternal(body1, body2, finalHeaders, finalAllCols, true, {}, false, (fullPairs) => {
                // Store full comparison results
                window.fullComparisonPairs = fullPairs;
                currentPairs = fullPairs; // Update current pairs for export
                
                // Update the info message
                if (quickModeInfoDiv) {
                    quickModeInfoDiv.innerHTML = `
                        <strong>✅ Full Comparison Complete:</strong> Analyzed all <strong>${Math.max(body1.length, body2.length).toLocaleString()}</strong> rows. 
                        Ready for complete Excel export with all comparison data.
                    `;
                }
                
                resolve();
            });
            
        } catch (error) {
            reject(error);
        }
    });
}

// Internal version of the comparison function for export use
function performFuzzyMatchingForExportInternal(body1, body2, finalHeaders, finalAllCols, isLargeFile, tableHeaders, quickModeOnly, callback) {
    // Store original data for later use
    window.originalBody1 = body1;
    window.originalBody2 = body2;
    
    const combinedData = [finalHeaders, ...body1, ...body2];
    const keyColumnIndexes = smartDetectKeyColumns ? smartDetectKeyColumns(finalHeaders, combinedData) : [];
    
    let used2 = new Array(body2.length).fill(false);
    let pairs = [];
    let processedRows = 0;
    
    function countMatches(rowA, rowB) {
        let matches = 0;
        let keyMatches = 0;
        
        for (let i = 0; i < finalAllCols; i++) {
            let valueA = (rowA[i] || '').toString();
            let valueB = (rowB[i] || '').toString();
            
            if (valueA.toUpperCase() === valueB.toUpperCase()) {
                matches++;
                if (keyColumnIndexes.includes(i)) {
                    keyMatches++;
                }
            }
        }
        
        return (keyMatches * 3) + (matches - keyMatches);
    }
    
    function processBatchInternal(startIndex, batchSize) {
        const endIndex = Math.min(startIndex + batchSize, body1.length);
        
        for (let i = startIndex; i < endIndex; i++) {
            let bestIdx = -1, bestScore = -1;
            for (let j = 0; j < body2.length; j++) {
                if (used2[j]) continue;
                let score = countMatches(body1[i], body2[j]);
                if (score > bestScore) {
                    bestScore = score;
                    bestIdx = j;
                }
            }
            
            const minTotalMatches = Math.ceil(finalAllCols * 0.5);
            
            if (bestScore >= minTotalMatches) {
                pairs.push({row1: body1[i], row2: body2[bestIdx]});
                used2[bestIdx] = true;
            } else {
                pairs.push({row1: body1[i], row2: null});
            }
            
            processedRows++;
        }
        
        // Update progress
        const progress = Math.round((endIndex / body1.length) * 100);
        const progressBar = document.getElementById('full-comparison-progress');
        if (progressBar) {
            progressBar.style.width = progress + '%';
        }
        
        if (endIndex < body1.length) {
            setTimeout(() => processBatchInternal(endIndex, batchSize), 10);
        } else {
            // Add remaining rows from body2
            for (let j = 0; j < body2.length; j++) {
                if (!used2[j]) {
                    pairs.push({row1: null, row2: body2[j]});
                }
            }
            
            callback(pairs);
        }
    }
    
    const batchSize = body1.length > 5000 ? 100 : body1.length > 1000 ? 250 : 1000;
    processBatchInternal(0, batchSize);
}