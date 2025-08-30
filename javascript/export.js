function exportToExcel() {
    if (!currentPairs || currentPairs.length === 0) {
        alert('No data to export. Please compare files first.');
        return;
    }
    
    try {
        const exportData = prepareDataForExport();
        const htmlContent = createStyledHTMLTable(exportData);
        downloadExcelFromHTML(htmlContent);
    } catch (error) {
        console.error('Export error:', error);
        alert('Error exporting to Excel: ' + error.message);
    }
}

function createStyledHTMLTable(exportData) {
    let html = `
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
    `;
    
    exportData.data.forEach((row, rowIndex) => {
        if (rowIndex === 0) {
            html += '<tr style="font-weight: bold;">';
        } else {
            html += '<tr>';
        }
        
        row.forEach((cellValue, colIndex) => {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            const formatting = exportData.formatting[cellAddress];
            
            let style = 'border: 0.5pt solid #D4D4D4; padding: 4px; vertical-align: top;';
            
            if (formatting && formatting.border) {
                if (formatting.border.top && formatting.border.top.style === 'thick') {
                    style = 'border-top: 2pt solid #000; border-bottom: 0.5pt solid #D4D4D4; border-left: 0.5pt solid #D4D4D4; border-right: 0.5pt solid #D4D4D4; padding:4px; vertical-align:top;';
                } else if (formatting.border.bottom && formatting.border.bottom.style === 'thick') {
                    style = 'border-top: 0.5pt solid #D4D4D4; border-bottom: 2pt solid #000; border-left: 0.5pt solid #D4D4D4; border-right: 0.5pt solid #D4D4D4; padding:4px; vertical-align:top;';
                }
            }
            if (formatting && formatting.fill && formatting.fill.fgColor) {
                let bgColor = formatting.fill.fgColor.rgb;
                switch (bgColor) {
                    case 'FF6B6B': bgColor = 'f8d7da'; break;
                    case '4CAF50': bgColor = 'd4edda'; break;
                    case 'ffeaa7': bgColor = 'ffeaa7'; break;
                }
                style += ` background-color:#${bgColor};`;
            }
            if (formatting && formatting.font) {
                if (formatting.font.bold) style += ' font-weight:bold;';
                style += ' color:#212529;';
            }
            const safeValue = String(cellValue || '').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
            if (rowIndex === 0) {
                html += `<td style="${style}" class="thin-border" x:autofilter="all">${safeValue}</td>`;
            } else {
                const textValue = ' \t' + safeValue;
                html += `<td style="${style}mso-number-format:'@';mso-data-type:string;white-space:pre;" class="thin-border text-format" x:str="${safeValue}" x:format="@">${textValue}</td>`;
            }
        });
        html += '</tr>';
    });
    
    html += `
        </table>
    </body>
    </html>
    `;
    
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
    currentPairs.forEach(pair => {
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

function setColumnWidths(ws, colWidths) { ws['!cols'] = colWidths; }

function showExportSuccess() {
    const successDiv = document.createElement('div');
    successDiv.style.cssText = 'position:fixed;top:20px;right:20px;background:#d4edda;color:#155724;padding:15px 25px;border:1px solid #c3e6cb;border-radius:8px;box-shadow:0 4px 15px rgba(0,0,0,0.15);z-index:10000;font-size:14px;font-weight:500;max-width:350px;line-height:1.4;';
    successDiv.innerHTML = '✅ <strong>Excel exported!</strong>';
    document.body.appendChild(successDiv);
    setTimeout(() => { successDiv.parentNode && successDiv.parentNode.removeChild(successDiv); }, 5000);
}