// Function to export comparison table to Excel with formatting
function exportToExcel() {
    if (!currentPairs || currentPairs.length === 0) {
        alert('No data to export. Please compare files first.');
        return;
    }
    
    try {
        // Prepare data for export
        const exportData = prepareDataForExport();
        
        // Create HTML table with styling
        let htmlContent = createStyledHTMLTable(exportData);
        
        // Create and download Excel file using HTML
        downloadExcelFromHTML(htmlContent);
        
    } catch (error) {
        console.error('Export error:', error);
        alert('Error exporting to Excel: ' + error.message);
    }
}

// Create styled HTML table for Excel export
function createStyledHTMLTable(exportData) {
    let html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
        <meta charset="utf-8">
        <style>
            .thick-top { border-top: 2pt solid black !important; }
            .thick-bottom { border-bottom: 2pt solid black !important; }
            .thin-border { border: 0.5pt solid #D4D4D4; }
        </style>
        <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>Comparison Results</x:Name><x:WorksheetOptions><x:DisplayGridlines/><x:Selected/><x:Panes><x:Pane><x:Number>0</x:Number><x:ActiveRow>0</x:ActiveRow><x:ActiveCol>0</x:ActiveCol></x:Pane></x:Panes></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
    </head>
    <body>
        <table border="0" style="border-collapse: collapse;" x:publishsource="Excel">
    `;
    
    // Add table rows with data and formatting
    exportData.data.forEach((row, rowIndex) => {
        // Add autofilter to header row
        if (rowIndex === 0) {
            html += '<tr style="font-weight: bold;">';
        } else {
            html += '<tr>';
        }
        
        row.forEach((cellValue, colIndex) => {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            const formatting = exportData.formatting[cellAddress];
            
            let style = 'border: 0.5pt solid #D4D4D4; padding: 4px; vertical-align: top;';
            
            // Check for thick borders from formatting
            if (formatting && formatting.border) {
                let borderStyle = 'border: ';
                if (formatting.border.top && formatting.border.top.style === 'thick') {
                    borderStyle = 'border-top: 2pt solid #000000; border-bottom: 0.5pt solid #D4D4D4; border-left: 0.5pt solid #D4D4D4; border-right: 0.5pt solid #D4D4D4; ';
                } else if (formatting.border.bottom && formatting.border.bottom.style === 'thick') {
                    borderStyle = 'border-top: 0.5pt solid #D4D4D4; border-bottom: 2pt solid #000000; border-left: 0.5pt solid #D4D4D4; border-right: 0.5pt solid #D4D4D4; ';
                } else {
                    borderStyle = 'border: 0.5pt solid #D4D4D4; ';
                }
                style = borderStyle + 'padding: 4px; vertical-align: top;';
            }
            
            if (formatting) {
                if (formatting.fill && formatting.fill.fgColor) {
                    // Use the same colors as in the comparison table
                    let bgColor = formatting.fill.fgColor.rgb;
                    
                    // Map bright colors to muted ones from CSS
                    switch(bgColor) {
                        case 'FF6B6B': // Red differences -> muted red
                            bgColor = 'f8d7da';
                            break;
                        case '4CAF50': // Green identical -> muted green
                            bgColor = 'd4edda';
                            break;
                        case 'd4edda': // Already correct muted green for identical
                            bgColor = 'd4edda';
                            break;
                        case '2196F3': // Blue file 1 -> muted yellow
                            bgColor = 'fff3cd';
                            break;
                        case 'FFC107': // Yellow file 2 -> muted yellow
                            bgColor = 'fff3cd';
                            break;
                        case 'fff3cd': // Already correct muted yellow for unique
                            bgColor = 'fff3cd';
                            break;
                        case 'f8d7da': // Already correct muted red for differences
                            bgColor = 'f8d7da';
                            break;
                        case 'f8f9fa': // Already correct muted gray for headers
                            bgColor = 'f8f9fa';
                            break;
                        case 'D3D3D3': // Gray header
                            bgColor = 'f8f9fa';
                            break;
                    }
                    style += ` background-color: #${bgColor};`;
                }
                if (formatting.font) {
                    if (formatting.font.color && formatting.fill && formatting.fill.fgColor.rgb !== 'D3D3D3') {
                        // For colored cells, use dark text instead of white for better readability
                        style += ' color: #212529;';
                    } else if (formatting.font.color) {
                        style += ` color: #${formatting.font.color.rgb};`;
                    }
                    if (formatting.font.bold) {
                        style += ' font-weight: bold;';
                    }
                }
            }
            
            // Add CSS classes for thick borders
            let cssClass = '';
            if (formatting && formatting.border) {
                if (formatting.border.top && formatting.border.top.style === 'thick') {
                    cssClass += ' thick-top';
                }
                if (formatting.border.bottom && formatting.border.bottom.style === 'thick') {
                    cssClass += ' thick-bottom';
                }
            }
            
            // Add autofilter attribute to header cells
            if (rowIndex === 0) {
                html += `<td style="${style}" class="thin-border${cssClass}" x:autofilter="all">${cellValue || ''}</td>`;
            } else {
                html += `<td style="${style}" class="thin-border${cssClass}">${cellValue || ''}</td>`;
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

// Download Excel file from HTML
function downloadExcelFromHTML(htmlContent) {
    const now = new Date();
    const date = now.toISOString().slice(0, 10); // YYYY-MM-DD format
    
    // Get the first file name without extension
    let firstFileName = 'file1';
    if (typeof fileName1 !== 'undefined' && fileName1) {
        firstFileName = fileName1.replace(/\.[^/.]+$/, ''); // Remove file extension
        firstFileName = firstFileName.replace(/[^a-zA-Z0-9_-]/g, '_'); // Replace special chars with underscore
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

// Function to prepare data for Excel export
function prepareDataForExport() {
    const data = [];
    const formatting = {};
    const colWidths = [];
    
    // Get filter states
    const hideSameEl = document.getElementById('hideSameRows');
    const hideDiffEl = document.getElementById('hideDiffColumns');
    const hideNewRows1El = document.getElementById('hideNewRows1');
    const hideNewRows2El = document.getElementById('hideNewRows2');
    
    const hideSame = hideSameEl ? hideSameEl.checked : false;
    const hideDiffRows = hideDiffEl ? hideDiffEl.checked : false;
    const hideNewRows1 = hideNewRows1El ? hideNewRows1El.checked : false;
    const hideNewRows2 = hideNewRows2El ? hideNewRows2El.checked : false;
    
    // Add headers
    const headers = ['Source'];
    for (let c = 0; c < currentFinalAllCols; c++) {
        headers.push(currentFinalHeaders[c] || '');
    }
    data.push(headers);
    
    // Set column widths (Source column wider, data columns standard)
    colWidths.push({ wch: 20 }); // Source column
    for (let c = 0; c < currentFinalAllCols; c++) {
        colWidths.push({ wch: 15 }); // Data columns
    }
    
    // Add header formatting
    for (let col = 0; col < headers.length; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
        formatting[cellAddress] = {
            fill: { fgColor: { rgb: "f8f9fa" } }, // Light gray background (same as CSS)
            font: { bold: true, color: { rgb: "212529" } }, // Dark bold text
            border: {
                top: { style: "thin", color: { rgb: "D4D4D4" } },
                bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                left: { style: "thin", color: { rgb: "D4D4D4" } },
                right: { style: "thin", color: { rgb: "D4D4D4" } }
            }
        };
    }
    
    let rowIndex = 1; // Start from row 1 (0 is headers)
    
    // Process comparison pairs
    currentPairs.forEach(pair => {
        const row1 = pair.row1;
        const row2 = pair.row2;
        
        // Pre-convert values to uppercase for comparison
        const row1Upper = row1 ? row1.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        const row2Upper = row2 ? row2.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        
        let isEmpty = true;
        let hasWarn = false;
        let allSame = true;
        
        // Check if row should be filtered out
        for (let c = 0; c < currentFinalAllCols; c++) {
            const v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
            const v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
            if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                isEmpty = false;
            }
            if (row1 && row2) {
                if (row1Upper[c] !== row2Upper[c]) {
                    hasWarn = true;
                    allSame = false;
                }
            } else {
                allSame = false;
            }
        }
        
        if (isEmpty) return;
        if (hideSame && row1 && row2 && allSame) return;
        if (hideNewRows1 && row1 && !row2) return;
        if (hideNewRows2 && !row1 && row2) return;
        if (hideDiffRows && row1 && row2 && hasWarn) return;
        
        // Add rows to export data
        if (row1 && row2 && hasWarn) {
            // Different data - create two rows
            // Row 1 data
            const dataRow1 = [fileName1 || 'File 1'];
            for (let c = 0; c < currentFinalAllCols; c++) {
                const v1 = row1[c] !== undefined ? row1[c] : '';
                const v2 = row2[c] !== undefined ? row2[c] : '';
                const v1Upper = row1Upper[c] || '';
                const v2Upper = row2Upper[c] || '';
                
                dataRow1.push(v1);
                
                // Apply formatting for different cells
                if (v1Upper !== v2Upper) {
                    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                    formatting[cellAddress] = {
                        fill: { fgColor: { rgb: "f8d7da" } }, // Muted red background for differences (same as CSS)
                        font: { color: { rgb: "212529" } }, // Dark text
                        border: {
                            top: { style: "thick", color: { rgb: "000000" } }, // Thick top border for comparison group
                            bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                            left: { style: "thin", color: { rgb: "D4D4D4" } },
                            right: { style: "thin", color: { rgb: "D4D4D4" } }
                        }
                    };
                } else {
                    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                    formatting[cellAddress] = {
                        border: {
                            top: { style: "thick", color: { rgb: "000000" } }, // Thick top border for comparison group
                            bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                            left: { style: "thin", color: { rgb: "D4D4D4" } },
                            right: { style: "thin", color: { rgb: "D4D4D4" } }
                        }
                    };
                }
            }
            data.push(dataRow1);
            
            // Source cell formatting for File 1
            const sourceCell1 = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
            formatting[sourceCell1] = {
                fill: { fgColor: { rgb: "f8d7da" } }, // Muted red background
                font: { color: { rgb: "212529" }, bold: true }, // Dark bold text
                border: {
                    top: { style: "thick", color: { rgb: "000000" } }, // Thick top border for comparison group
                    bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                    left: { style: "thin", color: { rgb: "D4D4D4" } },
                    right: { style: "thin", color: { rgb: "D4D4D4" } }
                }
            };
            rowIndex++;
            
            // Row 2 data
            const dataRow2 = [fileName2 || 'File 2'];
            for (let c = 0; c < currentFinalAllCols; c++) {
                const v1 = row1[c] !== undefined ? row1[c] : '';
                const v2 = row2[c] !== undefined ? row2[c] : '';
                const v1Upper = row1Upper[c] || '';
                const v2Upper = row2Upper[c] || '';
                
                dataRow2.push(v2);
                
                // Apply formatting for different cells
                if (v1Upper !== v2Upper) {
                    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                    formatting[cellAddress] = {
                        fill: { fgColor: { rgb: "f8d7da" } }, // Muted red background for differences
                        font: { color: { rgb: "212529" } }, // Dark text
                        border: {
                            top: { style: "thin", color: { rgb: "D4D4D4" } },
                            bottom: { style: "thick", color: { rgb: "000000" } }, // Thick bottom border for comparison group
                            left: { style: "thin", color: { rgb: "D4D4D4" } },
                            right: { style: "thin", color: { rgb: "D4D4D4" } }
                        }
                    };
                } else {
                    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                    formatting[cellAddress] = {
                        border: {
                            top: { style: "thin", color: { rgb: "D4D4D4" } },
                            bottom: { style: "thick", color: { rgb: "000000" } }, // Thick bottom border for comparison group
                            left: { style: "thin", color: { rgb: "D4D4D4" } },
                            right: { style: "thin", color: { rgb: "D4D4D4" } }
                        }
                    };
                }
            }
            data.push(dataRow2);
            
            // Source cell formatting for File 2
            const sourceCell2 = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
            formatting[sourceCell2] = {
                fill: { fgColor: { rgb: "f8d7da" } }, // Muted red background
                font: { color: { rgb: "212529" }, bold: true }, // Dark bold text
                border: {
                    top: { style: "thin", color: { rgb: "D4D4D4" } },
                    bottom: { style: "thick", color: { rgb: "000000" } }, // Thick bottom border for comparison group
                    left: { style: "thin", color: { rgb: "D4D4D4" } },
                    right: { style: "thin", color: { rgb: "D4D4D4" } }
                }
            };
            rowIndex++;
            
        } else {
            // Single row for identical data or data from only one file
            let source = '';
            let fillColor = '';
            
            if (row1 && row2 && allSame) {
                source = 'Both files';
                fillColor = "d4edda"; // Muted green for identical (same as CSS)
            } else if (row1 && !row2) {
                source = fileName1 || 'File 1';
                fillColor = "fff3cd"; // Muted yellow for new in file 1 (same as CSS)
            } else if (!row1 && row2) {
                source = fileName2 || 'File 2';
                fillColor = "fff3cd"; // Muted yellow for new in file 2 (same as CSS)
            }
            
            const dataRow = [source];
            for (let c = 0; c < currentFinalAllCols; c++) {
                let cellValue = '';
                if (row1 && !row2) {
                    cellValue = row1[c] !== undefined ? row1[c] : '';
                } else if (!row1 && row2) {
                    cellValue = row2[c] !== undefined ? row2[c] : '';
                } else if (row1 && row2) {
                    cellValue = row1[c] !== undefined ? row1[c] : '';
                }
                dataRow.push(cellValue);
                
                // Apply cell formatting
                const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: c + 1 });
                formatting[cellAddress] = {
                    fill: { fgColor: { rgb: fillColor } },
                    font: { color: { rgb: "212529" } }, // Dark text for better readability
                    border: {
                        top: { style: "thin", color: { rgb: "D4D4D4" } },
                        bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                        left: { style: "thin", color: { rgb: "D4D4D4" } },
                        right: { style: "thin", color: { rgb: "D4D4D4" } }
                    }
                };
            }
            data.push(dataRow);
            
            // Source cell formatting
            const sourceCell = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
            formatting[sourceCell] = {
                fill: { fgColor: { rgb: fillColor } },
                font: { color: { rgb: "212529" }, bold: true }, // Dark bold text
                border: {
                    top: { style: "thin", color: { rgb: "D4D4D4" } },
                    bottom: { style: "thin", color: { rgb: "D4D4D4" } },
                    left: { style: "thin", color: { rgb: "D4D4D4" } },
                    right: { style: "thin", color: { rgb: "D4D4D4" } }
                }
            };
            rowIndex++;
        }
    });
    
    return { data, formatting, colWidths };
}

// Function to set column widths
function setColumnWidths(ws, colWidths) {
    ws['!cols'] = colWidths;
}

// Function to show export success message
function showExportSuccess(filename) {
    // Create a temporary success message
    const successDiv = document.createElement('div');
    successDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #d4edda;
        color: #155724;
        padding: 12px 20px;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        z-index: 10000;
        font-size: 14px;
        font-weight: 500;
    `;
    successDiv.innerHTML = `✅ Excel file exported successfully: ${filename}`;
    
    document.body.appendChild(successDiv);
    
    // Remove message after 3 seconds
    setTimeout(() => {
        if (successDiv && successDiv.parentNode) {
            successDiv.parentNode.removeChild(successDiv);
        }
    }, 3000);
}
