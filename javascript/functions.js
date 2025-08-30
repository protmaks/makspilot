// Excel and CSV File Comparison Functions

// Configuration constants
const MAX_ROWS_LIMIT = 55000; // Maximum allowed rows per file
const MAX_COLS_LIMIT = 120; // Maximum allowed columns per file
const DETAILED_TABLE_LIMIT = 15000; // Limit for showing detailed comparison table (above this only summary and export)

let data1 = [], data2 = [];
let fileName1 = '', fileName2 = '';
let sheetName1 = '', sheetName2 = ''; // Store current sheet names
let workbook1 = null, workbook2 = null; // Store workbooks for sheet selection
let toleranceMode = false; // Global variable for tolerance comparison mode

// Global variables for sorting
let currentSortColumn = -1;
let currentSortDirection = 'asc';
let currentPairs = [];
let currentFinalHeaders = [];
let currentFinalAllCols = 0;
let currentDiffColumns1 = '-';
let currentDiffColumns2 = '-';

// Function to get translated summary table headers
function getSummaryTableHeaders() {
    // Detect current language based on page URL
    const currentLang = window.location.pathname.includes('/ru/') ? 'ru' : 
                       window.location.pathname.includes('/pl/') ? 'pl' :
                       window.location.pathname.includes('/es/') ? 'es' :
                       window.location.pathname.includes('/de/') ? 'de' :
                       window.location.pathname.includes('/ja/') ? 'ja' :
                       window.location.pathname.includes('/pt/') ? 'pt' :
                       window.location.pathname.includes('/zh/') ? 'zh' :
                       window.location.pathname.includes('/ar/') ? 'ar' : 'en';
                       
    const translations = {
        'ru': {
            file: 'Файл',
            rowCount: 'Количество строк',
            rowsOnlyInFile: 'Строки только в этом файле',
            identicalRows: 'Идентичные строки',
            similarity: '% Сходство',
            diffColumns: 'Различающиеся колонки',
            file1: 'Файл 1',
            file2: 'Файл 2',
            calculating: 'Вычисляется...'
        },
        'pl': {
            file: 'Plik',
            rowCount: 'Liczba wierszy',
            rowsOnlyInFile: 'Wiersze tylko w tym pliku',
            identicalRows: 'Identyczne wiersze',
            similarity: '% Podobieństwo',
            diffColumns: 'Różne kolumny',
            file1: 'Plik 1',
            file2: 'Plik 2',
            calculating: 'Obliczanie...'
        },
        'es': {
            file: 'Archivo',
            rowCount: 'Número de filas',
            rowsOnlyInFile: 'Filas solo en este archivo',
            identicalRows: 'Filas idénticas',
            similarity: '% Similitud',
            diffColumns: 'Columnas diferentes',
            file1: 'Archivo 1',
            file2: 'Archivo 2',
            calculating: 'Calculando...'
        },
        'de': {
            file: 'Datei',
            rowCount: 'Anzahl Zeilen',
            rowsOnlyInFile: 'Zeilen nur in dieser Datei',
            identicalRows: 'Identische Zeilen',
            similarity: '% Ähnlichkeit',
            diffColumns: 'Unterschiedliche Spalten',
            file1: 'Datei 1',
            file2: 'Datei 2',
            calculating: 'Berechnung...'
        },
        'ja': {
            file: 'ファイル',
            rowCount: '行数',
            rowsOnlyInFile: 'このファイルのみの行',
            identicalRows: '同一行',
            similarity: '% 類似度',
            diffColumns: '異なる列',
            file1: 'ファイル 1',
            file2: 'ファイル 2',
            calculating: '計算中...'
        },
        'pt': {
            file: 'Arquivo',
            rowCount: 'Número de linhas',
            rowsOnlyInFile: 'Linhas apenas neste arquivo',
            identicalRows: 'Linhas idênticas',
            similarity: '% Similaridade',
            diffColumns: 'Colunas diferentes',
            file1: 'Arquivo 1',
            file2: 'Arquivo 2',
            calculating: 'Calculando...'
        },
        'zh': {
            file: '文件',
            rowCount: '行数',
            rowsOnlyInFile: '仅此文件中的行',
            identicalRows: '相同行',
            similarity: '% 相似度',
            diffColumns: '不同列',
            file1: '文件 1',
            file2: '文件 2',
            calculating: '计算中...'
        },
        'ar': {
            file: 'ملف',
            rowCount: 'عدد الصفوف',
            rowsOnlyInFile: 'الصفوف في هذا الملف فقط',
            identicalRows: 'الصفوف المتطابقة',
            similarity: '% التشابه',
            diffColumns: 'الأعمدة المختلفة',
            file1: 'ملف 1',
            file2: 'ملف 2',
            calculating: 'جاري الحساب...'
        },
        'en': {
            file: 'File',
            rowCount: 'Row Count',
            rowsOnlyInFile: 'Rows only in this file',
            identicalRows: 'Identical rows',
            similarity: '% Similarity',
            diffColumns: 'Diff columns',
            file1: 'File 1',
            file2: 'File 2',
            calculating: 'Calculating...'
        }
    };
    
    return translations[currentLang] || translations['en'];
}

// Function to get file name with sheet name if exists
function getFileDisplayName(fileName, sheetName) {
    if (sheetName && sheetName.trim() !== '') {
        return `${fileName}:${sheetName}`;
    }
    return fileName;
}

// Function to update sheet information display
function updateSheetInfo(fileName, sheetNames, selectedSheet, fileNum) {
    const sheetInfoContainer = document.getElementById('sheetInfo');
    if (!sheetInfoContainer) return;
    
    if (sheetNames.length > 1) {
        const sheetInfo = document.createElement('div');
        sheetInfo.className = 'sheet-info';
        sheetInfo.dataset.fileNum = fileNum.toString();
        sheetInfo.innerHTML = `
            <div class="sheet-info-icon">ℹ</div>
            <div>
                <strong>${fileName}</strong> contains ${sheetNames.length} sheets. 
                Using: <strong>${selectedSheet}</strong>
                <br><small>Available sheets: ${sheetNames.join(', ')}</small>
            </div>
        `;
        
        // Remove any existing sheet info for this file
        const existingInfos = sheetInfoContainer.querySelectorAll('.sheet-info');
        existingInfos.forEach(info => {
            if (info.dataset.fileNum === fileNum.toString()) {
                info.remove();
            }
        });
        
        sheetInfoContainer.appendChild(sheetInfo);
        sheetInfoContainer.style.display = 'flex';
    } else {
        // If this file has only one sheet, remove its sheet info
        const allInfos = sheetInfoContainer.querySelectorAll('.sheet-info');
        allInfos.forEach(info => {
            if (info.dataset.fileNum === fileNum.toString()) {
                info.remove();
            }
        });
        
        if (sheetInfoContainer.children.length === 0) {
            sheetInfoContainer.style.display = 'none';
        }
    }
}

// Function to populate sheet selector dropdown
function populateSheetSelector(sheetNames, fileNum, selectedSheet) {
    const sheetSelector = document.getElementById(`sheetSelector${fileNum}`);
    const sheetSelect = document.getElementById(`sheetSelect${fileNum}`);
    
    if (!sheetSelector || !sheetSelect) return;
    
    if (sheetNames.length > 1) {
        // Clear existing options
        sheetSelect.innerHTML = '';
        
        // Add options for each sheet
        sheetNames.forEach((sheetName, index) => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            if (sheetName === selectedSheet) {
                option.selected = true;
            }
            sheetSelect.appendChild(option);
        });
        
        // Show the selector
        sheetSelector.style.display = 'block';
        
        // Add event listener for sheet selection change
        sheetSelect.onchange = function() {
            processSelectedSheet(fileNum, this.value);
        };
    } else {
        // Hide the selector for single-sheet files
        sheetSelector.style.display = 'none';
    }
}

// Function to process Excel sheet with optimized empty row/column detection
function processExcelSheetOptimized(sheet) {
    if (!sheet || !sheet['!ref']) {
        return [];
    }
    
    // Get the range of the sheet
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    // Find the actual data boundaries by scanning for non-empty cells
    let minRow = range.e.r + 1; // Start beyond the range
    let maxRow = -1;
    let minCol = range.e.c + 1; // Start beyond the range  
    let maxCol = -1;
    
    // Scan through all cells to find actual data boundaries
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
            const cell = sheet[cellAddress];
            
            // Check if cell has actual content (not just empty string or whitespace)
            if (cell && cell.v !== undefined && cell.v !== null) {
                const cellValue = cell.v.toString().trim();
                if (cellValue !== '') {
                    minRow = Math.min(minRow, row);
                    maxRow = Math.max(maxRow, row);
                    minCol = Math.min(minCol, col);
                    maxCol = Math.max(maxCol, col);
                }
            }
        }
    }
    
    // If no data found, return empty array
    if (minRow > maxRow || minCol > maxCol) {
        return [];
    }
    
    // Create the optimized range with only actual data
    const dataRange = {
        s: { r: minRow, c: minCol },
        e: { r: maxRow, c: maxCol }
    };
    
    // Create a new sheet object with only the data range
    const optimizedSheet = {};
    optimizedSheet['!ref'] = XLSX.utils.encode_range(dataRange);
    
    // Copy only the cells within the data range
    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const originalAddress = XLSX.utils.encode_cell({r: row, c: col});
            const newAddress = XLSX.utils.encode_cell({r: row - minRow, c: col - minCol});
            
            if (sheet[originalAddress]) {
                optimizedSheet[newAddress] = sheet[originalAddress];
            }
        }
    }
    
    // Update the range to start from 0,0
    optimizedSheet['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: maxRow - minRow, c: maxCol - minCol }
    });
    
    // Convert to JSON with the optimized range
    const json = XLSX.utils.sheet_to_json(optimizedSheet, {
        header: 1, 
        defval: '',
        raw: true,          // Get raw values to preserve full precision
        dateNF: 'yyyy-mm-dd hh:mm:ss'  // More complete date format
    });
    
    // Additional safety: filter out any remaining empty rows that might slip through
    const filteredJson = json.filter(row => 
        Array.isArray(row) && row.some(cell => 
            cell !== null && cell !== undefined && cell.toString().trim() !== ''
        )
    );
    
    return filteredJson;
}

// Function to process the selected sheet
function processSelectedSheet(fileNum, selectedSheetName) {
    // Clear previous comparison results when changing sheets
    clearComparisonResults();
    
    const workbook = fileNum === 1 ? workbook1 : workbook2;
    const fileName = fileNum === 1 ? fileName1 : fileName2;
    
    // Store the current sheet name
    if (fileNum === 1) {
        sheetName1 = selectedSheetName;
    } else {
        sheetName2 = selectedSheetName;
    }
    
    if (!workbook || !workbook.Sheets[selectedSheetName]) return;
    
    // Show loading indicator for large files
    const tableElement = document.getElementById(fileNum === 1 ? 'table1' : 'table2');
    tableElement.innerHTML = '<div style="text-align: center; padding: 20px;">Loading sheet...</div>';
    
    // Use setTimeout to allow UI to update
    setTimeout(() => {
        const sheet = workbook.Sheets[selectedSheetName];
        let json = processExcelSheetOptimized(sheet);
        
        // Check both file size and column count limits simultaneously
        let maxCols = 0;
        for (let i = 0; i < json.length; i++) {
            if (json[i] && json[i].length > maxCols) {
                maxCols = json[i].length;
            }
        }
        const rowsExceeded = json.length > MAX_ROWS_LIMIT;
        const colsExceeded = maxCols > MAX_COLS_LIMIT;
        
        if (rowsExceeded || colsExceeded) {
            tableElement.innerHTML = generateLimitErrorMessage(
                'rows', json.length, MAX_ROWS_LIMIT, '', 
                'columns', maxCols, MAX_COLS_LIMIT
            );
            // Clear data
            if (fileNum === 1) {
                data1 = [];
            } else {
                data2 = [];
            }
            return;
        }
        
        // Process the data efficiently
        json = json.filter(row => Array.isArray(row) && row.some(cell => (cell !== null && cell !== undefined && cell.toString().trim() !== '')));
        
        // Normalize headers to uppercase
        json = normalizeHeaders(json);
        
        // Analyze column types and apply appropriate conversions
        json = processColumnTypes(json);
        
        json = removeEmptyColumns(json);
        json = json.map(row => {
            if (!Array.isArray(row)) return row;
            return row.map(cell => roundDecimalNumbers(cell));
        });
        
        // Update the data and preview
        if (fileNum === 1) {
            data1 = json;
            renderPreview(json, 'table1');
        } else {
            data2 = json;
            renderPreview(json, 'table2');
        }
        
        // Update the sheet info
        updateSheetInfo(fileName, workbook.SheetNames, selectedSheetName, fileNum);
    }, 10);
}

// Function to analyze column types and apply appropriate conversions
function processColumnTypes(data) {
    if (!data || data.length === 0) return data;
    
    // Get column count
    const maxCols = Math.max(...data.map(row => Array.isArray(row) ? row.length : 0));
    if (maxCols === 0) return data;
    
    // Analyze each column
    const columnTypes = [];
    const headers = data[0] || [];
    
    for (let colIndex = 0; colIndex < maxCols; colIndex++) {
        // Extract column values (skip header row)
        const columnValues = [];
        for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
            const value = data[rowIndex] && data[rowIndex][colIndex];
            if (value !== null && value !== undefined && value.toString().trim() !== '') {
                columnValues.push(value);
            }
        }
        
        // Determine column type
        const columnHeader = headers[colIndex] || '';
        const isDateCol = isDateColumn(columnValues, columnHeader);
        columnTypes[colIndex] = isDateCol ? 'date' : 'other';
        
        // Column type detection completed
    }
    
    // Apply conversions based on column types
    const processedData = data.map((row, rowIndex) => {
        if (!Array.isArray(row) || rowIndex === 0) return row; // Skip header
        
        return row.map((cell, colIndex) => {
            if (columnTypes[colIndex] === 'date') {
                return convertExcelDateNormalized(cell, true); // Force date conversion
            } else {
                // For non-date columns, don't convert numbers to dates
                return cell;
            }
        });
    });
    
    return processedData;
}

// Helper function to detect if a column likely contains dates
function isDateColumn(columnValues, columnHeader = '') {
    if (!columnValues || columnValues.length === 0) return false;
    
    let dateCount = 0;
    let numberCount = 0;
    let totalCount = 0;
    let potentialExcelDates = 0;
    
    // Check if column header suggests this is a date column
    const headerLower = columnHeader.toString().toLowerCase();
    const dateKeywords = ['date', 'time', 'created', 'modified', 'updated', 'birth', 'дата', 'время', 'создан', 'изменен', 'обновлен'];
    const headerSuggestsDate = dateKeywords.some(keyword => headerLower.includes(keyword));
    
    for (let value of columnValues) {
        if (value && value.toString().trim() !== '') {
            totalCount++;
            
            // Check if value looks like a date
            if (typeof value === 'string') {
                // Check for date patterns
                if (value.match(/^\d{4}-\d{1,2}-\d{1,2}$/) || // YYYY-MM-DD
                    value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/) || // MM/DD/YYYY
                    value.match(/^\d{1,2}-\d{1,2}-\d{4}$/) || // MM-DD-YYYY
                    value.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/) || // YYYY/MM/DD
                    value.includes('T') || value.includes('GMT') || value.includes('UTC')) {
                    dateCount++;
                }
            } else if (value instanceof Date) {
                dateCount++;
            } else if (typeof value === 'number') {
                numberCount++;
                // Check if number could be an Excel date (exclude small numbers like 1015, 1011)
                // Start from 1980 (Excel serial ~29221) to avoid false positives
                if (value >= 29221 && value <= 219146) {
                    potentialExcelDates++;
                }
            }
        }
    }
    
    // Column is considered a date column if:
    // 1. Header suggests dates AND more than 30% are potential dates, OR
    // 2. More than 50% of values are explicit dates, OR
    // 3. More than 70% of values are numbers that could be Excel dates, OR
    // 4. Mix of dates and potential Excel dates makes up >60% of values
    if (totalCount === 0) return false;
    
    const dateRatio = dateCount / totalCount;
    const potentialDateRatio = potentialExcelDates / totalCount;
    const combinedDateRatio = (dateCount + potentialExcelDates) / totalCount;
    
    return (headerSuggestsDate && combinedDateRatio > 0.3) ||
           dateRatio > 0.5 || 
           (potentialDateRatio > 0.5 && numberCount > 0) || 
           combinedDateRatio > 0.6;
}

// Helper function to format date and time with optional seconds
function formatDate(year, month, day, hour = null, minute = null, seconds = null) {
    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    
    // Check if we have any meaningful time (not 00:00:00)
    const hasTime = (hour !== null && minute !== null) && 
                   !(hour === 0 && minute === 0 && (seconds === null || seconds === 0));
    
    if (hasTime) {
        result += ' ' + String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
        
        // Add seconds if provided
        if (seconds !== null) {
            result += ':' + String(seconds).padStart(2, '0');
        }
    }
    
    return result;
}

// Helper function to normalize date-time string and remove 00:00:00 consistently
function normalizeDateTime(dateTimeString) {
    // Check if it ends with 00:00:00 or 00:00 and remove it
    if (dateTimeString.match(/\s+00:00:00$/)) {
        return dateTimeString.replace(/\s+00:00:00$/, '');
    }
    if (dateTimeString.match(/\s+00:00$/)) {
        return dateTimeString.replace(/\s+00:00$/, '');
    }
    return dateTimeString;
}

// Function to convert Excel serial date to proper date format
function convertExcelDate(value, isInDateColumn = false) {
    // Handle Date objects that might be created by XLSX library
    if (value instanceof Date) {
        if (!isNaN(value.getTime())) {
            // Use local date components to avoid timezone shifts
            const year = value.getFullYear();
            const month = value.getMonth() + 1;
            const day = value.getDate();
            const hour = value.getHours();
            const minute = value.getMinutes();
            const second = value.getSeconds();
            
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                // Use formatDate which automatically handles 00:00:00 removal
                return formatDate(year, month, day, hour, minute, second);
            }
        }
    }
    
    // Check if value is a number that could be an Excel date serial number
    if (typeof value === 'number' && value > 1 && value < 300000) {
        // Only convert numbers to dates if explicitly in a date column context
        // Don't auto-convert numbers like 1015, 1011 unless we're sure it's a date column
        // Use realistic date range starting from 1980
        if (isInDateColumn && value >= 29221 && value <= 219146) { // Around 1980-2500, exclude unrealistic early dates
            
            // Handle Excel serial numbers with more precision for time parts
            const wholeDays = Math.floor(value);
            const timeFraction = value - wholeDays;
            
            // Excel epoch: January 1, 1900 (accounting for Excel's leap year bug)
            // Avoid timezone issues by calculating date components directly
            let days = wholeDays;
            
            // Excel incorrectly treats 1900 as a leap year, so for dates after Feb 28, 1900, subtract 1 day
            if (wholeDays > 59) { // After Feb 28, 1900 (Excel serial 59)
                days = wholeDays - 1; // Correct for Excel's leap year bug
            }
            
            // Calculate date components directly to avoid timezone shifts
            // Excel serial 1 = January 1, 1900
            const excelYear = 1900;
            let year = excelYear;
            let dayOfYear = days;
            
            // Calculate year and remaining days
            while (dayOfYear > 365) {
                const isLeapYear = (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
                const daysInYear = isLeapYear ? 366 : 365;
                if (dayOfYear > daysInYear) {
                    dayOfYear -= daysInYear;
                    year++;
                } else {
                    break;
                }
            }
            
            // Calculate month and day
            const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
            const isLeapYear = (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
            if (isLeapYear) daysInMonth[1] = 29;
            
            let month = 1;
            let day = dayOfYear;
            
            for (let i = 0; i < 12; i++) {
                if (day <= daysInMonth[i]) {
                    month = i + 1;
                    break;
                } else {
                    day -= daysInMonth[i];
                }
            }
            
            // Calculate time from fractional part
            const totalSeconds = Math.round(timeFraction * 24 * 60 * 60);
            const hours = Math.floor(totalSeconds / 3600);
            const minutes = Math.floor((totalSeconds % 3600) / 60);
            const seconds = totalSeconds % 60;
            
            if (year >= 1900 && year <= 2500) {
                // Use formatDate for consistent 00:00:00 handling
                const hasTime = timeFraction > 0 || hours !== 0 || minutes !== 0 || seconds !== 0;
                return formatDate(year, month, day, hasTime ? hours : null, hasTime ? minutes : null, hasTime ? seconds : null);
            }
        }
    }
    
    // Check if value is already a date string in various formats
    if (typeof value === 'string') {
        // First check if it's already in correct YYYY-MM-DD HH:MM:SS format - normalize it
        if (value.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
            return normalizeDateTime(value); // Remove 00:00:00 if present
        }
        
        // Check if it's in YYYY-MM-DD HH:MM format - add seconds and normalize
        if (value.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/)) {
            return normalizeDateTime(value + ':00'); // Add missing seconds and normalize
        }
        
        // Check if it's already in YYYY-MM-DD format - keep it as is
        let isoDateMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (isoDateMatch) {
            const year = parseInt(isoDateMatch[1]);
            const month = parseInt(isoDateMatch[2]);
            const day = parseInt(isoDateMatch[3]);
            if (day <= 31 && month <= 12 && year >= 1900 && year <= 2100) {
                // Already in correct format, just normalize
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Check for YYYY-MM-DD HH:MM format - keep it as is
        let isoDateTimeMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (isoDateTimeMatch) {
            const year = parseInt(isoDateTimeMatch[1]);
            const month = parseInt(isoDateTimeMatch[2]);
            const day = parseInt(isoDateTimeMatch[3]);
            const hour = parseInt(isoDateTimeMatch[4]);
            const minute = parseInt(isoDateTimeMatch[5]);
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                return value; // Already in correct format
            }
        }
        
        // Handle YYYY-MM-DD HH:MM:SS AM/PM format (like "2022-01-01 10:41:52 AM")
        let isoDateTimeAMPMMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+(AM|PM)$/i);
        if (isoDateTimeAMPMMatch) {
            const year = parseInt(isoDateTimeAMPMMatch[1]);
            const month = parseInt(isoDateTimeAMPMMatch[2]); 
            const day = parseInt(isoDateTimeAMPMMatch[3]);
            let hour = parseInt(isoDateTimeAMPMMatch[4]);
            const minute = parseInt(isoDateTimeAMPMMatch[5]);
            const second = isoDateTimeAMPMMatch[6] ? parseInt(isoDateTimeAMPMMatch[6]) : 0;
            const ampm = isoDateTimeAMPMMatch[7].toUpperCase();
            
            // Convert 12-hour to 24-hour format
            if (ampm === 'AM') {
                if (hour === 12) hour = 0; // 12:xx AM becomes 00:xx
            } else { // PM
                if (hour !== 12) hour += 12; // 1-11 PM becomes 13-23, 12 PM stays 12
            }
            
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59 && second >= 0 && second <= 59) {
                const result = formatDate(year, month, day, hour, minute, second);
                return result;
            }
        }
        
        // Handle MM/DD/YY HH:MM:SS format (with seconds)
        let dateTimeMatchSec = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
        if (dateTimeMatchSec) {
            const first = parseInt(dateTimeMatchSec[1]);
            const second = parseInt(dateTimeMatchSec[2]);
            let year = parseInt(dateTimeMatchSec[3]);
            const hour = parseInt(dateTimeMatchSec[4]);
            const minute = parseInt(dateTimeMatchSec[5]);
            const sec = parseInt(dateTimeMatchSec[6]);
            
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            // For American format like 8/1/22, interpret as MM/DD
            if (first <= 12 && second <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59 && sec >= 0 && sec <= 59) {
                const month = first;
                const day = second;
                return formatDate(year, month, day, hour, minute, sec);
            }
        }
        
        // Handle MM/DD/YY HH:MM:SS AM/PM format (like "8/1/22 10:41:52 AM")
        let shortYearAMPMMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+(AM|PM)$/i);
        if (shortYearAMPMMatch) {
            const first = parseInt(shortYearAMPMMatch[1]);
            const second = parseInt(shortYearAMPMMatch[2]);
            let year = parseInt(shortYearAMPMMatch[3]);
            let hour = parseInt(shortYearAMPMMatch[4]);
            const minute = parseInt(shortYearAMPMMatch[5]);
            const seconds = shortYearAMPMMatch[6] ? parseInt(shortYearAMPMMatch[6]) : 0;
            const ampm = shortYearAMPMMatch[7].toUpperCase();
            
            // Convert to 24-hour format
            if (ampm === 'AM' && hour === 12) {
                hour = 0;  // 12 AM becomes 00
            } else if (ampm === 'PM' && hour !== 12) {
                hour += 12;  // 1-11 PM becomes 13-23
            }
            // 12 PM stays as 12
            
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            // Smart format detection: prefer MM/DD for typical American patterns
            if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                // If first > 12, it must be DD/MM (European)
                if (first > 12) {
                    day = first;
                    month = second;
                } else if (second > 12) {
                    // If second > 12, it must be MM/DD (American)
                    month = first;
                    day = second;
                } else {
                    // Both <= 12, use American MM/DD format by default
                    month = first;
                    day = second;
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return formatDate(year, month, day, hour, minute, seconds);
                }
            }
        }

        // Handle MM/DD/YY HH:MM format (like "8/1/22 10:41")
        let dateTimeMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const first = parseInt(dateTimeMatch[1]);
            const second = parseInt(dateTimeMatch[2]);
            let year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            const seconds = dateTimeMatch[6] ? parseInt(dateTimeMatch[6]) : 0;
            
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            // Smart format detection: prefer MM/DD for typical American patterns
            if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                // If first > 12, it must be DD/MM (European)
                if (first > 12 && second <= 12) {
                    day = first;
                    month = second;
                } 
                // If second > 12, it must be MM/DD (American)
                else if (second > 12 && first <= 12) {
                    month = first;
                    day = second;
                }
                // If both <= 12, prefer MM/DD (American) as it's more common with time formats
                else if (first <= 12 && second <= 12) {
                    month = first;
                    day = second;
                }
                // If both > 12, invalid date - skip
                else {
                    return value;
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                    
                    // Always add seconds to preserve format consistency
                    result += ':' + String(seconds).padStart(2, '0');
                    
                    return result;
                }
            }
        }
        
        // Handle MM/DD/YYYY HH:MM:SS format (with seconds and full year)
        let dateTimeMatchFullSec = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
        if (dateTimeMatchFullSec) {
            const first = parseInt(dateTimeMatchFullSec[1]);
            const second = parseInt(dateTimeMatchFullSec[2]);
            const year = parseInt(dateTimeMatchFullSec[3]);
            const hour = parseInt(dateTimeMatchFullSec[4]);
            const minute = parseInt(dateTimeMatchFullSec[5]);
            const sec = parseInt(dateTimeMatchFullSec[6]);
            
            if (year >= 1900 && year <= 2100 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59 && sec >= 0 && sec <= 59) {
                let month, day;
                
                // Smart format detection
                if (first > 12 && second <= 12) {
                    // Must be DD/MM
                    day = first;
                    month = second;
                } else if (second > 12 && first <= 12) {
                    // Must be MM/DD
                    month = first;
                    day = second;
                } else if (first <= 12 && second <= 12) {
                    // Ambiguous case - prefer MM/DD for American format with time
                    month = first;
                    day = second;
                } else {
                    return value; // Invalid date
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return formatDate(year, month, day, hour, minute, sec);
                }
            }
        }
        
        // Handle MM/DD/YYYY HH:MM format (like "8/1/2022 10:41")
        dateTimeMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const first = parseInt(dateTimeMatch[1]);
            const second = parseInt(dateTimeMatch[2]);
            const year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            const seconds = dateTimeMatch[6] ? parseInt(dateTimeMatch[6]) : 0;
            
            if (year >= 1900 && year <= 2100 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                // Smart format detection
                if (first > 12 && second <= 12) {
                    // Must be DD/MM
                    day = first;
                    month = second;
                } else if (second > 12 && first <= 12) {
                    // Must be MM/DD
                    month = first;
                    day = second;
                } else if (first <= 12 && second <= 12) {
                    // Ambiguous case - prefer MM/DD for American format with time
                    month = first;
                    day = second;
                } else {
                    return value; // Invalid date
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                    
                    // Always add seconds to preserve format consistency
                    result += ':' + String(seconds).padStart(2, '0');
                    
                    return result;
                }
            }
        }
        
        // Handle MM/DD/YYYY HH:MM:SS AM/PM format (like "8/1/2022 10:41:52 AM")
        let dateTimeAMPMMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+(AM|PM)$/i);
        if (dateTimeAMPMMatch) {
            const first = parseInt(dateTimeAMPMMatch[1]);
            const second = parseInt(dateTimeAMPMMatch[2]);
            const year = parseInt(dateTimeAMPMMatch[3]);
            let hour = parseInt(dateTimeAMPMMatch[4]);
            const minute = parseInt(dateTimeAMPMMatch[5]);
            const second_time = dateTimeAMPMMatch[6] ? parseInt(dateTimeAMPMMatch[6]) : 0;
            const ampm = dateTimeAMPMMatch[7].toUpperCase();
            
            // Convert 12-hour to 24-hour format
            if (ampm === 'AM') {
                if (hour === 12) hour = 0; // 12:xx AM becomes 00:xx
            } else { // PM
                if (hour !== 12) hour += 12; // 1-11 PM becomes 13-23, 12 PM stays 12
            }
            
            if (year >= 1900 && year <= 2100 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                // Smart format detection
                if (first > 12 && second <= 12) {
                    // Must be DD/MM
                    day = first;
                    month = second;
                } else if (second > 12 && first <= 12) {
                    // Must be MM/DD
                    month = first;
                    day = second;
                } else if (first <= 12 && second <= 12) {
                    // Ambiguous case - prefer MM/DD for American format with AM/PM
                    month = first;
                    day = second;
                } else {
                    return value; // Invalid date
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                    
                    // Always add seconds to preserve format consistency
                    result += ':' + String(second_time).padStart(2, '0');
                    
                    return result;
                }
            }
        }
        
        // Handle AM/PM formats WITHOUT seconds (like "2022-01-01 10:41 AM")
        let ampmNoSecondsMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})\s+(AM|PM)$/i);
        if (ampmNoSecondsMatch) {
            const year = parseInt(ampmNoSecondsMatch[1]);
            const month = parseInt(ampmNoSecondsMatch[2]);
            const day = parseInt(ampmNoSecondsMatch[3]);
            let hour = parseInt(ampmNoSecondsMatch[4]);
            const minute = parseInt(ampmNoSecondsMatch[5]);
            const ampm = ampmNoSecondsMatch[6].toUpperCase();
            
            // Convert 12-hour to 24-hour format
            if (ampm === 'AM') {
                if (hour === 12) hour = 0; // 12:xx AM becomes 00:xx
            } else { // PM
                if (hour !== 12) hour += 12; // 1-11 PM becomes 13-23, 12 PM stays 12
            }
            
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                // Add :00 seconds since they weren't specified
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                       String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0') + ':00';
            }
        }
        
        // Handle MM/DD/YYYY AM/PM formats WITHOUT seconds (like "1/1/2022 10:41 AM")
        ampmNoSecondsMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})\s+(AM|PM)$/i);
        if (ampmNoSecondsMatch) {
            const first = parseInt(ampmNoSecondsMatch[1]);
            const second = parseInt(ampmNoSecondsMatch[2]);
            const year = parseInt(ampmNoSecondsMatch[3]);
            let hour = parseInt(ampmNoSecondsMatch[4]);
            const minute = parseInt(ampmNoSecondsMatch[5]);
            const ampm = ampmNoSecondsMatch[6].toUpperCase();
            
            // Convert 12-hour to 24-hour format
            if (ampm === 'AM') {
                if (hour === 12) hour = 0;
            } else { // PM
                if (hour !== 12) hour += 12;
            }
            
            if (year >= 1900 && year <= 2100 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                // Smart format detection
                if (first > 12 && second <= 12) {
                    day = first;
                    month = second;
                } else if (second > 12 && first <= 12) {
                    month = first;
                    day = second;
                } else if (first <= 12 && second <= 12) {
                    month = first;
                    day = second;
                } else {
                    return value;
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    // Add :00 seconds since they weren't specified
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0') + ':00';
                }
            }
        }
        
        // Handle MM/DD/YY AM/PM formats WITHOUT seconds (like "1/1/22 10:41 AM")
        ampmNoSecondsMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})\s+(AM|PM)$/i);
        if (ampmNoSecondsMatch) {
            const first = parseInt(ampmNoSecondsMatch[1]);
            const second = parseInt(ampmNoSecondsMatch[2]);
            let year = parseInt(ampmNoSecondsMatch[3]);
            let hour = parseInt(ampmNoSecondsMatch[4]);
            const minute = parseInt(ampmNoSecondsMatch[5]);
            const ampm = ampmNoSecondsMatch[6].toUpperCase();
            
            // Convert year
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            // Convert 12-hour to 24-hour format
            if (ampm === 'AM') {
                if (hour === 12) hour = 0;
            } else { // PM
                if (hour !== 12) hour += 12;
            }
            
            if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                if (first > 12) {
                    day = first;
                    month = second;
                } else if (second > 12) {
                    month = first;
                    day = second;
                } else {
                    month = first;
                    day = second;
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    // Add :00 seconds since they weren't specified
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0') + ':00';
                }
            }
        }
        
        // Handle DD/MM/YY HH:MM format (European with time)
        dateTimeMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const first = parseInt(dateTimeMatch[1]);
            const second = parseInt(dateTimeMatch[2]);
            let year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            // For European format, interpret as DD/MM if first > 12
            if (first > 12 && second <= 12 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                const day = first;
                const month = second;
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                       String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
            }
        }
        
        // Handle MM.DD.YY HH:MM format (dots with time)
        dateTimeMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const month = parseInt(dateTimeMatch[1]);
            const day = parseInt(dateTimeMatch[2]);
            let year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' +
                       String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
            }
        }
        
        // Handle MM.DD.YYYY HH:MM format (dots with full year and time)
        dateTimeMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const month = parseInt(dateTimeMatch[1]);
            const day = parseInt(dateTimeMatch[2]);
            const year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 1900 && year <= 2100 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' +
                       String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
            }
        }
        
        // Check if it's a Date object converted to string that might have timezone issues
        // This can happen when CSV parsing creates Date objects that get stringified
        if (value.includes('T') || value.includes('GMT') || value.includes('UTC')) {
            try {
                // For ISO strings, extract date part directly to avoid timezone conversion
                if (value.includes('T')) {
                    // Check if it's a full datetime ISO string
                    const fullDateTimeMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})T(\d{1,2}):(\d{1,2}):(\d{1,2})/);
                    if (fullDateTimeMatch) {
                        const year = parseInt(fullDateTimeMatch[1]);
                        const month = parseInt(fullDateTimeMatch[2]);
                        const day = parseInt(fullDateTimeMatch[3]);
                        const hour = parseInt(fullDateTimeMatch[4]);
                        const minute = parseInt(fullDateTimeMatch[5]);
                        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                            return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                                   String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                        }
                    }
                    
                    // Otherwise extract just date part
                    const datePart = value.split('T')[0];
                    const dateMatch = datePart.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
                    if (dateMatch) {
                        const year = parseInt(dateMatch[1]);
                        const month = parseInt(dateMatch[2]);
                        const day = parseInt(dateMatch[3]);
                        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                            return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                        }
                    }
                }
                
                // Fallback: Parse as Date and preserve both date and time if available
                const dateObj = new Date(value);
                if (!isNaN(dateObj.getTime())) {
                    // Use local date components to avoid timezone shifts
                    const year = dateObj.getFullYear();
                    const month = dateObj.getMonth() + 1;
                    const day = dateObj.getDate();
                    const hour = dateObj.getHours();
                    const minute = dateObj.getMinutes();
                    
                    if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                        // Check if original value contains time information
                        if (value.includes(':') || hour !== 0 || minute !== 0) {
                            return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                                   String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                        } else {
                            return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                        }
                    }
                }
            } catch (e) {
                // If parsing fails, continue with other methods
            }
        }
        
        // Try to parse common date formats
        let dateMatch;
        
        // DD.MM.YYYY format (dots as separators)
        dateMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const month = parseInt(dateMatch[2]);
            const year = parseInt(dateMatch[3]);
            if (day <= 31 && month <= 12 && year >= 1900 && year <= 2100) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle date formats with slashes (DD/MM/YYYY or MM/DD/YYYY) - smart detection
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dateMatch) {
            const first = parseInt(dateMatch[1]);
            const second = parseInt(dateMatch[2]);
            const year = parseInt(dateMatch[3]);
            
            if (year >= 1900 && year <= 2100) {
                let month, day;
                
                if (first > 12 && second <= 12) {
                    // Must be DD/MM
                    day = first;
                    month = second;
                } else if (second > 12 && first <= 12) {
                    // Must be MM/DD
                    month = first;
                    day = second;
                } else if (first <= 12 && second <= 12) {
                    // Ambiguous case - prefer MM/DD (American format for consistency with datetime)
                    month = first;
                    day = second;
                } else {
                    return value; // Invalid date
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                }
            }
        }
        
        // MM/DD/YY format (short American format) - often result of Excel reading
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
        if (dateMatch) {
            const first = parseInt(dateMatch[1]);
            const second = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            // Smart format detection for dates without time
            let month, day;
            
            if (first > 12 && second <= 12) {
                // Must be DD/MM
                day = first;
                month = second;
            } else if (second > 12 && first <= 12) {
                // Must be MM/DD
                month = first;
                day = second;
            } else if (first <= 12 && second <= 12) {
                // Ambiguous case - prefer MM/DD (American format)
                month = first;
                day = second;
            } else {
                return value; // Invalid date
            }
            
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // MM.DD.YY format (short year)
        dateMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);
        if (dateMatch) {
            const month = parseInt(dateMatch[1]);
            const day = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // MM/DD/YY format (short year with slashes)
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
        if (dateMatch) {
            const month = parseInt(dateMatch[1]);
            const day = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle Excel text dates like "1-May-2025", "15-Dec-2024"
        dateMatch = value.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            const year = parseInt(dateMatch[3]);
            
            const monthMap = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const month = monthMap[monthName];
            if (month && day <= 31 && year >= 1900 && year <= 2100) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle Excel text dates with short year like "01-May-25", "15-Dec-24"
        dateMatch = value.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            const monthMap = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const month = monthMap[monthName];
            if (month && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle Excel text dates like "May-1-2025", "Dec-15-2024"
        dateMatch = value.match(/^([A-Za-z]{3})-(\d{1,2})-(\d{4})$/);
        if (dateMatch) {
            const monthName = dateMatch[1].toLowerCase();
            const day = parseInt(dateMatch[2]);
            const year = parseInt(dateMatch[3]);
            
            const monthMap = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const month = monthMap[monthName];
            if (month && day <= 31 && year >= 1900 && year <= 2100) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle Excel text dates with short year like "May-01-25", "Dec-15-24"
        dateMatch = value.match(/^([A-Za-z]{3})-(\d{1,2})-(\d{2})$/);
        if (dateMatch) {
            const monthName = dateMatch[1].toLowerCase();
            const day = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            const monthMap = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const month = monthMap[monthName];
            if (month && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle Excel text dates with day first and short year like "05-May-25", "15-Dec-24"
        dateMatch = value.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            const monthMap = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const month = monthMap[monthName];
            if (month && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle full month names like "1 May 2025", "15 December 2024"
        dateMatch = value.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            const year = parseInt(dateMatch[3]);
            
            const monthMap = {
                'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
                'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const month = monthMap[monthName];
            if (month && day <= 31 && year >= 1900 && year <= 2100) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle full month names with short year like "1 May 25", "15 December 24"
        dateMatch = value.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{2})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            const monthMap = {
                'january': 1, 'february': 2, 'march': 3, 'april': 4, 'may': 5, 'june': 6,
                'july': 7, 'august': 8, 'september': 9, 'october': 10, 'november': 11, 'december': 12,
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            };
            
            const month = monthMap[monthName];
            if (month && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // Handle cases where Excel might have parsed "2025-05-01" as a calculation
        // Look for patterns like "2019" or "2020" etc that might be date calculations
        if (value.includes('-') && !isoDateMatch) {
            // Try to parse as potential date calculation result
            let parts = value.split('-');
            if (parts.length === 3) {
                let year = parseInt(parts[0]);
                let month = parseInt(parts[1]);
                let day = parseInt(parts[2]);
                
                // Check if this could be a valid date
                if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                }
            }
        }
    }
    
    // Return original value if it's not a date
    return value;
}

// Wrapper for convertExcelDate that ensures consistent 00:00:00 removal
function convertExcelDateNormalized(value, isInDateColumn = false) {
    const result = convertExcelDate(value, isInDateColumn);
    
    // If result is a date string (YYYY-MM-DD format), normalize it
    if (typeof result === 'string' && result.match(/^\d{4}-\d{2}-\d{2}/)) {
        return normalizeDateTime(result);
    }
    
    return result;
}

// Function to round decimal numbers to 2 decimal places and remove .0 for integers
function roundDecimalNumbers(value) {
    // If it's a number, process it
    if (typeof value === 'number') {
        // Round to 2 decimal places to fix floating point issues
        let rounded = Math.round(value * 100) / 100;
        
        // If it's an integer after rounding, return as integer
        if (Number.isInteger(rounded)) {
            return rounded;
        }
        
        return rounded;
    }
    
    // If it's a string that represents a number, convert and process
    if (typeof value === 'string' && !isNaN(value) && !isNaN(parseFloat(value))) {
        const numValue = parseFloat(value);
        
        // Round to 2 decimal places to fix floating point issues
        let rounded = Math.round(numValue * 100) / 100;
        
        // If it's an integer after rounding, return as integer
        if (Number.isInteger(rounded)) {
            return rounded;
        }
        
        return rounded;
    }
    
    return value;
}

function removeEmptyColumns(data) {
    if (!data || data.length === 0) return data;
    
    // Determine the maximum number of columns
    let maxCols = 0;
    for (let i = 0; i < data.length; i++) {
        if (data[i] && data[i].length > maxCols) {
            maxCols = data[i].length;
        }
    }
    if (maxCols === 0) return data;
    
    // Find columns that are completely empty (optimized for large files)
    let columnsToKeep = [];
    
    // For very large files (>1000 rows), use sampling approach
    const isLargeFile = data.length > 1000;
    const checkRows = isLargeFile ? Math.min(data.length, 100) : data.length;
    
    columnLoop: for (let col = 0; col < maxCols; col++) {
        // Quick check: examine sample rows first
        for (let row = 0; row < checkRows; row++) {
            if (data[row] && col < data[row].length) {
                let cellValue = data[row][col];
                if (cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '') {
                    columnsToKeep.push(col);
                    continue columnLoop;
                }
            }
        }
        
        // For large files, if no data found in sample, check every 50th row
        if (isLargeFile && data.length > checkRows) {
            for (let row = checkRows; row < data.length; row += 50) {
                if (data[row] && col < data[row].length) {
                    let cellValue = data[row][col];
                    if (cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '') {
                        columnsToKeep.push(col);
                        continue columnLoop;
                    }
                }
            }
        }
    }
    
    // Create new rows with only non-empty columns (use more efficient approach)
    let result = new Array(data.length);
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (!row) {
            result[i] = [];
            continue;
        }
        
        result[i] = new Array(columnsToKeep.length);
        for (let j = 0; j < columnsToKeep.length; j++) {
            const colIndex = columnsToKeep[j];
            result[i][j] = colIndex < row.length ? row[colIndex] : '';
        }
    }
    
    // Optimize date normalization - only for columns that might contain dates
    if (result.length > 0) {
        const numCols = result[0].length;
        
        for (let col = 0; col < numCols; col++) {
            // Sample only small portion for date detection (max 20 rows)
            const sampleSize = Math.min(result.length, 20);
            const columnValues = [];
            for (let i = 0; i < sampleSize; i++) {
                columnValues.push(result[i][col]);
            }
            
            const columnHeader = result.length > 0 ? result[0][col] : '';
            const isDateCol = isDateColumn(columnValues, columnHeader);
            
            // Only process dates if column is identified as date column
            if (isDateCol) {
                for (let row = 0; row < result.length; row++) {
                    if (result[row][col] !== null && result[row][col] !== undefined && result[row][col] !== '') {
                        result[row][col] = convertExcelDateNormalized(result[row][col], true);
                    }
                }
            }
        }
    }
    
    return result;
}

// Function to normalize row lengths (ensure all rows have same number of columns)
function normalizeRowLengths(data) {
    if (!data || data.length === 0) return data;
    
    // Find the maximum number of columns
    let maxCols = 0;
    for (let i = 0; i < data.length; i++) {
        if (data[i] && data[i].length > maxCols) {
            maxCols = data[i].length;
        }
    }
    
    // Pad all rows to have the same number of columns
    return data.map(row => {
        if (!row) return new Array(maxCols).fill('');
        while (row.length < maxCols) {
            row.push('');
        }
        return row;
    });
}

// Function to detect CSV delimiter
function detectCSVDelimiter(csvText) {
    const firstLine = csvText.split(/\r?\n/)[0];
    const delimiters = [',', ';', '\t', '|'];
    let maxCount = 0;
    let bestDelimiter = ',';
    
    for (let delimiter of delimiters) {
        const count = firstLine.split(delimiter).length - 1;
        if (count > maxCount) {
            maxCount = count;
            bestDelimiter = delimiter;
        }
    }
    
    return bestDelimiter;
}

// Function to parse CSV manually to avoid timezone issues
function parseCSV(csvText) {
    const lines = csvText.split(/\r?\n/);
    const result = [];
    
    // Detect delimiter from first line
    const delimiter = detectCSVDelimiter(csvText);
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.trim() === '') continue;
        
        const row = [];
        let current = '';
        let inQuotes = false;
        
        for (let j = 0; j < line.length; j++) {
            const char = line[j];
            
            if (char === '"') {
                if (inQuotes && line[j + 1] === '"') {
                    // Escaped quote
                    current += '"';
                    j++; // Skip next quote
                } else {
                    // Toggle quote state
                    inQuotes = !inQuotes;
                }
            } else if (char === delimiter && !inQuotes) {
                // End of field
                row.push(parseCSVValue(current.trim()));
                current = '';
            } else {
                current += char;
            }
        }
        
        // Add the last field
        row.push(parseCSVValue(current.trim()));
        
        // Only add non-empty rows
        if (row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== '')) {
            result.push(row);
        }
    }
    
    return result;
}

// Function to parse CSV value while preserving dates as strings
function parseCSVValue(value) {
    // Remove surrounding quotes if present
    if (value.startsWith('"') && value.endsWith('"')) {
        value = value.slice(1, -1);
    }
    
    // Handle empty values and 'null' strings
    if (value === '' || value === null || value === undefined || value.toLowerCase() === 'null') {
        return '';
    }
    
    // Clean up encoding issues - replace ? with common date separators
    if (value.includes('?')) {
        // Try to detect if this might be a date with encoding issues
        const cleanValue = value.replace(/\?/g, '.');
        value = cleanValue;
    }
    
    // Check if it looks like a date - if so, keep it as string
    if (value.match(/^\d{4}-\d{1,2}-\d{1,2}$/)) {
        // ISO format YYYY-MM-DD - keep as is
        return value;
    } else if (value.match(/^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{4}$/) ||
               value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/) ||
               value.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
        // Convert DD/MM/YYYY to YYYY-MM-DD
        const parts = value.split(/[\/\.]/);
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        
        // Validate date parts
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
        return value;
    } else if (value.match(/^\d{1,2}\.\d{1,2}\.\d{2}$/)) {
        // Handle MM.DD.YY format (e.g., "05.01.25" -> 2025-05-01)
        const parts = value.split('.');
        const month = parseInt(parts[0], 10);
        const day = parseInt(parts[1], 10);
        let year = parseInt(parts[2], 10);
        
        // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
        year = year <= 30 ? 2000 + year : 1900 + year;
        
        // Validate date parts
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
        return value;
    }
    return value; // Keep dates as strings
    
    // Try to parse as number (handle both . and , as decimal separators)
    let numValue = value.replace(',', '.'); // Convert comma to dot for parsing
    if (numValue !== '' && !isNaN(numValue) && !isNaN(parseFloat(numValue))) {
        const num = parseFloat(numValue);
        return Number.isInteger(num) ? num : num;
    }
    
    return value;
}

function handleFile(file, num) {
    if (!file) return;
    
    // Show loading indicator immediately
    const tableElement = document.getElementById(num === 1 ? 'table1' : 'table2');
    tableElement.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">📊 Loading file... Please wait</div>';
    
    // Clear previous data for security
    if (num === 1) {
        data1 = [];
        fileName1 = file.name;
        workbook1 = null;
        sheetName1 = ''; // Clear sheet name for CSV files
    } else {
        data2 = [];
        fileName2 = file.name;
        workbook2 = null;
        sheetName2 = ''; // Clear sheet name for CSV files
    }
    
    // Clear result tables
    document.getElementById('result').innerHTML = '';
    document.getElementById('summary').innerHTML = '';
    
    // Hide filter controls
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
    
    // Clear any existing sheet info for this file slot
    const sheetInfoContainer = document.getElementById('sheetInfo');
    if (sheetInfoContainer) {
        // Remove existing sheet info for this file number
        const existingInfos = Array.from(sheetInfoContainer.querySelectorAll('.sheet-info'));
        existingInfos.forEach(info => {
            // Remove sheet info that might be for the previous file in this slot
            if (num === 1 && info.dataset.fileNum === '1') {
                info.remove();
            } else if (num === 2 && info.dataset.fileNum === '2') {
                info.remove();
            }
        });
        
        if (sheetInfoContainer.children.length === 0) {
            sheetInfoContainer.style.display = 'none';
        }
    }
    
    // Hide sheet selector initially
    const sheetSelector = document.getElementById(`sheetSelector${num}`);
    if (sheetSelector) {
        sheetSelector.style.display = 'none';
    }
    
    // Check if it's a CSV file
    const isCSV = file.name.toLowerCase().endsWith('.csv');
    
    if (isCSV) {
        // Handle CSV files with text reading to avoid timezone issues
        const reader = new FileReader();
        reader.onload = function(e) {
            // Use setTimeout to allow UI to update with loading message
            setTimeout(() => {
                const csvText = e.target.result;
                const json = parseCSV(csvText);
                
                // Check both file size and column count limits simultaneously
                let maxCols = 0;
                for (let i = 0; i < json.length; i++) {
                    if (json[i] && json[i].length > maxCols) {
                        maxCols = json[i].length;
                    }
                }
                const rowsExceeded = json.length > MAX_ROWS_LIMIT;
                const colsExceeded = maxCols > MAX_COLS_LIMIT;
                
                if (rowsExceeded || colsExceeded) {
                    tableElement.innerHTML = generateLimitErrorMessage(
                        'rows', json.length, MAX_ROWS_LIMIT, '', 
                        'columns', maxCols, MAX_COLS_LIMIT
                    );
                    // Clear data
                    if (num === 1) {
                        data1 = [];
                        fileName1 = '';
                        workbook1 = null;
                    } else {
                        data2 = [];
                        fileName2 = '';
                        workbook2 = null;
                    }
                    return;
                }
                
                // Show progress for large files
                if (json.length > 1000) {
                    tableElement.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚙️ Processing large CSV file... Please wait</div>';
                }
                
                // Use setTimeout for each processing step to allow UI updates
                setTimeout(() => {
                    // Normalize headers to uppercase first
                    const headersNormalizedJson = normalizeHeaders(json);
                    
                    // Apply column-based date conversion (for CSV text dates like "05.01.25")
                    const dateConvertedJson = processColumnTypes(headersNormalizedJson);
                    
                    setTimeout(() => {
                        // Normalize row lengths
                        const normalizedJson = normalizeRowLengths(dateConvertedJson);
                        
                        // Remove empty columns and normalize dates intelligently
                        const cleanedJson = removeEmptyColumns(normalizedJson);
                        
                        setTimeout(() => {
                            // Round decimal numbers
                            const processedJson = cleanedJson.map(row => {
                                if (!Array.isArray(row)) return row;
                                return row.map(cell => roundDecimalNumbers(cell));
                            });
                            
                            if (num === 1) {
                                data1 = processedJson;
                                renderPreview(processedJson, 'table1');
                            } else {
                                data2 = processedJson;
                                renderPreview(processedJson, 'table2');
                            }
                        }, 10);
                    }, 10);
                }, 10);
            }, 10);
        };
        reader.readAsText(file, 'UTF-8');
    } else {
        // Handle Excel files with XLSX library
        const reader = new FileReader();
        reader.onload = function(e) {
            // Use setTimeout to allow UI to update
            setTimeout(() => {
                let data = new Uint8Array(e.target.result);
                // Use cellDates option to preserve date formatting and avoid timezone issues
                let workbook = XLSX.read(data, {
                    type: 'array',
                    cellDates: false,    // Get raw values instead of Date objects
                    UTC: false  // Use local timezone to avoid shifts
                });
                
                // Store workbook for sheet selection
                if (num === 1) {
                    workbook1 = workbook;
                    sheetName1 = workbook.SheetNames[0]; // Set initial sheet name
                } else {
                    workbook2 = workbook;
                    sheetName2 = workbook.SheetNames[0]; // Set initial sheet name
                }
                
                // Show information about sheets if there are multiple
                if (workbook.SheetNames.length > 1) {
                    // Multiple sheets detected - using first sheet
                }
                
                // Update UI to show sheet information
                updateSheetInfo(file.name, workbook.SheetNames, workbook.SheetNames[0], num);
                
                // Populate sheet selector
                populateSheetSelector(workbook.SheetNames, num, workbook.SheetNames[0]);
                
                setTimeout(() => {
                    // Process the first sheet directly (more efficient than calling processSelectedSheet)
                    let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    let json = processExcelSheetOptimized(firstSheet);
                    
                    // Check both file size and column count limits simultaneously
                    let maxCols = 0;
                    for (let i = 0; i < json.length; i++) {
                        if (json[i] && json[i].length > maxCols) {
                            maxCols = json[i].length;
                        }
                    }
                    const rowsExceeded = json.length > MAX_ROWS_LIMIT;
                    const colsExceeded = maxCols > MAX_COLS_LIMIT;
                    
                    if (rowsExceeded || colsExceeded) {
                        tableElement.innerHTML = generateLimitErrorMessage(
                            'rows', json.length, MAX_ROWS_LIMIT, '', 
                            'columns', maxCols, MAX_COLS_LIMIT
                        );
                        // Clear data
                        if (num === 1) {
                            data1 = [];
                            fileName1 = '';
                            workbook1 = null;
                        } else {
                            data2 = [];
                            fileName2 = '';
                            workbook2 = null;
                        }
                        return;
                    }
                    
                    // Show progress for large files
                    if (json.length > 1000) {
                        tableElement.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚙️ Processing large Excel file... Please wait</div>';
                    }
                    
                    setTimeout(() => {
                        // Process the data same as before
                        json = json.filter(row => Array.isArray(row) && row.some(cell => (cell !== null && cell !== undefined && cell.toString().trim() !== '')));
                        
                        // Normalize headers to uppercase
                        json = normalizeHeaders(json);
                        
                        // Apply column-based date conversion (for Excel text dates like "05.01.25")
                        json = processColumnTypes(json);
                        
                        setTimeout(() => {
                            json = removeEmptyColumns(json);
                            json = json.map(row => {
                                if (!Array.isArray(row)) return row;
                                return row.map(cell => roundDecimalNumbers(cell));
                            });
                            
                            if (num === 1) {
                                data1 = json;
                                renderPreview(json, 'table1');
                            } else {
                                data2 = json;
                                renderPreview(json, 'table2');
                            }
                        }, 10);
                    }, 10);
                }, 10);
            }, 10);
        };
        reader.readAsArrayBuffer(file);
    }
}

// Helper function to generate limit exceeded error message
function generateLimitErrorMessage(type, current, limit, additionalInfo = '', secondType = null, secondCurrent = null, secondLimit = null) {
    if (secondType && secondCurrent !== null && secondLimit !== null) {
        // Combined limits error message
        const rowsInfo = type === 'rows' ? { current, limit } : { current: secondCurrent, limit: secondLimit };
        const colsInfo = type === 'columns' ? { current, limit } : { current: secondCurrent, limit: secondLimit };
        
        let violatedLimits = [];
        if (rowsInfo.current > rowsInfo.limit) {
            violatedLimits.push(`<strong>${rowsInfo.current.toLocaleString()} rows</strong> (limit: ${rowsInfo.limit.toLocaleString()})`);
        }
        if (colsInfo.current > colsInfo.limit) {
            violatedLimits.push(`<strong>${colsInfo.current} columns</strong> (limit: ${colsInfo.limit})`);
        }
        
        const title = violatedLimits.length > 1 ? 'File Size and Column Limits Exceeded' : 
                     (rowsInfo.current > rowsInfo.limit ? 'File Size Limit Exceeded' : 'Too Many Columns');
        
        return `
            <div style="text-align: center; padding: 40px; background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px; margin: 20px 0;">
                <div style="font-size: 48px; margin-bottom: 16px;">⚠️</div>
                <div style="font-size: 18px; font-weight: 600; color: #856404; margin-bottom: 10px;">${title}</div>
                <div style="color: #856404; margin-bottom: 15px;">
                    ${additionalInfo ? additionalInfo + ' contains' : 'This file contains'}:<br>
                    ${violatedLimits.join('<br>')}
                </div>
                <div style="color: #856404; font-size: 14px;">
                    Please use a smaller file with fewer rows and columns, or contact support for enterprise solutions.
                </div>
            </div>
        `;
    }
    
    // Single limit error message (original functionality)
    const isRows = type === 'rows';
    const title = isRows ? 'File Size Limit Exceeded' : 'Too Many Columns';
    const currentText = isRows ? `${current.toLocaleString()} rows` : `${current} columns`;
    const limitText = isRows ? `${limit.toLocaleString()} rows` : `${limit} columns`;
    const suggestion = isRows 
        ? 'Please use a smaller file or contact support for enterprise solutions.'
        : 'Please reduce the number of columns or contact support for enterprise solutions.';
    
    return `
        <div style="text-align: center; padding: 40px; background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px; margin: 20px 0;">
            <div style="font-size: 28px; margin-bottom: 16px;">⚠️</div>
            <div style="font-size: 18px; font-weight: 600; color: #856404; margin-bottom: 10px;">${title}</div>
            <div style="color: #856404; margin-bottom: 15px;">
                ${additionalInfo ? additionalInfo + ' contains' : 'This file contains'} <strong>${currentText}</strong>, but the limit is <strong>${limitText}</strong>.
            </div>
            <div style="color: #856404; font-size: 14px;">
                ${suggestion}
            </div>
        </div>
    `;
}

// Function to create column mapping between two headers
function createColumnMapping(header1, header2) {
    const mapping = [];
    const header1Lower = header1.map(h => (h || '').toString().toLowerCase().trim());
    const header2Lower = header2.map(h => (h || '').toString().toLowerCase().trim());
    
    // Find common columns and their positions
    const commonColumns = [];
    const used2 = new Set();
    
    header1Lower.forEach((col1, index1) => {
        if (col1 === '') return; // Skip empty headers
        
        const index2 = header2Lower.findIndex((col2, idx) => 
            col2 === col1 && !used2.has(idx)
        );
        
        if (index2 !== -1) {
            commonColumns.push({
                name: header1[index1],
                index1: index1,
                index2: index2
            });
            used2.add(index2);
        }
    });
    
    return {
        commonColumns,
        onlyInFile1: header1.filter((col, idx) => 
            header1Lower[idx] !== '' && !commonColumns.some(c => c.index1 === idx)
        ),
        onlyInFile2: header2.filter((col, idx) => 
            header2Lower[idx] !== '' && !commonColumns.some(c => c.index2 === idx)
        )
    };
}

// Function to reorder data based on column mapping
function reorderDataByColumns(data, originalHeader, commonColumns, targetOrder) {
    if (!data || data.length === 0) return data;
    
    // Create new header based on target order
    const newHeader = targetOrder.map(colInfo => colInfo.name);
    
    // Reorder all data rows
    const reorderedData = [newHeader];
    
    for (let i = 1; i < data.length; i++) {
        const newRow = targetOrder.map(colInfo => {
            const sourceIndex = colInfo.sourceIndex;
            return sourceIndex < data[i].length ? data[i][sourceIndex] : '';
        });
        reorderedData.push(newRow);
    }
    
    return reorderedData;
}

// Function to prepare data for comparison with column alignment
function prepareDataForComparison(data1, data2) {
    if (!data1.length || !data2.length) {
        return { data1, data2, columnInfo: null };
    }
    
    const header1 = data1[0] || [];
    const header2 = data2[0] || [];
    
    // Create column mapping
    const mapping = createColumnMapping(header1, header2);
    
    if (mapping.commonColumns.length === 0) {
        // No common columns found - return original data with column differences
        return { 
            data1, 
            data2, 
            columnInfo: {
                hasCommonColumns: false,
                onlyInFile1: mapping.onlyInFile1,
                onlyInFile2: mapping.onlyInFile2,
                message: "No common column names found. Comparison will be done by position."
            }
        };
    }
    
    // Create unified column order based on common columns
    const unifiedOrder = mapping.commonColumns.map(col => ({
        name: col.name,
        sourceIndex: col.index1 // Use file1 as reference
    }));
    
    // Reorder data2 to match data1 column order
    const reorderedData1 = reorderDataByColumns(data1, header1, mapping.commonColumns, unifiedOrder);
    
    // For data2, we need to map the source indices correctly
    const unifiedOrderForData2 = mapping.commonColumns.map(col => ({
        name: col.name,
        sourceIndex: col.index2 // Use file2 indices
    }));
    
    const reorderedData2 = reorderDataByColumns(data2, header2, mapping.commonColumns, unifiedOrderForData2);
    
    return {
        data1: reorderedData1,
        data2: reorderedData2,
        columnInfo: {
            hasCommonColumns: true,
            commonCount: mapping.commonColumns.length,
            onlyInFile1: mapping.onlyInFile1,
            onlyInFile2: mapping.onlyInFile2,
            reordered: header1.join(',').toLowerCase() !== header2.join(',').toLowerCase()
        }
    };
}

// Function to normalize headers to uppercase
function normalizeHeaders(data) {
    if (!data || data.length === 0) return data;
    
    // Convert first row (headers) to uppercase
    if (data[0] && Array.isArray(data[0])) {
        data[0] = data[0].map(header => {
            if (header && typeof header === 'string') {
                return header.toUpperCase();
            }
            return header;
        });
    }
    
    return data;
}

function renderTable(data, elementId, diffRowsSet) {
    let html = '<table>';
    if (data.length > 0) {
        html += '<tr>';
        for (let j = 0; j < data[0].length; j++) {
            html += `<th>${data[0][j]}</th>`;
        }
        html += '</tr>';
    }
    for (let i = 1; i < data.length; i++) { // start from 1 to avoid duplicating headers
        let rowClass = (diffRowsSet && diffRowsSet.has(i)) ? 'diff' : '';
        html += `<tr class="${rowClass}">`;
        for (let j = 0; j < data[i].length; j++) {
            html += `<td>${data[i][j]}</td>`;
        }
        html += '</tr>';
    }
    html += '</table>';
    document.getElementById(elementId).innerHTML = html;
}

function renderPreview(data, elementId, title) {
    let html = '';
    if (data.length > 0) {
        html += '<div class="preview-wrapper">';
        html += '<table class="preview-table">';
        
        // Headers
        if (data[0] && data[0].length > 0) {
            html += '<tr>';
            for (let j = 0; j < data[0].length; j++) {
                let headerValue = data[0][j];
                headerValue = (headerValue !== null && headerValue !== undefined) ? headerValue.toString() : '';
                let displayValue = headerValue.length > 20 ? headerValue.substring(0, 20) + '...' : headerValue;
                let titleAttr = headerValue.length > 20 ? ` title="${headerValue.replace(/"/g, '&quot;')}"` : '';
                html += `<th${titleAttr}>${displayValue}</th>`;
            }
            html += '</tr>';
        }
        
        // Show only first 3 data rows for performance
        const previewRows = Math.min(4, data.length);
        for (let i = 1; i < previewRows; i++) {
            html += '<tr>';
            for (let j = 0; j < data[i].length; j++) {
                let cellValue = data[i][j];
                cellValue = (cellValue !== null && cellValue !== undefined) ? cellValue.toString() : '';
                let displayValue = cellValue.length > 25 ? cellValue.substring(0, 25) + '...' : cellValue;
                let titleAttr = cellValue.length > 25 ? ` title="${cellValue.replace(/"/g, '&quot;')}"` : '';
                html += `<td${titleAttr}>${displayValue}</td>`;
            }
            html += '</tr>';
        }
        
        // Add info about total rows if file is large
        if (data.length > 4) {
            html += `<tr><td colspan="${data[0].length}" style="text-align: center; font-style: italic; color: #666;">
                ... and ${data.length - 1} more rows (showing first 3 for preview)
            </td></tr>`;
        }
        
        html += '</table>';
        html += '</div>';
    }
    document.getElementById(elementId).innerHTML = html;
}

function rowToKey(row) {
    // Convert row to key string for comparison
    return JSON.stringify(row);
}

function showPlaceholderMessage() {
    document.getElementById('result').innerHTML = '';
    document.getElementById('summary').innerHTML = '';
    
    // Get the translated text from HTML data attributes or use default English
    const placeholderText = document.documentElement.getAttribute('data-placeholder-text') || 'Choose files on your computer and click Compare';
    
    document.getElementById('diffTable').innerHTML = `
        <div class="placeholder-message">
            <div class="placeholder-icon">📊</div>
            <div class="placeholder-text">${placeholderText}</div>
        </div>
    `;
    
    // Hide filter controls when showing placeholder
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
}

// Centralized function to clear comparison results and UI state
function clearComparisonResults() {
    // Clear result containers
    document.getElementById('result').innerHTML = '';
    document.getElementById('summary').innerHTML = '';
    document.getElementById('diffTable').innerHTML = '';
    
    // Hide filter controls
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
    
    // Hide export button
    const exportBtn = document.getElementById('exportExcelBtn');
    const buttonsContainer = document.querySelector('.buttons-container');
    const exportButtonHalf = exportBtn ? exportBtn.closest('.button-half') : null;
    
    if (exportBtn && buttonsContainer) {
        exportBtn.style.display = 'none';
        if (exportButtonHalf) {
            exportButtonHalf.classList.add('export-hidden');
        }
        buttonsContainer.classList.add('export-hidden');
    }
    
    // Reset global comparison state
    currentPairs = [];
    currentSortColumn = -1;
    currentSortDirection = 'asc';
    currentDiffColumns1 = '-';
    currentDiffColumns2 = '-';
}

function restoreTableStructure() {
    document.getElementById('diffTable').innerHTML = `
        <div class="table-container-sync">
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
        </div>
    `;
}

// Helper function to get excluded columns from the input field
function getExcludedColumns() {
    const excludeInput = document.getElementById('excludeColumns');
    if (!excludeInput || !excludeInput.value.trim()) {
        return [];
    }
    return excludeInput.value.trim().split(',').map(col => col.trim()).filter(col => col.length > 0);
}

// Helper function to get column indexes to exclude
function getExcludedColumnIndexes(headers, excludedColumns) {
    if (!excludedColumns.length) return [];
    const excludedIndexes = [];
    excludedColumns.forEach(excludedCol => {
        const index = headers.findIndex(header => 
            header && header.toString().trim().toLowerCase() === excludedCol.toLowerCase()
        );
        if (index !== -1) {
            excludedIndexes.push(index);
        }
    });
    return excludedIndexes;
}

// Helper function to filter row data excluding specified columns
function filterRowExcludingColumns(row, excludedIndexes) {
    if (!excludedIndexes.length) return row;
    return row.filter((_, index) => !excludedIndexes.includes(index));
}

// Helper function to identify different column types
function getColumnTypes(header1, header2) {
    const header1Set = new Set(header1.map(h => h ? h.toString().trim().toLowerCase() : ''));
    const header2Set = new Set(header2.map(h => h ? h.toString().trim().toLowerCase() : ''));
    
    const columnTypes = {};
    
    // Mark columns by type
    header1.forEach((header, index) => {
        const headerKey = header ? header.toString().trim().toLowerCase() : '';
        if (headerKey) {
            if (header2Set.has(headerKey)) {
                columnTypes[index] = 'common'; // Column exists in both files
            } else {
                columnTypes[index] = 'diff'; // Column only in file 1
            }
        }
    });
    
    header2.forEach((header, index) => {
        const headerKey = header ? header.toString().trim().toLowerCase() : '';
        if (headerKey && !header1Set.has(headerKey)) {
            // Find the actual index in the combined headers array
            const combinedIndex = header1.length > header2.length ? 
                header1.findIndex(h => !h || h.toString().trim() === '') : 
                header1.length + index;
            columnTypes[combinedIndex] = 'new'; // Column only in file 2
        }
    });
    
    return columnTypes;
}

// Helper function to filter columns based on hide options
function getColumnsToHide(headers, columnTypes, hideDiff, hideNew) {
    const columnsToHide = [];
    
    headers.forEach((header, index) => {
        const columnType = columnTypes[index];
        if ((hideDiff && columnType === 'diff') || (hideNew && columnType === 'new')) {
            columnsToHide.push(index);
        }
    });
    
    return columnsToHide;
}

function compareTables(useTolerance = false) {
    // Set tolerance mode
    toleranceMode = useTolerance;
    
    // Clear previous results immediately when starting new comparison
    clearComparisonResults();
    
    // Show immediate loading indicator
    let resultDiv = document.getElementById('result');
    let summaryDiv = document.getElementById('summary');
    
    resultDiv.innerHTML = '<div id="comparison-loading" style="text-align: center; padding: 20px; font-size: 16px;">🔄 Starting comparison... Please wait</div>';
    summaryDiv.innerHTML = '<div style="text-align: center; padding: 10px;">Initializing...</div>';
    
    if (!data1.length || !data2.length) {
        document.getElementById('result').innerText = 'Please, load both files.';
        document.getElementById('summary').innerHTML = '';
        showPlaceholderMessage();
        return;
    }
    
    // Check combined size limit before comparison
    const totalRows = Math.max(data1.length, data2.length);
    if (totalRows > MAX_ROWS_LIMIT) {
        document.getElementById('result').innerHTML = `
            <div style="text-align: center; padding: 40px; background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px; margin: 20px 0;">
                <div style="font-size: 28px; margin-bottom: 16px;">⚠️</div>
                <div style="font-size: 18px; font-weight: 600; color: #856404; margin-bottom: 10px;">File Size Limit Exceeded</div>
                <div style="color: #856404; margin-bottom: 15px;">
                    Cannot compare files with more than <strong>${MAX_ROWS_LIMIT.toLocaleString()}</strong> rows.<br>
                    Current files have <strong>${data1.length.toLocaleString()}</strong> and <strong>${data2.length.toLocaleString()}</strong> rows.
                </div>
                <div style="color: #856404; font-size: 14px;">
                    Please use smaller files or contact support for enterprise solutions.
                </div>
            </div>
        `;
        document.getElementById('summary').innerHTML = '';
        return;
    }
    
    // Check column count limit before comparison
    let maxCols1 = 0;
    for (let i = 0; i < data1.length; i++) {
        if (data1[i] && data1[i].length > maxCols1) {
            maxCols1 = data1[i].length;
        }
    }
    let maxCols2 = 0;
    for (let i = 0; i < data2.length; i++) {
        if (data2[i] && data2[i].length > maxCols2) {
            maxCols2 = data2[i].length;
        }
    }
    const totalCols = Math.max(maxCols1, maxCols2);
    if (totalCols > MAX_COLS_LIMIT) {
        document.getElementById('result').innerHTML = `
            <div style="text-align: center; padding: 40px; background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px; margin: 20px 0;">
                <div style="font-size: 28px; margin-bottom: 16px;">⚠️</div>
                <div style="font-size: 18px; font-weight: 600; color: #856404; margin-bottom: 10px;">Too Many Columns</div>
                <div style="color: #856404; margin-bottom: 15px;">
                    Cannot compare files with more than <strong>${MAX_COLS_LIMIT}</strong> columns.<br>
                    Current files have <strong>${maxCols1}</strong> and <strong>${maxCols2}</strong> columns.
                </div>
                <div style="color: #856404; font-size: 14px;">
                    Please reduce the number of columns or contact support for enterprise solutions.
                </div>
            </div>
        `;
        document.getElementById('summary').innerHTML = '';
        return;
    }
    
    // Show loading indicator for large files
    if (totalRows > 1000) {
        const fileInfo = `${data1.length.toLocaleString()} vs ${data2.length.toLocaleString()} rows`;
        const loadingMessage = totalRows > DETAILED_TABLE_LIMIT ? 
            `🔄 Processing large files (${fileInfo}) - Summary mode... Please wait` :
            `🔄 Comparing files (${fileInfo})... Please wait`;
        
        document.getElementById('result').innerHTML = `<div style="text-align: center; padding: 20px; font-size: 16px;">${loadingMessage}</div>`;
        document.getElementById('summary').innerHTML = '<div style="text-align: center; padding: 10px;">Analyzing data...</div>';
    }
    
    // Use setTimeout to allow UI to update
    setTimeout(() => {
        performComparison();
    }, 10);
}

function performComparison() {
    // Get translated table headers
    const tableHeaders = getSummaryTableHeaders();
    
    // Restore table structure for comparison results
    restoreTableStructure();
    
    // Automatically enable "Hide COMMA rows" (hide identical rows) after comparison
    const hideSameRowsCheckbox = document.getElementById('hideSameRows');
    if (hideSameRowsCheckbox) {
        hideSameRowsCheckbox.checked = true;
    }
    
    // Get excluded columns
    const excludedColumns = getExcludedColumns();
    
    // Get filter options
    const hideDiffRowsEl = document.getElementById('hideDiffColumns');
    const hideNewRows1El = document.getElementById('hideNewRows1');
    const hideNewRows2El = document.getElementById('hideNewRows2');
    
    const hideDiffRows = hideDiffRowsEl ? hideDiffRowsEl.checked : false; // Hide rows with differences
    const hideNewRows1 = hideNewRows1El ? hideNewRows1El.checked : false; // Hide rows only in file 1
    const hideNewRows2 = hideNewRows2El ? hideNewRows2El.checked : false; // Hide rows only in file 2
    
    // Prepare data for comparison with column alignment
    const { data1: alignedData1, data2: alignedData2, columnInfo } = prepareDataForComparison(data1, data2);
    
    // Show column alignment info to user
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
    
    // Use aligned data for comparison
    const workingData1 = alignedData1;
    const workingData2 = alignedData2;
    
    // --- Statistics ---
    let header1 = workingData1[0] || [];
    let header2 = workingData2[0] || [];
    
    // Get excluded column indexes
    const excludedIndexes1 = getExcludedColumnIndexes(header1, excludedColumns);
    const excludedIndexes2 = getExcludedColumnIndexes(header2, excludedColumns);
    
    // Filter headers to exclude specified columns first
    const filteredHeader1 = filterRowExcludingColumns(header1, excludedIndexes1);
    const filteredHeader2 = filterRowExcludingColumns(header2, excludedIndexes2);
    
    // Determine which headers to use for column analysis
    let allCols = Math.max(filteredHeader1.length, filteredHeader2.length);
    let headers = (filteredHeader1.length >= filteredHeader2.length) ? filteredHeader1 : filteredHeader2;
    
    // Use original filtered headers without additional column hiding
    const finalHeaders = headers;
    const finalAllCols = finalHeaders.length;
    
    // Adjust body data with only exclusion filters applied (no column hiding)
    let body1 = workingData1.slice(1).map(row => filterRowExcludingColumns(row, excludedIndexes1));
    let body2 = workingData2.slice(1).map(row => filterRowExcludingColumns(row, excludedIndexes2));
    
    // Find columns that exist in one file but not in the other
    // Use columnInfo if available (from prepareDataForComparison), otherwise calculate manually
    let onlyInFile1, onlyInFile2;
    
    if (columnInfo && (columnInfo.hasCommonColumns || columnInfo.onlyInFile1 || columnInfo.onlyInFile2)) {
        // Use pre-calculated column differences from prepareDataForComparison
        onlyInFile1 = columnInfo.onlyInFile1 || [];
        onlyInFile2 = columnInfo.onlyInFile2 || [];
    } else {
        // Calculate manually for cases where no column alignment was done
        let header1Set = new Set(filteredHeader1.map(h => h ? h.toString().trim().toLowerCase() : ''));
        let header2Set = new Set(filteredHeader2.map(h => h ? h.toString().trim().toLowerCase() : ''));
        
        onlyInFile1 = filteredHeader1.filter(h => h && h.toString().trim() !== '' && !header2Set.has(h.toString().trim().toLowerCase()));
        onlyInFile2 = filteredHeader2.filter(h => h && h.toString().trim() !== '' && !header1Set.has(h.toString().trim().toLowerCase()));
    }
    
    let diffColumns1 = onlyInFile1.length > 0 ? onlyInFile1.join(', ') : '-';
    let diffColumns2 = onlyInFile2.length > 0 ? onlyInFile2.join(', ') : '-';
    
    // Save diff columns info globally for later use
    currentDiffColumns1 = diffColumns1;
    currentDiffColumns2 = diffColumns2;
    
    // Apply red highlighting if there are diff columns
    let diffColumns1Html = onlyInFile1.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns1}</span>` : diffColumns1;
    let diffColumns2Html = onlyInFile2.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns2}</span>` : diffColumns2;
    
    // Form keys for exact match with case-insensitive comparison
    // Clear the loading message if it exists
    const loadingDiv = document.getElementById('comparison-loading');
    if (loadingDiv) {
        loadingDiv.remove();
    }

    function rowKey(row) { 
        if (toleranceMode) {
            // In tolerance mode, we cannot use simple string comparison for stats
            // We'll need to compare rows individually in the matching logic
            return JSON.stringify(row.map(x => (x !== undefined ? x.toString() : ''))); 
        } else {
            return JSON.stringify(row.map(x => (x !== undefined ? x.toString().toUpperCase() : ''))); 
        }
    }
    
    let set1, set2, only1, only2, both;
    
    if (toleranceMode) {
        // In tolerance mode, we need to compare each row individually
        // Show temporary statistics until actual matching is completed
        set1 = new Set(body1.map((row, index) => `row1_${index}`));
        set2 = new Set(body2.map((row, index) => `row2_${index}`));
        only1 = 0; // Will be calculated after matching
        only2 = 0; // Will be calculated after matching
        both = 0; // Will be calculated after matching
    } else {
        set1 = new Set(body1.map(rowKey));
        set2 = new Set(body2.map(rowKey));
        only1 = 0; only2 = 0; both = 0;
        set1.forEach(k => { if (set2.has(k)) both++; else only1++; });
        set2.forEach(k => { if (!set1.has(k)) only2++; });
    }
    // --- Top table ---
    let totalRows = Math.max(data1.length, data2.length) - 1;
    let maxRows = Math.max(body1.length, body2.length);
    
    // Calculate percentage similarity: identical rows / max file size
    // This shows how similar the files are (higher % = more similar)
    let totalUniqueRows = only1 + only2 + both;
    let differentRows = only1 + only2;
    let percentDiff, percentClass;
    
    if (toleranceMode && totalUniqueRows === 0) {
        // In tolerance mode, show temporary message until matching is complete
        percentDiff = tableHeaders.calculating;
        percentClass = 'percent-low';
    } else {
        // Use max file size as denominator: (identical rows / max file size) * 100
        const maxFileSize = Math.max(body1.length, body2.length);
        percentDiff = maxFileSize > 0 ? Math.min(((both / maxFileSize) * 100), 100).toFixed(2) + '%' : '0.00%';
        percentClass = 'percent-high';
        if (parseFloat(percentDiff) < 30) percentClass = 'percent-low';
        else if (parseFloat(percentDiff) < 70) percentClass = 'percent-medium';
        else percentClass = 'percent-high';
    }
    
    // Add excluded columns info
    let excludedInfo = '';
    if (excludedColumns.length > 0) {
        excludedInfo = `<div class="excluded-info">
            <strong>Excluded from comparison:</strong> ${excludedColumns.join(', ')}
        </div>`;
    }
    
    // Add tolerance mode info
    let toleranceInfo = '';
    if (toleranceMode) {
        // Detect current language based on page URL or title
        const currentLang = window.location.pathname.includes('/ru/') ? 'ru' : 
                           window.location.pathname.includes('/pl/') ? 'pl' :
                           window.location.pathname.includes('/es/') ? 'es' :
                           window.location.pathname.includes('/de/') ? 'de' :
                           window.location.pathname.includes('/ja/') ? 'ja' :
                           window.location.pathname.includes('/pt/') ? 'pt' :
                           window.location.pathname.includes('/zh/') ? 'zh' :
                           window.location.pathname.includes('/ar/') ? 'ar' : 'en';
                           
        const toleranceMessages = {
            'ru': 'Режим погрешности активен: Числа с разницей в 1,5% и даты с одинаковой датой (разное время) показаны',
            'pl': 'Tryb tolerancji aktywny: Liczby różniące się o 1,5% i daty z tą samą datą (różny czas) są pokazane',
            'es': 'Modo de tolerancia activo: Números con diferencia del 1,5% y fechas con la misma fecha (diferente hora) se muestran',
            'de': 'Toleranzmodus aktiv: Zahlen mit 1,5% Unterschied und Daten mit gleichem Datum (unterschiedliche Zeit) werden angezeigt',
            'ja': '許容差モードがアクティブ: 1,5%の差がある数値と同じ日付（異なる時刻）の日付が表示されます',
            'pt': 'Modo de tolerância ativo: Números com diferença de 1,5% e datas com a mesma data (horário diferente) são mostrados',
            'zh': '容差模式激活：1,5%差异的数字和相同日期（不同时间）的日期显示为',
            'ar': 'وضع التسامح نشط: الأرقام التي تختلف بنسبة 1,5% والتواريخ بنفس التاريخ (وقت مختلف) تظهر',
            'en': 'Tolerance Mode Active: Numbers within 1.5% difference and dates with same date (different time) are shown'
        };
        
        toleranceInfo = `<div style="background: #fff3cd; border: 1px solid #ffeaa7; color: #856404; padding: 12px; margin: 15px 0; border-radius: 6px; font-size: 14px;">
            <strong>🔄 ${toleranceMessages[currentLang]}</strong> <span style="background: #ffeaa7; padding: 2px 4px; border-radius: 3px;">${currentLang === 'ru' ? 'оранжевым' : currentLang === 'pl' ? 'na pomarańczowo' : currentLang === 'es' ? 'en naranja' : currentLang === 'de' ? 'in Orange' : currentLang === 'ja' ? 'オレンジ色で' : currentLang === 'pt' ? 'em laranja' : currentLang === 'zh' ? '橙色' : currentLang === 'ar' ? 'باللون البرتقالي' : 'in orange'}</span> ${currentLang === 'ru' ? 'вместо красного' : currentLang === 'pl' ? 'zamiast na czerwono' : currentLang === 'es' ? 'en lugar de rojo' : currentLang === 'de' ? 'statt in Rot' : currentLang === 'ja' ? '赤の代わりに' : currentLang === 'pt' ? 'em vez de vermelho' : currentLang === 'zh' ? '而不是红色' : currentLang === 'ar' ? 'بدلاً من الأحمر' : 'instead of red'}.
        </div>`;
    }
    
    let htmlSummary = `
        ${excludedInfo}
        ${toleranceInfo}
        <table style="margin-bottom:20px; border: 1px solid #ccc;">
            <tr><th>${tableHeaders.file}</th><th>${tableHeaders.rowCount}</th><th>${tableHeaders.rowsOnlyInFile}</th><th>${tableHeaders.identicalRows}</th><th>${tableHeaders.similarity}</th><th>${tableHeaders.diffColumns}</th></tr>
            <tr><td>${fileName1 || tableHeaders.file1}</td><td>${body1.length}</td><td>${toleranceMode && totalUniqueRows === 0 ? tableHeaders.calculating : only1}</td><td rowspan="2">${toleranceMode && totalUniqueRows === 0 ? tableHeaders.calculating : both}</td><td rowspan="2" class="percent-cell ${percentClass}">${percentDiff}</td><td>${diffColumns1Html}</td></tr>
            <tr><td>${fileName2 || tableHeaders.file2}</td><td>${body2.length}</td><td>${toleranceMode && totalUniqueRows === 0 ? tableHeaders.calculating : only2}</td><td>${diffColumns2Html}</td></tr>
        </table>
    `;
    document.getElementById('summary').innerHTML = htmlSummary;
    
    // Show filter controls after comparison
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'block';
    }
    
    // Show export button after successful comparison
    const exportBtn = document.getElementById('exportExcelBtn');
    const buttonsContainer = document.querySelector('.buttons-container');
    const exportButtonHalf = exportBtn ? exportBtn.closest('.button-half') : null;
    
    if (exportBtn && buttonsContainer) {
        exportBtn.style.display = 'inline-block';
        // Remove classes that hide export button container
        if (exportButtonHalf) {
            exportButtonHalf.classList.remove('export-hidden');
        }
        buttonsContainer.classList.remove('export-hidden');
    }
    
    // Don't automatically check "Hide same rows" - let user decide
    if (hideSameRowsCheckbox) {
        // Reset checkbox state and render all rows initially
        hideSameRowsCheckbox.checked = false;
    }
    
    // Check if files are too large for detailed table rendering
    const totalRowsForCheck = Math.max(body1.length, body2.length);
    const isLargeFile = totalRowsForCheck > DETAILED_TABLE_LIMIT;
    
    if (isLargeFile) {
        // For large files (15K-45K rows), perform fuzzy matching for export but don't show table
        performFuzzyMatchingForExport(body1, body2, finalHeaders, finalAllCols, true, tableHeaders);
        return;
    } else {
        // For smaller files, continue with detailed comparison table and full processing
        performFuzzyMatchingForExport(body1, body2, finalHeaders, finalAllCols, false, tableHeaders);
    }
}

// Function to update summary statistics based on actual matching results
function updateSummaryStatistics(tableHeaders, file1Size = null, file2Size = null) {
    if (!currentPairs || currentPairs.length === 0) return;
    
    let identicalCount = 0;
    let onlyInFile1 = 0;
    let onlyInFile2 = 0;
    let toleranceCount = 0;
    let differentCount = 0;
    
    currentPairs.forEach(pair => {
        const row1 = pair.row1;
        const row2 = pair.row2;
        
        if (row1 && row2) {
            // Both files have data - check if identical or different
            let isIdentical = true;
            let hasTolerance = false;
            
            for (let c = 0; c < currentFinalAllCols; c++) {
                const v1 = row1[c] !== undefined ? row1[c] : '';
                const v2 = row2[c] !== undefined ? row2[c] : '';
                
                if (toleranceMode) {
                    const compResult = compareValuesWithTolerance(v1, v2);
                    if (compResult === 'different') {
                        isIdentical = false;
                        break;
                    } else if (compResult === 'tolerance') {
                        hasTolerance = true;
                        isIdentical = false;
                    }
                } else {
                    if (v1.toString().toUpperCase() !== v2.toString().toUpperCase()) {
                        isIdentical = false;
                        break;
                    }
                }
            }
            
            if (isIdentical && !hasTolerance) {
                identicalCount++;
            } else if (hasTolerance && toleranceMode) {
                toleranceCount++;
            } else {
                differentCount++;
            }
        } else if (row1 && !row2) {
            onlyInFile1++;
        } else if (!row1 && row2) {
            onlyInFile2++;
        }
    });
    
    // Update the global variables used by the summary table
    if (toleranceMode) {
        // In tolerance mode, identical includes both exact matches and tolerance matches
        window.summaryStats = {
            only1: onlyInFile1,
            only2: onlyInFile2,
            both: identicalCount + toleranceCount, // Both exact and tolerance matches count as "both"
            exact: identicalCount,
            tolerance: toleranceCount,
            different: differentCount
        };
    } else {
        window.summaryStats = {
            only1: onlyInFile1,
            only2: onlyInFile2,
            both: identicalCount,
            different: differentCount
        };
    }
    
    // Recalculate percentage based on similarity (identical rows / max file size)
    // Use the maximum number of rows between the two files as the denominator
    const maxFileSize = file1Size && file2Size ? Math.max(file1Size, file2Size) : (window.summaryStats.only1 + window.summaryStats.only2 + window.summaryStats.both);
    const percentDiff = maxFileSize > 0 ? Math.min(((window.summaryStats.both / maxFileSize) * 100), 100).toFixed(2) : '0.00';
    
    let percentClass = 'percent-high';
    if (parseFloat(percentDiff) < 30) percentClass = 'percent-low';
    else if (parseFloat(percentDiff) < 70) percentClass = 'percent-medium';
    else percentClass = 'percent-high';
    
    // Update the summary table HTML
    updateSummaryTable(window.summaryStats.only1, window.summaryStats.only2, window.summaryStats.both, percentDiff, percentClass, tableHeaders);
}

// Function to update the summary table HTML
function updateSummaryTable(only1, only2, both, percentDiff, percentClass, tableHeaders) {
    // Get existing excluded info and tolerance info
    const summaryDiv = document.getElementById('summary');
    const existingHTML = summaryDiv.innerHTML;
    
    // Extract excluded info and tolerance info if they exist
    const excludedMatch = existingHTML.match(/<div class="excluded-info">.*?<\/div>/);
    const toleranceMatch = existingHTML.match(/<div style="background: #fff3cd;.*?<\/div>/);
    
    const excludedInfo = excludedMatch ? excludedMatch[0] : '';
    const toleranceInfo = toleranceMatch ? toleranceMatch[0] : '';
    
    // Get diff columns info from saved global values
    const diffColumns1 = currentDiffColumns1 || '-';
    const diffColumns2 = currentDiffColumns2 || '-';
    
    // Apply red highlighting if there are diff columns (only if not just '-')
    const diffColumns1Html = (diffColumns1 !== '-') ? `<span style="color: red; font-weight: bold;">${diffColumns1}</span>` : diffColumns1;
    const diffColumns2Html = (diffColumns2 !== '-') ? `<span style="color: red; font-weight: bold;">${diffColumns2}</span>` : diffColumns2;
    
    let htmlSummary = `
        ${excludedInfo}
        ${toleranceInfo}
        <table style="margin-bottom:20px; border: 1px solid #ccc;">
            <tr><th>${tableHeaders.file}</th><th>${tableHeaders.rowCount}</th><th>${tableHeaders.rowsOnlyInFile}</th><th>${tableHeaders.identicalRows}</th><th>${tableHeaders.similarity}</th><th>${tableHeaders.diffColumns}</th></tr>
            <tr><td>${getFileDisplayName(fileName1, sheetName1) || tableHeaders.file1}</td><td>${currentPairs.filter(p => p.row1).length}</td><td>${only1}</td><td rowspan="2">${both}</td><td rowspan="2" class="percent-cell ${percentClass}">${percentDiff}%</td><td>${diffColumns1Html}</td></tr>
            <tr><td>${getFileDisplayName(fileName2, sheetName2) || tableHeaders.file2}</td><td>${currentPairs.filter(p => p.row2).length}</td><td>${only2}</td><td>${diffColumns2Html}</td></tr>
        </table>
    `;
    
    summaryDiv.innerHTML = htmlSummary;
}

// Smart detection of key columns based on headers, uniqueness, and position
function smartDetectKeyColumns(headers, data) {
    if (!headers || !data || data.length < 2) {
        return [0]; // Default to first column if no data
    }
    
    const columnCount = headers.length;
    const bodyData = data.slice(1); // Skip header row
    
    // Define key indicators for headers
    const keyIndicators = {
        high: ['id', 'uid', 'key', 'primary', 'identifier', 'код', 'номер', 'артикул', 'pk', 'primarykey'],
        medium: ['name', 'title', 'label', 'имя', 'название', 'наименование', 'фио', 'customer', 'client', 'клиент'],
        low: ['date', 'time', 'created', 'modified', 'дата', 'время', 'создан', 'изменен'],
        // New: Aggregation patterns for YYYYMM and multi-column aggregation
        aggregation: ['yyyymm', 'yyyymmdd', 'year_month', 'yearmonth', 'period', 'период', 'reporting_period', 'отчетный_период',
                     'year', 'год', 'month', 'месяц', 'quarter', 'квартал', 'partition', 'раздел']
    };
    
    const columnScores = [];
    
    for (let colIndex = 0; colIndex < columnCount; colIndex++) {
        const header = (headers[colIndex] || '').toString().toLowerCase();
        let score = 0;
        
        // 1. Header analysis (40% weight)
        let headerScore = 0;
        
        // Check for aggregation patterns (highest priority for YYYYMM type fields)
        if (keyIndicators.aggregation.some(keyword => header.includes(keyword))) {
            // Special boost for YYYYMM pattern
            if (header.match(/^y{4}m{2}$|yyyymm|year.*month|месяц.*год|period|период/i)) {
                headerScore = 12; // Highest priority for temporal aggregation
            } else {
                headerScore = 9; // High priority for other aggregation components
            }
        }
        // Check for high priority keywords
        else if (keyIndicators.high.some(keyword => header.includes(keyword))) {
            headerScore = 10;
        }
        // Check for medium priority keywords
        else if (keyIndicators.medium.some(keyword => header.includes(keyword))) {
            headerScore = 6;
        }
        // Check for low priority keywords
        else if (keyIndicators.low.some(keyword => header.includes(keyword))) {
            headerScore = 3;
        }
        // No keywords found
        else {
            headerScore = 1;
        }
        
        score += headerScore * 0.4;
        
        // 2. Uniqueness analysis (40% weight)
        const columnValues = bodyData.map(row => (row[colIndex] || '').toString().trim()).filter(val => val !== '');
        const uniqueValues = new Set(columnValues);
        const uniquenessRatio = columnValues.length > 0 ? uniqueValues.size / columnValues.length : 0;
        
        // Special analysis for YYYYMM and aggregation patterns
        let isAggregationField = false;
        let aggregationConfidence = 0;
        
        if (columnValues.length > 0) {
            // Check for YYYYMM pattern (202301, 202302, etc.)
            const yyyymmPattern = /^20\d{4}$/; // Matches 6-digit numbers starting with 20 (2000-2099 years)
            const yyyymmCount = columnValues.filter(val => yyyymmPattern.test(val)).length;
            const hasYYYYMMPattern = yyyymmCount > columnValues.length * 0.7; // 70% match threshold
            
            if (hasYYYYMMPattern) {
                // Additional validation for sequential YYYYMM values
                const yyyymmValues = columnValues.filter(val => yyyymmPattern.test(val))
                    .map(val => parseInt(val))
                    .sort((a, b) => a - b);
                
                let validMonthCount = 0;
                for (const value of yyyymmValues) {
                    const month = value % 100;
                    if (month >= 1 && month <= 12) {
                        validMonthCount++;
                    }
                }
                
                if (validMonthCount > yyyymmValues.length * 0.8) {
                    isAggregationField = true;
                    aggregationConfidence = 0.9;
                }
            }
            
            // Check for other aggregation patterns
            const aggregationPatterns = [
                { pattern: /^20\d{2}$/, type: 'year' },          // 2023, 2024
                { pattern: /^[1-9]$|^1[0-2]$/, type: 'month' },  // 1-12
                { pattern: /^[1-4]$/, type: 'quarter' },         // 1-4
                { pattern: /^20\d{2}-[01]\d$/, type: 'year-month' }, // 2023-01
                { pattern: /^Q[1-4]$/, type: 'quarter' }         // Q1, Q2, Q3, Q4
            ];
            
            for (const { pattern, type } of aggregationPatterns) {
                const matches = columnValues.filter(val => pattern.test(val)).length;
                if (matches > columnValues.length * 0.6) {
                    isAggregationField = true;
                    aggregationConfidence = Math.max(aggregationConfidence, 0.7);
                    break;
                }
            }
        }
        
        let uniquenessScore = 0;
        if (isAggregationField) {
            // Aggregation fields are excellent keys even with lower uniqueness
            const baseScore = uniquenessRatio >= 0.3 ? 9 : 7;
            uniquenessScore = Math.floor(baseScore * aggregationConfidence);
        } else if (uniquenessRatio >= 0.95) {
            uniquenessScore = 10; // Very unique - likely a key
        } else if (uniquenessRatio >= 0.8) {
            uniquenessScore = 8;
        } else if (uniquenessRatio >= 0.6) {
            uniquenessScore = 6;
        } else if (uniquenessRatio >= 0.4) {
            uniquenessScore = 4;
        } else {
            uniquenessScore = 1; // Low uniqueness - unlikely to be a key
        }
        
        score += uniquenessScore * 0.4;
        
        // 3. Position weight (20% weight) - earlier columns are more likely to be keys
        const positionScore = Math.max(1, 10 - colIndex * 2); // Decreases as position increases
        score += positionScore * 0.2;
        
        columnScores.push({
            index: colIndex,
            header: headers[colIndex],
            score: score,
            uniquenessRatio: uniquenessRatio,
            headerScore: headerScore,
            uniquenessScore: uniquenessScore,
            positionScore: positionScore,
            isAggregationField: isAggregationField,
            aggregationConfidence: aggregationConfidence
        });
    }
    
    // Sort by score descending
    columnScores.sort((a, b) => b.score - a.score);
    
    // Log the analysis results (commented out for production)
    // console.log('🔍 Smart Key Column Detection Results:');
    // console.log('Columns analyzed:', columnScores.map(col => ({
    //     column: `"${col.header}"`,
    //     index: col.index,
    //     totalScore: col.score.toFixed(2),
    //     uniqueness: `${(col.uniquenessRatio * 100).toFixed(1)}%`,
    //     headerMatch: col.headerScore,
    //     position: col.positionScore,
    //     aggregation: col.isAggregationField ? `✅ YES (${(col.aggregationConfidence * 100).toFixed(0)}%)` : '❌ NO'
    // })));
    
    // Determine key columns based on scores with priority for aggregation fields
    let keyColumns = [];
    
    // First, look for aggregation fields - they are excellent keys for YYYYMM type data
    const aggregationFields = columnScores.filter(col => col.isAggregationField);
    if (aggregationFields.length > 0) {
        // console.log('🎯 Aggregation fields found:', aggregationFields.map(col => `"${col.header}" (${(col.aggregationConfidence * 100).toFixed(0)}% confidence)`));
        
        // Prioritize aggregation fields, especially YYYYMM patterns
        keyColumns.push(aggregationFields[0].index);
        
        // Add additional aggregation fields if they are high quality
        for (let i = 1; i < Math.min(3, aggregationFields.length); i++) {
            const col = aggregationFields[i];
            if (col.score >= 8 && col.aggregationConfidence >= 0.7) {
                keyColumns.push(col.index);
            }
        }
    }
    
    // If no aggregation fields found or we need more keys, use traditional approach
    if (keyColumns.length === 0) {
        // Always include the highest scoring column
        if (columnScores.length > 0) {
            keyColumns.push(columnScores[0].index);
        }
    }
    
    // Add additional high-scoring columns if needed (avoid duplicates)
    for (let i = 0; i < Math.min(3, columnScores.length) && keyColumns.length < 3; i++) {
        const col = columnScores[i];
        if (!keyColumns.includes(col.index)) {
            // Include if score is high and uniqueness is good
            if (col.score >= 6 && col.uniquenessRatio >= 0.7) {
                keyColumns.push(col.index);
            }
        }
    }
    
    // Ensure we have at least one key column
    if (keyColumns.length === 0) {
        keyColumns = [0];
    }
    
    // Sort key columns by their original position
    keyColumns.sort((a, b) => a - b);
    
    // console.log(`✅ Selected key columns: [${keyColumns.join(', ')}] = [${keyColumns.map(i => `"${headers[i]}"`).join(', ')}]`);
    // if (aggregationFields.length > 0) {
    //     console.log(`🗓️ Aggregation priority applied for: ${keyColumns.filter(i => aggregationFields.some(col => col.index === i)).map(i => `"${headers[i]}"`).join(', ')}`);
    // }
    
    return keyColumns;
}

// Function to perform fuzzy matching for both export and display purposes
function performFuzzyMatchingForExport(body1, body2, finalHeaders, finalAllCols, isLargeFile, tableHeaders) {
    
    // Smart detection of key columns
    const combinedData = [finalHeaders, ...body1, ...body2];
    const keyColumnIndexes = smartDetectKeyColumns(finalHeaders, combinedData);
    
    // --- Fuzzy matching for bottom table ---
    let used2 = new Array(body2.length).fill(false);
    let pairs = [];
    
    function countMatches(rowA, rowB) {
        let matches = 0;
        let keyMatches = 0; // Count matches in key columns
        let toleranceMatches = 0; // Count tolerance matches
        let keyToleranceMatches = 0; // Count tolerance matches in key columns
        
        // Use smart detected key columns instead of 70% rule
        const keyColumnsIndexes = keyColumnIndexes;
        
        for (let i = 0; i < finalAllCols; i++) {
            let valueA = (rowA[i] || '').toString();
            let valueB = (rowB[i] || '').toString();
            
            if (toleranceMode) {
                // Use tolerance comparison
                const compResult = compareValuesWithTolerance(valueA, valueB);
                
                if (compResult === 'identical') {
                    matches++;
                    if (keyColumnsIndexes.includes(i)) {
                        keyMatches++;
                    }
                } else if (compResult === 'tolerance') {
                    toleranceMatches++;
                    if (keyColumnsIndexes.includes(i)) {
                        keyToleranceMatches++;
                    }
                }
            } else {
                // Use exact comparison (case-insensitive)
                if (valueA.toUpperCase() === valueB.toUpperCase()) {
                    matches++;
                    if (keyColumnsIndexes.includes(i)) {
                        keyMatches++;
                    }
                }
            }
        }
        
        if (toleranceMode) {
            // In tolerance mode, both exact matches and tolerance matches contribute to score
            // Exact matches get full weight, tolerance matches get 70% weight
            const adjustedToleranceMatches = toleranceMatches * 0.7;
            const adjustedKeyToleranceMatches = keyToleranceMatches * 0.7;
            
            const totalKeyScore = (keyMatches * 3) + (adjustedKeyToleranceMatches * 3);
            const totalOtherScore = (matches - keyMatches) + (adjustedToleranceMatches - adjustedKeyToleranceMatches);
            
            return totalKeyScore + totalOtherScore;
        } else {
            // Original logic for exact matching
            const otherMatches = matches - keyMatches;
            const weightedScore = (keyMatches * 3) + otherMatches;
            return weightedScore;
        }
    }
    
    // Process large files in batches to avoid freezing the UI
    function processBatch(startIndex, batchSize) {
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
            // Calculate threshold for matching based on smart detected key columns
            const keyColumnsCount = keyColumnIndexes.length;
            let minKeyMatches = Math.ceil(keyColumnsCount * 0.6) * 3; // 60% of detected key columns with 3x weight
            let minTotalMatches = Math.ceil(finalAllCols * 0.5); // 50% of all columns
            
            // In tolerance mode, require higher threshold since tolerance matches have reduced weight
            if (toleranceMode) {
                minKeyMatches = Math.ceil(keyColumnsCount * 0.8) * 3; // 80% of detected key columns with 3x weight
                minTotalMatches = Math.ceil(finalAllCols * 0.7); // 70% of all columns
            }
            
            if (bestScore >= minKeyMatches || bestScore >= minTotalMatches) {
                pairs.push({row1: body1[i], row2: body2[bestIdx]});
                used2[bestIdx] = true;
            } else {
                pairs.push({row1: body1[i], row2: null});
            }
        }
        
        // Continue with next batch or finish
        if (endIndex < body1.length) {
            // Show progress for large files
            if (body1.length > 1000) {
                const progress = Math.round((endIndex / body1.length) * 100);
                const progressMessage = isLargeFile ? 
                    `🔄 Processing large files for export... ${progress}% complete` :
                    `🔄 Comparing files... ${progress}% complete`;
                document.getElementById('result').innerHTML = `<div style="text-align: center; padding: 20px; font-size: 16px;">${progressMessage}</div>`;
            }
            setTimeout(() => processBatch(endIndex, batchSize), 10);
        } else {
            // Add unmatched rows from body2
            for (let j = 0; j < body2.length; j++) {
                if (!used2[j]) {
                    pairs.push({row1: null, row2: body2[j]});
                }
            }
            
            // Save pairs and headers for sorting
            currentPairs = pairs;
            currentFinalHeaders = finalHeaders;
            currentFinalAllCols = finalAllCols;
            
            // Update summary statistics based on actual matching results
            updateSummaryStatistics(tableHeaders, body1.length, body2.length);
            
            if (isLargeFile) {
                // Show large file message instead of table
                showLargeFileMessage(Math.max(body1.length, body2.length));
            } else {
                // Continue with rendering using universal function
                renderComparisonTable();
            }
        }
    }
    
    // Determine batch size based on file size
    const batchSize = body1.length > 5000 ? 100 : body1.length > 1000 ? 250 : 1000;
    processBatch(0, batchSize);
}

// Helper functions for tolerance comparison
function isDateString(str) {
    if (!str || typeof str !== 'string') return false;
    
    // Remove quotes and trim
    const cleanStr = str.replace(/['"]/g, '').trim();
    
    // Check for common date patterns
    const datePatterns = [
        /^\d{4}-\d{2}-\d{2}/, // 2023-12-31
        /^\d{2}\/\d{2}\/\d{4}/, // 12/31/2023
        /^\d{2}\.\d{2}\.\d{4}/, // 31.12.2023
        /^\d{1,2}\/\d{1,2}\/\d{4}/, // 1/1/2023
        /^\d{1,2}-\d{1,2}-\d{4}/, // 1-1-2023
    ];
    
    return datePatterns.some(pattern => pattern.test(cleanStr));
}

function isNumericString(str) {
    if (!str || typeof str !== 'string') return false;
    
    // Remove quotes, commas, and trim
    const cleanStr = str.replace(/['",$\s]/g, '').trim();
    
    // Check if it's a valid number
    return !isNaN(cleanStr) && !isNaN(parseFloat(cleanStr)) && isFinite(cleanStr);
}

function extractDateOnly(str) {
    if (!str) return '';
    
    const cleanStr = str.toString().replace(/['"]/g, '').trim();
    
    // Extract date part from datetime strings
    const dateMatch = cleanStr.match(/(\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}|\d{2}\.\d{2}\.\d{4}|\d{1,2}\/\d{1,2}\/\d{4}|\d{1,2}-\d{1,2}-\d{4})/);
    
    return dateMatch ? dateMatch[0] : cleanStr;
}

function parseNumber(str) {
    if (!str) return 0;
    
    // Remove quotes, commas, currency symbols and trim
    const cleanStr = str.toString().replace(/['",$\s]/g, '').trim();
    
    return parseFloat(cleanStr) || 0;
}

function isWithinTolerance(val1, val2, tolerance = 0.015) {
    const num1 = parseNumber(val1);
    const num2 = parseNumber(val2);
    
    if (num1 === 0 && num2 === 0) return true;
    if (num1 === 0 || num2 === 0) return false;
    
    const diff = Math.abs(num1 - num2);
    const avg = (Math.abs(num1) + Math.abs(num2)) / 2;
    
    return (diff / avg) <= tolerance;
}

function compareValuesWithTolerance(v1, v2) {
    if (!v1 && !v2) return 'identical';
    if (!v1 || !v2) return 'different';
    
    const str1 = v1.toString().trim();
    const str2 = v2.toString().trim();
    
    // Exact match (case-insensitive)
    if (str1.toUpperCase() === str2.toUpperCase()) {
        return 'identical';
    }
    
    // Check if both are dates
    if (isDateString(str1) && isDateString(str2)) {
        const date1 = extractDateOnly(str1);
        const date2 = extractDateOnly(str2);
        
        if (date1 === date2) {
            return 'tolerance'; // Same date, different time
        }
    }
    
    // Check if both are numbers
    if (isNumericString(str1) && isNumericString(str2)) {
        if (isWithinTolerance(str1, str2, 0.015)) {
            return 'tolerance'; // Within 1.5% tolerance
        }
    }
    
    return 'different';
}

// Universal function to render comparison table with consistent styling
function renderComparisonTable() {
    if (!currentPairs || currentPairs.length === 0) {
        document.querySelector('.diff-table-header thead').innerHTML = '';
        document.querySelector('.filter-row').innerHTML = '';
        document.querySelector('.diff-table-body tbody').innerHTML = '<tr><td colspan="100" style="text-align:center; padding:20px;">No data to display</td></tr>';
        return;
    }
    
    // Get filter states
    const hideSameEl = document.getElementById('hideSameRows');
    const hideDiffEl = document.getElementById('hideDiffColumns');
    const hideNewRows1El = document.getElementById('hideNewRows1');
    const hideNewRows2El = document.getElementById('hideNewRows2');
    
    let hideSame = hideSameEl ? hideSameEl.checked : false;
    const hideDiffRows = hideDiffEl ? hideDiffEl.checked : false;
    const hideNewRows1 = hideNewRows1El ? hideNewRows1El.checked : false;
    const hideNewRows2 = hideNewRows2El ? hideNewRows2El.checked : false;
    
    // Table headers with consistent styling
    let headerHtml = '<tr><th title="Source - shows which file the data comes from">Source</th>';
    for (let c = 0; c < currentFinalAllCols; c++) {
        let sortClass = 'sortable';
        if (c === currentSortColumn) {
            sortClass += currentSortDirection === 'asc' ? ' sort-asc' : ' sort-desc';
        }
        let headerText = currentFinalHeaders[c] !== undefined ? currentFinalHeaders[c] : '';
        let titleAttr = headerText ? ` title="${headerText.toString().replace(/"/g, '&quot;')}"` : '';
        headerHtml += `<th class="${sortClass}" onclick="sortTable(${c})"${titleAttr}>${headerText}</th>`;
    }
    headerHtml += '</tr>';
    document.querySelector('.diff-table-header thead').innerHTML = headerHtml;
    
    // Filter row with consistent styling
    let filterHtml = '<tr><td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>';
    for (let c = 0; c < currentFinalAllCols; c++) {
        filterHtml += `<td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>`;
    }
    filterHtml += '</tr>';
    document.querySelector('.filter-row').innerHTML = filterHtml;
    
    // Count visible rows first to handle "no results" properly
    let visibleRowCount = 0;
    let tempBodyHtml = '';
    
    currentPairs.forEach(pair => {
        let row1 = pair.row1;
        let row2 = pair.row2;
        
        // Pre-convert values to uppercase for performance
        let row1Upper = row1 ? row1.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        let row2Upper = row2 ? row2.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        
        let isEmpty = true;
        let hasWarn = false;
        let allSame = true;
        let hasDiff = false;
        
        for (let c = 0; c < currentFinalAllCols; c++) {
            let v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
            let v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
            if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                isEmpty = false;
            }
            if (row1 && row2) {
                if (row1Upper[c] !== row2Upper[c]) {
                    hasWarn = true;
                    allSame = false;
                } 
            } else {
                hasDiff = true;
                allSame = false;
            }
        }
        
        if (isEmpty) return;
        if (hideSame && row1 && row2 && allSame) return;
        if (hideNewRows1 && row1 && !row2) return;
        if (hideNewRows2 && !row1 && row2) return;
        if (hideDiffRows && row1 && row2 && hasWarn) return;
        
        visibleRowCount++;
    });
    
    // If no visible rows after filtering, show appropriate message
    if (visibleRowCount === 0) {
        const activeFilters = [];
        if (hideSame) activeFilters.push('Hide identical rows');
        if (hideNewRows1) activeFilters.push('Hide rows only in File 1');
        if (hideNewRows2) activeFilters.push('Hide rows only in File 2');
        if (hideDiffRows) activeFilters.push('Hide rows with differences');
        
        let message = 'No rows match the current filters';
        if (activeFilters.length > 0) {
            message += ': ' + activeFilters.join(', ');
        }
        
        document.querySelector('.diff-table-body tbody').innerHTML = `<tr><td colspan="100" style="text-align:center; padding:20px;">${message}</td></tr>`;
        return;
    }

    // Table body with consistent styling
    let bodyHtml = '';
    currentPairs.forEach(pair => {
        let row1 = pair.row1;
        let row2 = pair.row2;
        
        // Pre-convert values to uppercase for performance (for non-tolerance mode)
        let row1Upper = row1 ? row1.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        let row2Upper = row2 ? row2.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        
        let isEmpty = true;
        let hasWarn = false;
        let hasTolerance = false;
        let allSame = true;
        let hasDiff = false;
        
        // Store comparison results for each column
        let columnComparisons = [];
        
        for (let c = 0; c < currentFinalAllCols; c++) {
            let v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
            let v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
            
            if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                isEmpty = false;
            }
            
            if (row1 && row2) {
                if (toleranceMode) {
                    // Use tolerance comparison
                    const comparisonResult = compareValuesWithTolerance(v1, v2);
                    columnComparisons[c] = comparisonResult;
                    
                    if (comparisonResult === 'different') {
                        hasWarn = true;
                        allSame = false;
                    } else if (comparisonResult === 'tolerance') {
                        hasTolerance = true;
                        allSame = false;
                    }
                } else {
                    // Use exact comparison
                    if (row1Upper[c] !== row2Upper[c]) {
                        hasWarn = true;
                        allSame = false;
                        columnComparisons[c] = 'different';
                    } else {
                        columnComparisons[c] = 'identical';
                    }
                }
            } else {
                hasDiff = true;
                allSame = false;
                columnComparisons[c] = 'different';
            }
        }
        
        if (isEmpty) return;
        if (hideSame && row1 && row2 && allSame) return;
        if (hideNewRows1 && row1 && !row2) return;
        if (hideNewRows2 && !row1 && row2) return;
        if (hideDiffRows && row1 && row2 && (hasWarn || hasTolerance)) return;
        
        // If both files have data and they differ, create two separate rows
        if (row1 && row2 && (hasWarn || hasTolerance)) {
            // Row for File 1
            bodyHtml += `<tr class="warn-row warn-row-group-start">`;
            bodyHtml += `<td class="warn-cell">${getFileDisplayName(fileName1, sheetName1) || 'File 1'}</td>`;
            for (let c = 0; c < currentFinalAllCols; c++) {
                let v1 = row1[c] !== undefined ? row1[c] : '';
                let compResult = columnComparisons[c];
                
                if (compResult === 'identical') {
                    // Same values - will be merged with rowspan
                    bodyHtml += `<td class="identical" rowspan="2" style="vertical-align: middle; text-align: center;">${v1}</td>`;
                } else if (compResult === 'tolerance') {
                    // Values within tolerance - show in orange
                    bodyHtml += `<td class="tolerance-cell">${v1}</td>`;
                } else {
                    // Different values - show in red
                    bodyHtml += `<td class="warn-cell">${v1}</td>`;
                }
            }
            bodyHtml += '</tr>';
            
            // Row for File 2
            bodyHtml += `<tr class="warn-row warn-row-group-end">`;
            bodyHtml += `<td class="warn-cell">${getFileDisplayName(fileName2, sheetName2) || 'File 2'}</td>`;
            for (let c = 0; c < currentFinalAllCols; c++) {
                let v2 = row2[c] !== undefined ? row2[c] : '';
                let compResult = columnComparisons[c];
                
                if (compResult === 'tolerance') {
                    // Values within tolerance - show in orange
                    bodyHtml += `<td class="tolerance-cell">${v2}</td>`;
                } else if (compResult === 'different') {
                    // Different values - show in red
                    bodyHtml += `<td class="warn-cell">${v2}</td>`;
                }
                // For identical values, we already added rowspan=2 in the first row, so skip here
            }
            bodyHtml += '</tr>';
        } else {
            // Single row for identical data or data from only one file
            let source = '';
            let rowClass = '';
            
            if (row1 && row2 && allSame) {
                source = 'Both files';
                rowClass = 'row-identical';
            } else if (row1 && !row2) {
                source = getFileDisplayName(fileName1, sheetName1) || 'File 1';
                rowClass = 'new-row1';
            } else if (!row1 && row2) {
                source = getFileDisplayName(fileName2, sheetName2) || 'File 2';
                rowClass = 'new-row2';
            }
            
            bodyHtml += `<tr class="${rowClass}">`;
            if (row1 && row2 && allSame) {
                bodyHtml += `<td class="identical">${source}</td>`;
            } else if (row1 && !row2) {
                bodyHtml += `<td class="new-cell1">${source}</td>`;
            } else if (!row1 && row2) {
                bodyHtml += `<td class="new-cell2">${source}</td>`;
            } else {
                bodyHtml += `<td>${source}</td>`;
            }
            
            // Data columns
            for (let c = 0; c < currentFinalAllCols; c++) {
                let cellValue = '';
                let cellClass = '';
                
                if (row1 && !row2) {
                    cellValue = row1[c] !== undefined ? row1[c] : '';
                    cellClass = 'new-cell1';
                } else if (!row1 && row2) {
                    cellValue = row2[c] !== undefined ? row2[c] : '';
                    cellClass = 'new-cell2';
                } else if (row1 && row2) {
                    // Both files have the same data
                    cellValue = row1[c] !== undefined ? row1[c] : '';
                    cellClass = 'identical';
                }
                
                bodyHtml += `<td class="${cellClass}">${cellValue}</td>`;
            }
            bodyHtml += '</tr>';
        }
    });
    
    document.querySelector('.diff-table-body tbody').innerHTML = bodyHtml;
    
    // Apply consistent styling and synchronization immediately
    setTimeout(() => {
        // Ensure tables have proper CSS classes and structure
        const headerTable = document.querySelector('.diff-table-header');
        const bodyTable = document.querySelector('.diff-table-body');
        const container = document.querySelector('.table-container-sync');
        
        if (container) {
            container.style.position = 'relative';
            container.style.overflow = 'hidden';
            container.style.border = '1px solid #e0e0e0';
            container.style.borderRadius = '8px';
        }
        
        if (headerTable) {
            headerTable.style.borderCollapse = 'separate';
            headerTable.style.borderSpacing = '0';
            headerTable.style.tableLayout = 'fixed';
            // Don't set width to 100% - let syncColumnWidths handle it
        }
        
        if (bodyTable) {
            bodyTable.style.borderCollapse = 'separate';
            bodyTable.style.borderSpacing = '0';
            bodyTable.style.tableLayout = 'fixed';
            // Don't set width to 100% - let syncColumnWidths handle it
        }
        
        // Apply full synchronization with enhanced width enforcement
        // Order matters: syncColumnWidths first, then forceTableWidthSync for consistency
        syncColumnWidths();
        forceTableWidthSync();
        syncTableScroll();
    }, 50);
}

// Sorting function
function sortTable(column) {
    if (currentSortColumn === column) {
        currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortColumn = column;
        currentSortDirection = 'asc';
    }
    
    // Sort pairs
    currentPairs.sort((a, b) => {
        let aVal = '';
        let bVal = '';
        
        // Use value from first row if available, otherwise from second
        if (a.row1 && a.row1[column] !== undefined) {
            aVal = a.row1[column].toString();
        } else if (a.row2 && a.row2[column] !== undefined) {
            aVal = a.row2[column].toString();
        }
        if (b.row1 && b.row1[column] !== undefined) {
            bVal = b.row1[column].toString();
        } else if (b.row2 && b.row2[column] !== undefined) {
            bVal = b.row2[column].toString();
        }
        
        if (currentSortDirection === 'asc') {
            return aVal.localeCompare(bVal);
        } else {
            return bVal.localeCompare(aVal);
        }
    });
    
    // Update display using universal rendering function
    renderComparisonTable();
}

// Table filtering function
function filterTable() {
    let filters = [];
    let filterInputs = document.querySelectorAll('.filter-row input[type="text"]');
    filterInputs.forEach(input => {
        filters.push(input.value.toLowerCase());
    });
    
    let rows = document.querySelectorAll('.diff-table-body tbody tr');
    
    // Process rows in pairs or individually
    for (let i = 0; i < rows.length; i++) {
        let row = rows[i];
        let cells = row.querySelectorAll('td');
        
        // Check if this row has rowspan cells (first row of a pair)
        let hasRowspan = Array.from(cells).some(cell => cell.hasAttribute('rowspan'));
        
        if (hasRowspan && i + 1 < rows.length) {
            // Process pair of rows together
            let nextRow = rows[i + 1];
            let show = false;
            
            // Check first row
            show = checkRowMatchesFilters(row, filters) || checkRowMatchesFilters(nextRow, filters, row);
            
            // Apply visibility to both rows
            row.style.display = show ? '' : 'none';
            nextRow.style.display = show ? '' : 'none';
            
            i++; // Skip next row since we processed it
        } else {
            // Process single row
            let show = checkRowMatchesFilters(row, filters);
            row.style.display = show ? '' : 'none';
        }
    }
    
    // Force comprehensive table reset and resync after filtering
    refreshTableLayout();
}

// Helper function to check if a row matches filters
function checkRowMatchesFilters(row, filters, rowspanRow = null) {
    let cells = row.querySelectorAll('td');
    
    for (let j = 0; j < filters.length; j++) {
        if (filters[j] && filters[j].trim() !== '') {
            let cellText = '';
            
            if (j < cells.length) {
                cellText = cells[j].textContent.toLowerCase();
            } else if (rowspanRow) {
                // This might be a case where cell is covered by rowspan from previous row
                let rowspanCells = rowspanRow.querySelectorAll('td');
                
                // Try to find the corresponding cell accounting for rowspan
                let adjustedIndex = j;
                let currentCellIndex = 0;
                
                for (let k = 0; k <= j && currentCellIndex < rowspanCells.length; k++) {
                    if (k === j) {
                        if (rowspanCells[currentCellIndex] && rowspanCells[currentCellIndex].hasAttribute('rowspan')) {
                            // This cell spans down to current row
                            cellText = rowspanCells[currentCellIndex].textContent.toLowerCase();
                        }
                        break;
                    }
                    
                    // Move to next cell if current one doesn't have rowspan
                    if (!rowspanCells[currentCellIndex] || !rowspanCells[currentCellIndex].hasAttribute('rowspan')) {
                        currentCellIndex++;
                    }
                }
            }
            
            if (cellText && !cellText.includes(filters[j])) {
                return false;
            }
        }
    }
    
    return true;
}

// New function for comprehensive table layout refresh
function refreshTableLayout() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    const headerContainer = document.querySelector('.table-header-fixed');
    const bodyContainer = document.querySelector('.table-body-scrollable');
    
    if (!headerTable || !bodyTable) return;
    
    // Step 1: Reset table layout completely
    headerTable.style.tableLayout = 'auto';
    bodyTable.style.tableLayout = 'auto';
    
    // Step 2: Force a reflow
    headerTable.offsetHeight;
    bodyTable.offsetHeight;
    
    // Step 3: Apply fixed layout and sync column widths
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    // Step 4: Apply our synchronized column widths
    setTimeout(() => {
        syncColumnWidths();
        forceTableWidthSync();
        
        // Step 5: Ensure containers are properly sized
        if (headerContainer && bodyContainer) {
            headerContainer.scrollLeft = bodyContainer.scrollLeft;
        }
    }, 10);
}

// Function to set correct table width (only called once during initial render)
function adjustTableWidth() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    
    if (headerTable && bodyTable) {
        // Get the number of columns
        const headerCells = headerTable.querySelectorAll('th');
        const numColumns = headerCells.length;
        
        if (numColumns > 0) {
            // Calculate fixed total width: Source column (240px) + data columns (180px each)
            const sourceColumnWidth = 240;
            const dataColumnWidth = 180;
            const totalWidth = sourceColumnWidth + ((numColumns - 1) * dataColumnWidth);
            
            // Set consistent table widths
            const widthStyle = totalWidth + 'px';
            headerTable.style.width = widthStyle;
            bodyTable.style.width = widthStyle;
            headerTable.style.minWidth = widthStyle;
            bodyTable.style.minWidth = widthStyle;
            
            // Synchronize column widths after setting table width
            setTimeout(() => {
                syncColumnWidths();
            }, 10);
        }
    }
}

// Performance monitoring functions for large files
function getMemoryUsage() {
    if (performance && performance.memory) {
        return {
            used: Math.round(performance.memory.usedJSHeapSize / 1048576), // MB
            total: Math.round(performance.memory.totalJSHeapSize / 1048576), // MB
            limit: Math.round(performance.memory.jsHeapSizeLimit / 1048576) // MB
        };
    }
    return null;
}

function showMemoryWarning() {
    const memory = getMemoryUsage();
    if (memory && memory.used > memory.limit * 0.8) {
        const warningDiv = document.createElement('div');
        warningDiv.innerHTML = `
            <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; padding: 10px; margin: 10px 0; border-radius: 5px;">
                <strong>⚠️ High Memory Usage:</strong> The browser is using ${memory.used}MB of ${memory.limit}MB available memory.
                <br><small>Consider using filters to reduce the amount of data displayed.</small>
            </div>
        `;
        document.getElementById('result').appendChild(warningDiv);
    }
}

// Optimized CSV parsing for large files
function parseCSVChunked(csvText, chunkSize = 1000) {
    const lines = csvText.split(/\r?\n/);
    const result = [];
    const delimiter = detectCSVDelimiter(csvText);
    
    // Process in chunks to avoid blocking the UI
    function processChunk(startIndex) {
        const endIndex = Math.min(startIndex + chunkSize, lines.length);
        
        for (let i = startIndex; i < endIndex; i++) {
            const line = lines[i];
            if (line.trim() === '') continue;
            
            const row = [];
            let current = '';
            let inQuotes = false;
            
            for (let j = 0; j < line.length; j++) {
                const char = line[j];
                
                if (char === '"') {
                    if (inQuotes && line[j + 1] === '"') {
                        current += '"';
                        j++;
                    } else {
                        inQuotes = !inQuotes;
                    }
                } else if (char === delimiter && !inQuotes) {
                    row.push(parseCSVValue(current.trim()));
                    current = '';
                } else {
                    current += char;
                }
            }
            
            row.push(parseCSVValue(current.trim()));
            
            if (row.some(cell => cell !== null && cell !== undefined && cell.toString().trim() !== '')) {
                result.push(row);
            }
        }
        
        if (endIndex < lines.length) {
            // Continue processing next chunk
            setTimeout(() => processChunk(endIndex), 10);
        }
    }
    
    return new Promise((resolve) => {
        processChunk(0);
        // For now, return synchronous parsing for compatibility
        // In future, this could be made fully asynchronous
        resolve(parseCSV(csvText));
    });
}

// Initialize synchronization and placeholder on page load
window.addEventListener('load', function() {
    syncTableScroll();
    showPlaceholderMessage(); // Show placeholder message on page load
});

// Initialize file handlers
document.addEventListener('DOMContentLoaded', function() {
    // Initialize button layout - hide export button container initially
    const exportBtn = document.getElementById('exportExcelBtn');
    const buttonsContainer = document.querySelector('.buttons-container');
    const exportButtonHalf = exportBtn ? exportBtn.closest('.button-half') : null;
    
    if (exportBtn && buttonsContainer && exportButtonHalf) {
        exportBtn.style.display = 'none';
        exportButtonHalf.classList.add('export-hidden');
        buttonsContainer.classList.add('export-hidden');
    }
    
    document.getElementById('file1').addEventListener('change', function(e) {
        handleFile(e.target.files[0], 1);
    });
    document.getElementById('file2').addEventListener('change', function(e) {
        handleFile(e.target.files[0], 2);
    });
    
    // Add event listeners for filter checkboxes
    const hideSameRowsEl = document.getElementById('hideSameRows');
    if (hideSameRowsEl) {
        hideSameRowsEl.addEventListener('change', function() {
            if (currentPairs && currentPairs.length > 0) {
                renderComparisonTable();
            }
        });
    }
    
    const hideDiffColumnsEl = document.getElementById('hideDiffColumns');
    if (hideDiffColumnsEl) {
        hideDiffColumnsEl.addEventListener('change', function() {
            if (currentPairs && currentPairs.length > 0) {
                renderComparisonTable();
            }
        });
    }
    
    const hideNewRows1El = document.getElementById('hideNewRows1');
    if (hideNewRows1El) {
        hideNewRows1El.addEventListener('change', function() {
            if (currentPairs && currentPairs.length > 0) {
                renderComparisonTable();
            }
        });
    }
    
    const hideNewRows2El = document.getElementById('hideNewRows2');
    if (hideNewRows2El) {
        hideNewRows2El.addEventListener('change', function() {
            if (currentPairs && currentPairs.length > 0) {
                renderComparisonTable();
            }
        });
    }
    
    // Initialize table with empty values on page load
    const tableHeaders = getSummaryTableHeaders();
    let htmlSummary = `
        <table style="margin-bottom:20px; border: 1px solid #ccc;">
            <tr><th>${tableHeaders.file}</th><th>${tableHeaders.rowCount}</th><th>${tableHeaders.rowsOnlyInFile}</th><th>${tableHeaders.identicalRows}</th><th>${tableHeaders.similarity}</th><th>${tableHeaders.diffColumns}</th></tr>
            <tr><td>-</td><td>-</td><td>-</td><td rowspan="2">-</td><td rowspan="2" class="percent-cell">-</td><td>-</td></tr>
            <tr><td>-</td><td>-</td><td>-</td><td>-</td></tr>
        </table>
    `;
    document.getElementById('summary').innerHTML = htmlSummary;
});

// Function to synchronize column widths between header and body tables
function syncColumnWidths() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    
    if (!headerTable || !bodyTable) return;
    
    const headerCells = headerTable.querySelectorAll('th');
    const allBodyRows = bodyTable.querySelectorAll('tbody tr'); // Get ALL rows, not just visible ones
    
    if (!headerCells.length) return;
    
    // Force table layout to fixed for consistent column behavior
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    const numColumns = headerCells.length;
    
    // Force table layout to fixed for consistent column behavior
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    // Calculate total table width and apply it
    const sourceColumnWidth = 240;
    const dataColumnWidth = 180;
    const calculatedTotalWidth = sourceColumnWidth + ((numColumns - 1) * dataColumnWidth);
    
    // Set consistent table widths
    headerTable.style.width = calculatedTotalWidth + 'px';
    bodyTable.style.width = calculatedTotalWidth + 'px';
    headerTable.style.minWidth = calculatedTotalWidth + 'px';
    bodyTable.style.minWidth = calculatedTotalWidth + 'px';
}

// Function to synchronize table scroll
function syncTableScroll() {
    const headerTable = document.querySelector('.table-header-fixed');
    const bodyTable = document.querySelector('.table-body-scrollable');
    
    if (headerTable && bodyTable) {
        // Remove existing scroll listener to avoid duplicates
        bodyTable.removeEventListener('scroll', syncScrollHandler);
        
        // Add the scroll synchronization
        bodyTable.addEventListener('scroll', syncScrollHandler);
    }
}

// Define scroll handler function separately to allow removal
function syncScrollHandler() {
    const headerTable = document.querySelector('.table-header-fixed');
    const bodyTable = document.querySelector('.table-body-scrollable');
    
    if (headerTable && bodyTable) {
        headerTable.scrollLeft = bodyTable.scrollLeft;
    }
}

// Function to force table width synchronization with enhanced column width enforcement
function forceTableWidthSync() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    
    if (!headerTable || !bodyTable) return;
    
    const headerCells = headerTable.querySelectorAll('th');
    const allBodyRows = bodyTable.querySelectorAll('tbody tr');
    
    if (headerCells.length === 0) return;
    
    // Force table layout to fixed
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    const numColumns = headerCells.length;
    
    // Apply table layout and calculate total width
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    // Calculate and set total table width
    const sourceColumnWidth = 240;
    const dataColumnWidth = 180;
    const calculatedTotalWidth = sourceColumnWidth + ((numColumns - 1) * dataColumnWidth);
    
    // Set consistent table widths
    headerTable.style.width = calculatedTotalWidth + 'px';
    bodyTable.style.width = calculatedTotalWidth + 'px';
    headerTable.style.minWidth = calculatedTotalWidth + 'px';
    bodyTable.style.minWidth = calculatedTotalWidth + 'px';
}

// Function to show message for large files (15K-45K rows) with summary and export only
function showLargeFileMessage(totalRows) {
    // Hide the detailed comparison table
    document.getElementById('diffTable').innerHTML = `
        <div style="text-align: center; padding: 40px; background-color: #e7f3ff; border: 1px solid #bee5eb; border-radius: 8px; margin: 20px 0;">
            <div style="font-size: 24px; margin-bottom: 16px;">📊</div>
            <div style="font-size: 18px; font-weight: 600; color: #0c5460; margin-bottom: 10px;">Large File Mode</div>
            <div style="color: #0c5460; margin-bottom: 15px; line-height: 1.5;">
                Your files contain <strong>${totalRows.toLocaleString()}</strong> rows.<br>
                For optimal performance, the detailed row-by-row comparison table is hidden.<br>
                <strong>The comparison has been completed successfully!</strong>
            </div>
            <div style="color: #0c5460; font-size: 14px; margin-bottom: 20px;">
                ✅ View the summary table above for statistics<br>
                ✅ Use the Export to Excel button to download detailed results<br>
                ✅ All comparison data is available in the export
            </div>
            <div style="background-color: #d1ecf1; padding: 15px; border-radius: 6px; margin-top: 15px;">
                <strong>💡 Performance Tip:</strong> Files with fewer than ${DETAILED_TABLE_LIMIT.toLocaleString()} rows 
                will show the detailed comparison table for interactive browsing.
            </div>
        </div>
    `;
    
    // Hide filter controls since there's no table to filter
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
    
    // Make sure export button is prominently visible
    const exportBtn = document.getElementById('exportExcelBtn');
    const buttonsContainer = document.querySelector('.buttons-container');
    const exportButtonHalf = exportBtn ? exportBtn.closest('.button-half') : null;
    
    if (exportBtn && buttonsContainer) {
        exportBtn.style.display = 'inline-block';
        exportBtn.style.fontSize = '16px';
        exportBtn.style.padding = '12px 24px';
        exportBtn.style.backgroundColor = '#28a745';
        exportBtn.style.fontWeight = 'bold';
        
        // Remove classes that hide export button container
        if (exportButtonHalf) {
            exportButtonHalf.classList.remove('export-hidden');
        }
        buttonsContainer.classList.remove('export-hidden');
    }
}
