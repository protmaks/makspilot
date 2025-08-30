const MAX_ROWS_LIMIT = 55000; 
const MAX_COLS_LIMIT = 120; 
const DETAILED_TABLE_LIMIT = 15000; 

let data1 = [], data2 = [];
let fileName1 = '', fileName2 = '';
let sheetName1 = '', sheetName2 = ''; 
let workbook1 = null, workbook2 = null; 
let toleranceMode = false; 

let currentSortColumn = -1;
let currentSortDirection = 'asc';
let currentPairs = [];
let currentFinalHeaders = [];
let currentFinalAllCols = 0;
let currentDiffColumns1 = '-';
let currentDiffColumns2 = '-';

function getSummaryTableHeaders() {
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

function getFileDisplayName(fileName, sheetName) {
    if (sheetName && sheetName.trim() !== '') {
        return `${fileName}:${sheetName}`;
    }
    return fileName;
}

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
        
        const existingInfos = sheetInfoContainer.querySelectorAll('.sheet-info');
        existingInfos.forEach(info => {
            if (info.dataset.fileNum === fileNum.toString()) {
                info.remove();
            }
        });
        
        sheetInfoContainer.appendChild(sheetInfo);
        sheetInfoContainer.style.display = 'flex';
    } else {
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

function populateSheetSelector(sheetNames, fileNum, selectedSheet) {
    const sheetSelector = document.getElementById(`sheetSelector${fileNum}`);
    const sheetSelect = document.getElementById(`sheetSelect${fileNum}`);
    
    if (!sheetSelector || !sheetSelect) return;
    
    if (sheetNames.length > 1) {
        sheetSelect.innerHTML = '';
        
        sheetNames.forEach((sheetName, index) => {
            const option = document.createElement('option');
            option.value = sheetName;
            option.textContent = sheetName;
            if (sheetName === selectedSheet) {
                option.selected = true;
            }
            sheetSelect.appendChild(option);
        });
        
        sheetSelector.style.display = 'block';
        
        sheetSelect.onchange = function() {
            processSelectedSheet(fileNum, this.value);
        };
    } else {
        sheetSelector.style.display = 'none';
    }
}

function processExcelSheetOptimized(sheet) {
    if (!sheet || !sheet['!ref']) {
        return [];
    }
    
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    let minRow = range.e.r + 1; 
    let maxRow = -1;
    let minCol = range.e.c + 1; 
    let maxCol = -1;
    
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
            const cell = sheet[cellAddress];
            
            
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
    y
    if (minRow > maxRow || minCol > maxCol) {
        return [];
    }
    
    const dataRange = {
        s: { r: minRow, c: minCol },
        e: { r: maxRow, c: maxCol }
    };
    e
    const optimizedSheet = {};
    optimizedSheet['!ref'] = XLSX.utils.encode_range(dataRange);
    
    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const originalAddress = XLSX.utils.encode_cell({r: row, c: col});
            const newAddress = XLSX.utils.encode_cell({r: row - minRow, c: col - minCol});
            
            if (sheet[originalAddress]) {
                optimizedSheet[newAddress] = sheet[originalAddress];
            }
        }
    }
    
    optimizedSheet['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: maxRow - minRow, c: maxCol - minCol }
    });
    
    const json = XLSX.utils.sheet_to_json(optimizedSheet, {
        header: 1, 
        defval: '',
        raw: true,          
        dateNF: 'yyyy-mm-dd hh:mm:ss'  
    });
    gh
    const filteredJson = json.filter(row => 
        Array.isArray(row) && row.some(cell => 
            cell !== null && cell !== undefined && cell.toString().trim() !== ''
        )
    );
    
    return filteredJson;
}

function processSelectedSheet(fileNum, selectedSheetName) {
    clearComparisonResults();
    
    const workbook = fileNum === 1 ? workbook1 : workbook2;
    const fileName = fileNum === 1 ? fileName1 : fileName2;
    
    if (fileNum === 1) {
        sheetName1 = selectedSheetName;
    } else {
        sheetName2 = selectedSheetName;
    }
    
    if (!workbook || !workbook.Sheets[selectedSheetName]) return;
    
    const tableElement = document.getElementById(fileNum === 1 ? 'table1' : 'table2');
    tableElement.innerHTML = '<div style="text-align: center; padding: 20px;">Loading sheet...</div>';
    
    setTimeout(() => {
        const sheet = workbook.Sheets[selectedSheetName];
        let json = processExcelSheetOptimized(sheet);
        
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

            if (fileNum === 1) {
                data1 = [];
            } else {
                data2 = [];
            }
            return;
        }
        
        json = json.filter(row => Array.isArray(row) && row.some(cell => (cell !== null && cell !== undefined && cell.toString().trim() !== '')));
        
        json = normalizeHeaders(json);
        
        json = processColumnTypes(json);
        
        json = removeEmptyColumns(json);
        json = json.map(row => {
            if (!Array.isArray(row)) return row;
            return row.map(cell => roundDecimalNumbers(cell));
        });
        
        if (fileNum === 1) {
            data1 = json;
            renderPreview(json, 'table1');
        } else {
            data2 = json;
            renderPreview(json, 'table2');
        }
        
        updateSheetInfo(fileName, workbook.SheetNames, selectedSheetName, fileNum);
    }, 10);
}

function processColumnTypes(data) {
    if (!data || data.length === 0) return data;
    
    const maxCols = Math.max(...data.map(row => Array.isArray(row) ? row.length : 0));
    if (maxCols === 0) return data;
    
    const columnTypes = [];
    const headers = data[0] || [];
    
    for (let colIndex = 0; colIndex < maxCols; colIndex++) {
        const columnValues = [];
        for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
            const value = data[rowIndex] && data[rowIndex][colIndex];
            if (value !== null && value !== undefined && value.toString().trim() !== '') {
                columnValues.push(value);
            }
        }
        
        const columnHeader = headers[colIndex] || '';
        const isDateCol = isDateColumn(columnValues, columnHeader);
        columnTypes[colIndex] = isDateCol ? 'date' : 'other';
    }
    
    const processedData = data.map((row, rowIndex) => {
        if (!Array.isArray(row) || rowIndex === 0) return row;
        
        return row.map((cell, colIndex) => {
            if (columnTypes[colIndex] === 'date') {
                return convertExcelDateNormalized(cell, true);
            } else {
                return cell;
            }
        });
    });
    
    return processedData;
}

function isDateColumn(columnValues, columnHeader = '') {
    if (!columnValues || columnValues.length === 0) return false;
    
    let dateCount = 0;
    let numberCount = 0;
    let totalCount = 0;
    let potentialExcelDates = 0;
    
    const headerLower = columnHeader.toString().toLowerCase();
    const dateKeywords = ['date', 'time', 'created', 'modified', 'updated', 'birth', 'дата', 'время', 'создан', 'изменен', 'обновлен'];
    const headerSuggestsDate = dateKeywords.some(keyword => headerLower.includes(keyword));
    
    for (let value of columnValues) {
        if (value && value.toString().trim() !== '') {
            totalCount++;
            
            if (typeof value === 'string') {
                if (value.match(/^\d{4}-\d{1,2}-\d{1,2}$/) || 
                    value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/) || 
                    value.match(/^\d{1,2}-\d{1,2}-\d{4}$/) || 
                    value.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/) || 
                    value.includes('T') || value.includes('GMT') || value.includes('UTC')) {
                    dateCount++;
                }
            } else if (value instanceof Date) {
                dateCount++;
            } else if (typeof value === 'number') {
                numberCount++;
                if (value >= 29221 && value <= 219146) {
                    potentialExcelDates++;
                }
            }
        }
    }
    
    if (totalCount === 0) return false;
    
    const dateRatio = dateCount / totalCount;
    const potentialDateRatio = potentialExcelDates / totalCount;
    const combinedDateRatio = (dateCount + potentialExcelDates) / totalCount;
    
    return (headerSuggestsDate && combinedDateRatio > 0.3) ||
           dateRatio > 0.5 || 
           (potentialDateRatio > 0.5 && numberCount > 0) || 
           combinedDateRatio > 0.6;
}


function formatDate(year, month, day, hour = null, minute = null, seconds = null) {
    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    
    
    const hasTime = (hour !== null && minute !== null) && 
                   !(hour === 0 && minute === 0 && (seconds === null || seconds === 0));
    
    if (hasTime) {
        result += ' ' + String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
        
        
        if (seconds !== null) {
            result += ':' + String(seconds).padStart(2, '0');
        }
    }
    
    return result;
}


function normalizeDateTime(dateTimeString) {
    
    if (dateTimeString.match(/\s+00:00:00$/)) {
        return dateTimeString.replace(/\s+00:00:00$/, '');
    }
    if (dateTimeString.match(/\s+00:00$/)) {
        return dateTimeString.replace(/\s+00:00$/, '');
    }
    return dateTimeString;
}


function convertExcelDate(value, isInDateColumn = false) {
    
    if (value instanceof Date) {
        if (!isNaN(value.getTime())) {
            
            const year = value.getFullYear();
            const month = value.getMonth() + 1;
            const day = value.getDate();
            const hour = value.getHours();
            const minute = value.getMinutes();
            const second = value.getSeconds();
            
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                
                return formatDate(year, month, day, hour, minute, second);
            }
        }
    }
    
    
    if (typeof value === 'number' && value > 1 && value < 300000) {
        
        
        
        if (isInDateColumn && value >= 29221 && value <= 219146) { 
            
            
            const wholeDays = Math.floor(value);
            const timeFraction = value - wholeDays;
            
            
            
            let days = wholeDays;
            
            
            if (wholeDays > 59) { 
                days = wholeDays - 1; 
            }
            
            
            
            const excelYear = 1900;
            let year = excelYear;
            let dayOfYear = days;
            
            
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
            
            
            const totalSeconds = Math.round(timeFraction * 24 * 60 * 60);
            const hours = Math.floor(totalSeconds / 3600);
            const minutes = Math.floor((totalSeconds % 3600) / 60);
            const seconds = totalSeconds % 60;
            
            if (year >= 1900 && year <= 2500) {
                
                const hasTime = timeFraction > 0 || hours !== 0 || minutes !== 0 || seconds !== 0;
                return formatDate(year, month, day, hasTime ? hours : null, hasTime ? minutes : null, hasTime ? seconds : null);
            }
        }
    }
    
    
    if (typeof value === 'string') {
        
        if (value.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
            return normalizeDateTime(value); 
        }
        
        
        if (value.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/)) {
            return normalizeDateTime(value + ':00'); 
        }
        
        
        let isoDateMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (isoDateMatch) {
            const year = parseInt(isoDateMatch[1]);
            const month = parseInt(isoDateMatch[2]);
            const day = parseInt(isoDateMatch[3]);
            if (day <= 31 && month <= 12 && year >= 1900 && year <= 2100) {
                
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        
        let isoDateTimeMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (isoDateTimeMatch) {
            const year = parseInt(isoDateTimeMatch[1]);
            const month = parseInt(isoDateTimeMatch[2]);
            const day = parseInt(isoDateTimeMatch[3]);
            const hour = parseInt(isoDateTimeMatch[4]);
            const minute = parseInt(isoDateTimeMatch[5]);
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                return value; 
            }
        }
        
        
        let isoDateTimeAMPMMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+(AM|PM)$/i);
        if (isoDateTimeAMPMMatch) {
            const year = parseInt(isoDateTimeAMPMMatch[1]);
            const month = parseInt(isoDateTimeAMPMMatch[2]); 
            const day = parseInt(isoDateTimeAMPMMatch[3]);
            let hour = parseInt(isoDateTimeAMPMMatch[4]);
            const minute = parseInt(isoDateTimeAMPMMatch[5]);
            const second = isoDateTimeAMPMMatch[6] ? parseInt(isoDateTimeAMPMMatch[6]) : 0;
            const ampm = isoDateTimeAMPMMatch[7].toUpperCase();
            
            
            if (ampm === 'AM') {
                if (hour === 12) hour = 0; 
            } else { 
                if (hour !== 12) hour += 12; 
            }
            
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59 && second >= 0 && second <= 59) {
                const result = formatDate(year, month, day, hour, minute, second);
                return result;
            }
        }
        
        
        let dateTimeMatchSec = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
        if (dateTimeMatchSec) {
            const first = parseInt(dateTimeMatchSec[1]);
            const second = parseInt(dateTimeMatchSec[2]);
            let year = parseInt(dateTimeMatchSec[3]);
            const hour = parseInt(dateTimeMatchSec[4]);
            const minute = parseInt(dateTimeMatchSec[5]);
            const sec = parseInt(dateTimeMatchSec[6]);
            
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            
            if (first <= 12 && second <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59 && sec >= 0 && sec <= 59) {
                const month = first;
                const day = second;
                return formatDate(year, month, day, hour, minute, sec);
            }
        }
        
        
        let shortYearAMPMMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+(AM|PM)$/i);
        if (shortYearAMPMMatch) {
            const first = parseInt(shortYearAMPMMatch[1]);
            const second = parseInt(shortYearAMPMMatch[2]);
            let year = parseInt(shortYearAMPMMatch[3]);
            let hour = parseInt(shortYearAMPMMatch[4]);
            const minute = parseInt(shortYearAMPMMatch[5]);
            const seconds = shortYearAMPMMatch[6] ? parseInt(shortYearAMPMMatch[6]) : 0;
            const ampm = shortYearAMPMMatch[7].toUpperCase();
            
            
            if (ampm === 'AM' && hour === 12) {
                hour = 0;  
            } else if (ampm === 'PM' && hour !== 12) {
                hour += 12;  
            }
            
            
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            
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
                    return formatDate(year, month, day, hour, minute, seconds);
                }
            }
        }

        
        let dateTimeMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const first = parseInt(dateTimeMatch[1]);
            const second = parseInt(dateTimeMatch[2]);
            let year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            const seconds = dateTimeMatch[6] ? parseInt(dateTimeMatch[6]) : 0;
            
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            
            if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                
                if (first > 12 && second <= 12) {
                    day = first;
                    month = second;
                } 
                
                else if (second > 12 && first <= 12) {
                    month = first;
                    day = second;
                }
                
                else if (first <= 12 && second <= 12) {
                    month = first;
                    day = second;
                }
                
                else {
                    return value;
                }
                
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                    
                    
                    result += ':' + String(seconds).padStart(2, '0');
                    
                    return result;
                }
            }
        }
        
        
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
                    return formatDate(year, month, day, hour, minute, sec);
                }
            }
        }
        
        
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
                    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                    
                    
                    result += ':' + String(seconds).padStart(2, '0');
                    
                    return result;
                }
            }
        }
        
        
        let dateTimeAMPMMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+(AM|PM)$/i);
        if (dateTimeAMPMMatch) {
            const first = parseInt(dateTimeAMPMMatch[1]);
            const second = parseInt(dateTimeAMPMMatch[2]);
            const year = parseInt(dateTimeAMPMMatch[3]);
            let hour = parseInt(dateTimeAMPMMatch[4]);
            const minute = parseInt(dateTimeAMPMMatch[5]);
            const second_time = dateTimeAMPMMatch[6] ? parseInt(dateTimeAMPMMatch[6]) : 0;
            const ampm = dateTimeAMPMMatch[7].toUpperCase();
            
            
            if (ampm === 'AM') {
                if (hour === 12) hour = 0; 
            } else { 
                if (hour !== 12) hour += 12; 
            }
            
            if (year >= 1900 && year <= 2100 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                
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
                    let result = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                    
                    
                    result += ':' + String(second_time).padStart(2, '0');
                    
                    return result;
                }
            }
        }
        
        
        let ampmNoSecondsMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})\s+(AM|PM)$/i);
        if (ampmNoSecondsMatch) {
            const year = parseInt(ampmNoSecondsMatch[1]);
            const month = parseInt(ampmNoSecondsMatch[2]);
            const day = parseInt(ampmNoSecondsMatch[3]);
            let hour = parseInt(ampmNoSecondsMatch[4]);
            const minute = parseInt(ampmNoSecondsMatch[5]);
            const ampm = ampmNoSecondsMatch[6].toUpperCase();
            
            
            if (ampm === 'AM') {
                if (hour === 12) hour = 0; 
            } else { 
                if (hour !== 12) hour += 12; 
            }
            
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                       String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0') + ':00';
            }
        }
        
        
        ampmNoSecondsMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{1,2})\s+(AM|PM)$/i);
        if (ampmNoSecondsMatch) {
            const first = parseInt(ampmNoSecondsMatch[1]);
            const second = parseInt(ampmNoSecondsMatch[2]);
            const year = parseInt(ampmNoSecondsMatch[3]);
            let hour = parseInt(ampmNoSecondsMatch[4]);
            const minute = parseInt(ampmNoSecondsMatch[5]);
            const ampm = ampmNoSecondsMatch[6].toUpperCase();
            
            
            if (ampm === 'AM') {
                if (hour === 12) hour = 0;
            } else { 
                if (hour !== 12) hour += 12;
            }
            
            if (year >= 1900 && year <= 2100 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                let month, day;
                
                
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
                    
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0') + ':00';
                }
            }
        }
        
        
        ampmNoSecondsMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})\s+(AM|PM)$/i);
        if (ampmNoSecondsMatch) {
            const first = parseInt(ampmNoSecondsMatch[1]);
            const second = parseInt(ampmNoSecondsMatch[2]);
            let year = parseInt(ampmNoSecondsMatch[3]);
            let hour = parseInt(ampmNoSecondsMatch[4]);
            const minute = parseInt(ampmNoSecondsMatch[5]);
            const ampm = ampmNoSecondsMatch[6].toUpperCase();
            
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            
            if (ampm === 'AM') {
                if (hour === 12) hour = 0;
            } else { 
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
                    
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                           String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0') + ':00';
                }
            }
        }
        
        
        dateTimeMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const first = parseInt(dateTimeMatch[1]);
            const second = parseInt(dateTimeMatch[2]);
            let year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            
            if (first > 12 && second <= 12 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                const day = first;
                const month = second;
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                       String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
            }
        }
        
        
        dateTimeMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/);
        if (dateTimeMatch) {
            const month = parseInt(dateTimeMatch[1]);
            const day = parseInt(dateTimeMatch[2]);
            let year = parseInt(dateTimeMatch[3]);
            const hour = parseInt(dateTimeMatch[4]);
            const minute = parseInt(dateTimeMatch[5]);
            
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' +
                       String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
            }
        }
        
        
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
        
        
        
        if (value.includes('T') || value.includes('GMT') || value.includes('UTC')) {
            try {
                
                if (value.includes('T')) {
                    
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
                
                
                const dateObj = new Date(value);
                if (!isNaN(dateObj.getTime())) {
                    
                    const year = dateObj.getFullYear();
                    const month = dateObj.getMonth() + 1;
                    const day = dateObj.getDate();
                    const hour = dateObj.getHours();
                    const minute = dateObj.getMinutes();
                    
                    if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                        
                        if (value.includes(':') || hour !== 0 || minute !== 0) {
                            return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0') + ' ' + 
                                   String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
                        } else {
                            return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                        }
                    }
                }
            } catch (e) {
                
            }
        }
        
        
        let dateMatch;
        
        
        dateMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const month = parseInt(dateMatch[2]);
            const year = parseInt(dateMatch[3]);
            if (day <= 31 && month <= 12 && year >= 1900 && year <= 2100) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dateMatch) {
            const first = parseInt(dateMatch[1]);
            const second = parseInt(dateMatch[2]);
            const year = parseInt(dateMatch[3]);
            
            if (year >= 1900 && year <= 2100) {
                let month, day;
                
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
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                }
            }
        }
        
        
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
        if (dateMatch) {
            const first = parseInt(dateMatch[1]);
            const second = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            
            
            let month, day;
            
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
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        
        dateMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);
        if (dateMatch) {
            const month = parseInt(dateMatch[1]);
            const day = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
        if (dateMatch) {
            const month = parseInt(dateMatch[1]);
            const day = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            
            year = year <= 30 ? 2000 + year : 1900 + year;
            if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        
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
        
        
        dateMatch = value.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            let year = parseInt(dateMatch[3]);
            
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
        
        
        dateMatch = value.match(/^([A-Za-z]{3})-(\d{1,2})-(\d{2})$/);
        if (dateMatch) {
            const monthName = dateMatch[1].toLowerCase();
            const day = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            
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
        
        
        dateMatch = value.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            let year = parseInt(dateMatch[3]);
            
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
        
        
        dateMatch = value.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{2})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const monthName = dateMatch[2].toLowerCase();
            let year = parseInt(dateMatch[3]);
            
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
        
        
        
        if (value.includes('-') && !isoDateMatch) {
            
            let parts = value.split('-');
            if (parts.length === 3) {
                let year = parseInt(parts[0]);
                let month = parseInt(parts[1]);
                let day = parseInt(parts[2]);
                
                
                if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                }
            }
        }
    }
    
    
    return value;
}


function convertExcelDateNormalized(value, isInDateColumn = false) {
    const result = convertExcelDate(value, isInDateColumn);
    
    
    if (typeof result === 'string' && result.match(/^\d{4}-\d{2}-\d{2}/)) {
        return normalizeDateTime(result);
    }
    
    return result;
}


function roundDecimalNumbers(value) {
    
    if (typeof value === 'number') {
        
        let rounded = Math.round(value * 100) / 100;
        
        
        if (Number.isInteger(rounded)) {
            return rounded;
        }
        
        return rounded;
    }
    
    
    if (typeof value === 'string' && !isNaN(value) && !isNaN(parseFloat(value))) {
        const numValue = parseFloat(value);
        
        
        let rounded = Math.round(numValue * 100) / 100;
        
        
        if (Number.isInteger(rounded)) {
            return rounded;
        }
        
        return rounded;
    }
    
    return value;
}

function removeEmptyColumns(data) {
    if (!data || data.length === 0) return data;
    
    
    let maxCols = 0;
    for (let i = 0; i < data.length; i++) {
        if (data[i] && data[i].length > maxCols) {
            maxCols = data[i].length;
        }
    }
    if (maxCols === 0) return data;
    
    
    let columnsToKeep = [];
    
    
    const isLargeFile = data.length > 1000;
    const checkRows = isLargeFile ? Math.min(data.length, 100) : data.length;
    
    columnLoop: for (let col = 0; col < maxCols; col++) {
        
        for (let row = 0; row < checkRows; row++) {
            if (data[row] && col < data[row].length) {
                let cellValue = data[row][col];
                if (cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '') {
                    columnsToKeep.push(col);
                    continue columnLoop;
                }
            }
        }
        
        
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
    
    
    if (result.length > 0) {
        const numCols = result[0].length;
        
        for (let col = 0; col < numCols; col++) {
            
            const sampleSize = Math.min(result.length, 20);
            const columnValues = [];
            for (let i = 0; i < sampleSize; i++) {
                columnValues.push(result[i][col]);
            }
            
            const columnHeader = result.length > 0 ? result[0][col] : '';
            const isDateCol = isDateColumn(columnValues, columnHeader);
            
            
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


function normalizeRowLengths(data) {
    if (!data || data.length === 0) return data;
    
    
    let maxCols = 0;
    for (let i = 0; i < data.length; i++) {
        if (data[i] && data[i].length > maxCols) {
            maxCols = data[i].length;
        }
    }
    
    
    return data.map(row => {
        if (!row) return new Array(maxCols).fill('');
        while (row.length < maxCols) {
            row.push('');
        }
        return row;
    });
}


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


function parseCSV(csvText) {
    const lines = csvText.split(/\r?\n/);
    const result = [];
    
    
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
    
    return result;
}


function parseCSVValue(value) {
    
    if (value.startsWith('"') && value.endsWith('"')) {
        value = value.slice(1, -1);
    }
    
    
    if (value === '' || value === null || value === undefined || value.toLowerCase() === 'null') {
        return '';
    }
    
    
    if (value.includes('?')) {
        
        const cleanValue = value.replace(/\?/g, '.');
        value = cleanValue;
    }
    
    
    if (value.match(/^\d{4}-\d{1,2}-\d{1,2}$/)) {
        
        return value;
    } else if (value.match(/^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{4}$/) ||
               value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/) ||
               value.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
        
        const parts = value.split(/[\/\.]/);
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        
        
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
        return value;
    } else if (value.match(/^\d{1,2}\.\d{1,2}\.\d{2}$/)) {
        
        const parts = value.split('.');
        const month = parseInt(parts[0], 10);
        const day = parseInt(parts[1], 10);
        let year = parseInt(parts[2], 10);
        
        
        year = year <= 30 ? 2000 + year : 1900 + year;
        
        
        if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        }
        return value;
    }
    return value; 
    
    
    let numValue = value.replace(',', '.'); 
    if (numValue !== '' && !isNaN(numValue) && !isNaN(parseFloat(numValue))) {
        const num = parseFloat(numValue);
        return Number.isInteger(num) ? num : num;
    }
    
    return value;
}

function handleFile(file, num) {
    if (!file) return;
    
    
    const tableElement = document.getElementById(num === 1 ? 'table1' : 'table2');
    tableElement.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">📊 Loading file... Please wait</div>';
    
    
    if (num === 1) {
        data1 = [];
        fileName1 = file.name;
        workbook1 = null;
        sheetName1 = ''; 
    } else {
        data2 = [];
        fileName2 = file.name;
        workbook2 = null;
        sheetName2 = ''; 
    }
    
    
    document.getElementById('result').innerHTML = '';
    document.getElementById('summary').innerHTML = '';
    
    
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
    
    
    const sheetInfoContainer = document.getElementById('sheetInfo');
    if (sheetInfoContainer) {
        
        const existingInfos = Array.from(sheetInfoContainer.querySelectorAll('.sheet-info'));
        existingInfos.forEach(info => {
            
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
    
    
    const sheetSelector = document.getElementById(`sheetSelector${num}`);
    if (sheetSelector) {
        sheetSelector.style.display = 'none';
    }
    
    
    const isCSV = file.name.toLowerCase().endsWith('.csv');
    
    if (isCSV) {
        
        const reader = new FileReader();
        reader.onload = function(e) {
            
            setTimeout(() => {
                const csvText = e.target.result;
                const json = parseCSV(csvText);
                
                
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
                
                
                if (json.length > 1000) {
                    tableElement.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚙️ Processing large CSV file... Please wait</div>';
                }
                
                
                setTimeout(() => {
                    
                    const headersNormalizedJson = normalizeHeaders(json);
                    
                    
                    const dateConvertedJson = processColumnTypes(headersNormalizedJson);
                    
                    setTimeout(() => {
                        
                        const normalizedJson = normalizeRowLengths(dateConvertedJson);
                        
                        
                        const cleanedJson = removeEmptyColumns(normalizedJson);
                        
                        setTimeout(() => {
                            
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
        
        const reader = new FileReader();
        reader.onload = function(e) {
            
            setTimeout(() => {
                let data = new Uint8Array(e.target.result);
                
                let workbook = XLSX.read(data, {
                    type: 'array',
                    cellDates: false,    
                    UTC: false  
                });
                
                
                if (num === 1) {
                    workbook1 = workbook;
                    sheetName1 = workbook.SheetNames[0]; 
                } else {
                    workbook2 = workbook;
                    sheetName2 = workbook.SheetNames[0]; 
                }
                
                
                if (workbook.SheetNames.length > 1) {
                    
                }
                
                
                updateSheetInfo(file.name, workbook.SheetNames, workbook.SheetNames[0], num);
                
                
                populateSheetSelector(workbook.SheetNames, num, workbook.SheetNames[0]);
                
                setTimeout(() => {
                    
                    let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    let json = processExcelSheetOptimized(firstSheet);
                    
                    
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
                    
                    
                    if (json.length > 1000) {
                        tableElement.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚙️ Processing large Excel file... Please wait</div>';
                    }
                    
                    setTimeout(() => {
                        
                        json = json.filter(row => Array.isArray(row) && row.some(cell => (cell !== null && cell !== undefined && cell.toString().trim() !== '')));
                        
                        
                        json = normalizeHeaders(json);
                        
                        
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


function generateLimitErrorMessage(type, current, limit, additionalInfo = '', secondType = null, secondCurrent = null, secondLimit = null) {
    if (secondType && secondCurrent !== null && secondLimit !== null) {
        
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


function createColumnMapping(header1, header2) {
    const mapping = [];
    const header1Lower = header1.map(h => (h || '').toString().toLowerCase().trim());
    const header2Lower = header2.map(h => (h || '').toString().toLowerCase().trim());
    
    
    const commonColumns = [];
    const used2 = new Set();
    
    header1Lower.forEach((col1, index1) => {
        if (col1 === '') return; 
        
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


function reorderDataByColumns(data, originalHeader, commonColumns, targetOrder) {
    if (!data || data.length === 0) return data;
    
    
    const newHeader = targetOrder.map(colInfo => colInfo.name);
    
    
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


function prepareDataForComparison(data1, data2) {
    if (!data1.length || !data2.length) {
        return { data1, data2, columnInfo: null };
    }
    
    const header1 = data1[0] || [];
    const header2 = data2[0] || [];
    
    
    const mapping = createColumnMapping(header1, header2);
    
    if (mapping.commonColumns.length === 0) {
        
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
    
    
    const unifiedOrder = mapping.commonColumns.map(col => ({
        name: col.name,
        sourceIndex: col.index1 
    }));
    
    
    const reorderedData1 = reorderDataByColumns(data1, header1, mapping.commonColumns, unifiedOrder);
    
    
    const unifiedOrderForData2 = mapping.commonColumns.map(col => ({
        name: col.name,
        sourceIndex: col.index2 
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


function normalizeHeaders(data) {
    if (!data || data.length === 0) return data;
    
    
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
    for (let i = 1; i < data.length; i++) { 
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
    
    return JSON.stringify(row);
}

function showPlaceholderMessage() {
    document.getElementById('result').innerHTML = '';
    document.getElementById('summary').innerHTML = '';
    
    
    const placeholderText = document.documentElement.getAttribute('data-placeholder-text') || 'Choose files on your computer and click Compare';
    
    document.getElementById('diffTable').innerHTML = `
        <div class="placeholder-message">
            <div class="placeholder-icon">📊</div>
            <div class="placeholder-text">${placeholderText}</div>
        </div>
    `;
    
    
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
}


function clearComparisonResults() {
    
    document.getElementById('result').innerHTML = '';
    document.getElementById('summary').innerHTML = '';
    document.getElementById('diffTable').innerHTML = '';
    
    
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
    
    
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


function getExcludedColumns() {
    const excludeInput = document.getElementById('excludeColumns');
    if (!excludeInput || !excludeInput.value.trim()) {
        return [];
    }
    return excludeInput.value.trim().split(',').map(col => col.trim()).filter(col => col.length > 0);
}


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


function filterRowExcludingColumns(row, excludedIndexes) {
    if (!excludedIndexes.length) return row;
    return row.filter((_, index) => !excludedIndexes.includes(index));
}


function getColumnTypes(header1, header2) {
    const header1Set = new Set(header1.map(h => h ? h.toString().trim().toLowerCase() : ''));
    const header2Set = new Set(header2.map(h => h ? h.toString().trim().toLowerCase() : ''));
    
    const columnTypes = {};
    
    
    header1.forEach((header, index) => {
        const headerKey = header ? header.toString().trim().toLowerCase() : '';
        if (headerKey) {
            if (header2Set.has(headerKey)) {
                columnTypes[index] = 'common'; 
            } else {
                columnTypes[index] = 'diff'; 
            }
        }
    });
    
    header2.forEach((header, index) => {
        const headerKey = header ? header.toString().trim().toLowerCase() : '';
        if (headerKey && !header1Set.has(headerKey)) {
            
            const combinedIndex = header1.length > header2.length ? 
                header1.findIndex(h => !h || h.toString().trim() === '') : 
                header1.length + index;
            columnTypes[combinedIndex] = 'new'; 
        }
    });
    
    return columnTypes;
}


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
    
    toleranceMode = useTolerance;
    
    
    clearComparisonResults();
    
    
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
    
    
    if (totalRows > 1000) {
        const fileInfo = `${data1.length.toLocaleString()} vs ${data2.length.toLocaleString()} rows`;
        const loadingMessage = totalRows > DETAILED_TABLE_LIMIT ? 
            `🔄 Processing large files (${fileInfo}) - Summary mode... Please wait` :
            `🔄 Comparing files (${fileInfo})... Please wait`;
        
        document.getElementById('result').innerHTML = `<div style="text-align: center; padding: 20px; font-size: 16px;">${loadingMessage}</div>`;
        document.getElementById('summary').innerHTML = '<div style="text-align: center; padding: 10px;">Analyzing data...</div>';
    }
    
    
    setTimeout(() => {
        performComparison();
    }, 10);
}

function performComparison() {
    
    const tableHeaders = getSummaryTableHeaders();
    
    
    restoreTableStructure();
    
    
    const hideSameRowsCheckbox = document.getElementById('hideSameRows');
    if (hideSameRowsCheckbox) {
        hideSameRowsCheckbox.checked = true;
    }
    
    
    const excludedColumns = getExcludedColumns();
    
    
    const hideDiffRowsEl = document.getElementById('hideDiffColumns');
    const hideNewRows1El = document.getElementById('hideNewRows1');
    const hideNewRows2El = document.getElementById('hideNewRows2');
    
    const hideDiffRows = hideDiffRowsEl ? hideDiffRowsEl.checked : false; 
    const hideNewRows1 = hideNewRows1El ? hideNewRows1El.checked : false; 
    const hideNewRows2 = hideNewRows2El ? hideNewRows2El.checked : false; 
    
    
    const { data1: alignedData1, data2: alignedData2, columnInfo } = prepareDataForComparison(data1, data2);
    
    
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
    
    
    const workingData1 = alignedData1;
    const workingData2 = alignedData2;
    
    
    let header1 = workingData1[0] || [];
    let header2 = workingData2[0] || [];
    
    
    const excludedIndexes1 = getExcludedColumnIndexes(header1, excludedColumns);
    const excludedIndexes2 = getExcludedColumnIndexes(header2, excludedColumns);
    
    
    const filteredHeader1 = filterRowExcludingColumns(header1, excludedIndexes1);
    const filteredHeader2 = filterRowExcludingColumns(header2, excludedIndexes2);
    
    
    let allCols = Math.max(filteredHeader1.length, filteredHeader2.length);
    let headers = (filteredHeader1.length >= filteredHeader2.length) ? filteredHeader1 : filteredHeader2;
    
    
    const finalHeaders = headers;
    const finalAllCols = finalHeaders.length;
    
    
    let body1 = workingData1.slice(1).map(row => filterRowExcludingColumns(row, excludedIndexes1));
    let body2 = workingData2.slice(1).map(row => filterRowExcludingColumns(row, excludedIndexes2));
    
    
    
    let onlyInFile1, onlyInFile2;
    
    if (columnInfo && (columnInfo.hasCommonColumns || columnInfo.onlyInFile1 || columnInfo.onlyInFile2)) {
        
        onlyInFile1 = columnInfo.onlyInFile1 || [];
        onlyInFile2 = columnInfo.onlyInFile2 || [];
    } else {
        
        let header1Set = new Set(filteredHeader1.map(h => h ? h.toString().trim().toLowerCase() : ''));
        let header2Set = new Set(filteredHeader2.map(h => h ? h.toString().trim().toLowerCase() : ''));
        
        onlyInFile1 = filteredHeader1.filter(h => h && h.toString().trim() !== '' && !header2Set.has(h.toString().trim().toLowerCase()));
        onlyInFile2 = filteredHeader2.filter(h => h && h.toString().trim() !== '' && !header1Set.has(h.toString().trim().toLowerCase()));
    }
    
    let diffColumns1 = onlyInFile1.length > 0 ? onlyInFile1.join(', ') : '-';
    let diffColumns2 = onlyInFile2.length > 0 ? onlyInFile2.join(', ') : '-';
    
    
    currentDiffColumns1 = diffColumns1;
    currentDiffColumns2 = diffColumns2;
    
    
    let diffColumns1Html = onlyInFile1.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns1}</span>` : diffColumns1;
    let diffColumns2Html = onlyInFile2.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns2}</span>` : diffColumns2;
    
    
    
    const loadingDiv = document.getElementById('comparison-loading');
    if (loadingDiv) {
        loadingDiv.remove();
    }

    function rowKey(row) { 
        if (toleranceMode) {
            
            
            return JSON.stringify(row.map(x => (x !== undefined ? x.toString() : ''))); 
        } else {
            return JSON.stringify(row.map(x => (x !== undefined ? x.toString().toUpperCase() : ''))); 
        }
    }
    
    let set1, set2, only1, only2, both;
    
    if (toleranceMode) {
        
        
        set1 = new Set(body1.map((row, index) => `row1_${index}`));
        set2 = new Set(body2.map((row, index) => `row2_${index}`));
        only1 = 0; 
        only2 = 0; 
        both = 0; 
    } else {
        set1 = new Set(body1.map(rowKey));
        set2 = new Set(body2.map(rowKey));
        only1 = 0; only2 = 0; both = 0;
        set1.forEach(k => { if (set2.has(k)) both++; else only1++; });
        set2.forEach(k => { if (!set1.has(k)) only2++; });
    }
    
    let totalRows = Math.max(data1.length, data2.length) - 1;
    let maxRows = Math.max(body1.length, body2.length);
    
    
    
    let totalUniqueRows = only1 + only2 + both;
    let differentRows = only1 + only2;
    let percentDiff, percentClass;
    
    if (toleranceMode && totalUniqueRows === 0) {
        
        percentDiff = tableHeaders.calculating;
        percentClass = 'percent-low';
    } else {
        
        const maxFileSize = Math.max(body1.length, body2.length);
        percentDiff = maxFileSize > 0 ? Math.min(((both / maxFileSize) * 100), 100).toFixed(2) + '%' : '0.00%';
        percentClass = 'percent-high';
        if (parseFloat(percentDiff) < 30) percentClass = 'percent-low';
        else if (parseFloat(percentDiff) < 70) percentClass = 'percent-medium';
        else percentClass = 'percent-high';
    }
    
    
    let excludedInfo = '';
    if (excludedColumns.length > 0) {
        excludedInfo = `<div class="excluded-info">
            <strong>Excluded from comparison:</strong> ${excludedColumns.join(', ')}
        </div>`;
    }
    
    
    let toleranceInfo = '';
    if (toleranceMode) {
        
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
    
    
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'block';
    }
    
    
    const exportBtn = document.getElementById('exportExcelBtn');
    const buttonsContainer = document.querySelector('.buttons-container');
    const exportButtonHalf = exportBtn ? exportBtn.closest('.button-half') : null;
    
    if (exportBtn && buttonsContainer) {
        exportBtn.style.display = 'inline-block';
        
        if (exportButtonHalf) {
            exportButtonHalf.classList.remove('export-hidden');
        }
        buttonsContainer.classList.remove('export-hidden');
    }
    
    
    if (hideSameRowsCheckbox) {
        
        hideSameRowsCheckbox.checked = false;
    }
    
    
    const totalRowsForCheck = Math.max(body1.length, body2.length);
    const isLargeFile = totalRowsForCheck > DETAILED_TABLE_LIMIT;
    
    if (isLargeFile) {
        
        performFuzzyMatchingForExport(body1, body2, finalHeaders, finalAllCols, true, tableHeaders);
        return;
    } else {
        
        performFuzzyMatchingForExport(body1, body2, finalHeaders, finalAllCols, false, tableHeaders);
    }
}


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
    
    
    if (toleranceMode) {
        
        window.summaryStats = {
            only1: onlyInFile1,
            only2: onlyInFile2,
            both: identicalCount + toleranceCount, 
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
    
    
    
    const maxFileSize = file1Size && file2Size ? Math.max(file1Size, file2Size) : (window.summaryStats.only1 + window.summaryStats.only2 + window.summaryStats.both);
    const percentDiff = maxFileSize > 0 ? Math.min(((window.summaryStats.both / maxFileSize) * 100), 100).toFixed(2) : '0.00';
    
    let percentClass = 'percent-high';
    if (parseFloat(percentDiff) < 30) percentClass = 'percent-low';
    else if (parseFloat(percentDiff) < 70) percentClass = 'percent-medium';
    else percentClass = 'percent-high';
    
    
    updateSummaryTable(window.summaryStats.only1, window.summaryStats.only2, window.summaryStats.both, percentDiff, percentClass, tableHeaders);
}


function updateSummaryTable(only1, only2, both, percentDiff, percentClass, tableHeaders) {
    
    const summaryDiv = document.getElementById('summary');
    const existingHTML = summaryDiv.innerHTML;
    
    
    const excludedMatch = existingHTML.match(/<div class="excluded-info">.*?<\/div>/);
    const toleranceMatch = existingHTML.match(/<div style="background: #fff3cd;.*?<\/div>/);
    
    const excludedInfo = excludedMatch ? excludedMatch[0] : '';
    const toleranceInfo = toleranceMatch ? toleranceMatch[0] : '';
    
    
    const diffColumns1 = currentDiffColumns1 || '-';
    const diffColumns2 = currentDiffColumns2 || '-';
    
    
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


function smartDetectKeyColumns(headers, data) {
    if (!headers || !data || data.length < 2) {
        return [0]; 
    }
    
    const columnCount = headers.length;
    const bodyData = data.slice(1); 
    
    
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
        
        
        const columnValues = bodyData.map(row => (row[colIndex] || '').toString().trim()).filter(val => val !== '');
        const uniqueValues = new Set(columnValues);
        const uniquenessRatio = columnValues.length > 0 ? uniqueValues.size / columnValues.length : 0;
        
        
        let isAggregationField = false;
        let aggregationConfidence = 0;
        
        if (columnValues.length > 0) {
            
            const yyyymmPattern = /^20\d{4}$/; 
            const yyyymmCount = columnValues.filter(val => yyyymmPattern.test(val)).length;
            const hasYYYYMMPattern = yyyymmCount > columnValues.length * 0.7; 
            
            if (hasYYYYMMPattern) {
                
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
            
            
            const aggregationPatterns = [
                { pattern: /^20\d{2}$/, type: 'year' },          
                { pattern: /^[1-9]$|^1[0-2]$/, type: 'month' },  
                { pattern: /^[1-4]$/, type: 'quarter' },         
                { pattern: /^20\d{2}-[01]\d$/, type: 'year-month' }, 
                { pattern: /^Q[1-4]$/, type: 'quarter' }         
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
            
            const baseScore = uniquenessRatio >= 0.3 ? 9 : 7;
            uniquenessScore = Math.floor(baseScore * aggregationConfidence);
        } else if (uniquenessRatio >= 0.95) {
            uniquenessScore = 10; 
        } else if (uniquenessRatio >= 0.8) {
            uniquenessScore = 8;
        } else if (uniquenessRatio >= 0.6) {
            uniquenessScore = 6;
        } else if (uniquenessRatio >= 0.4) {
            uniquenessScore = 4;
        } else {
            uniquenessScore = 1; 
        }
        
        score += uniquenessScore * 0.4;
        
        
        const positionScore = Math.max(1, 10 - colIndex * 2); 
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
    
    
    columnScores.sort((a, b) => b.score - a.score);
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    let keyColumns = [];
    
    
    const aggregationFields = columnScores.filter(col => col.isAggregationField);
    if (aggregationFields.length > 0) {
        
        
        
        keyColumns.push(aggregationFields[0].index);
        
        
        for (let i = 1; i < Math.min(3, aggregationFields.length); i++) {
            const col = aggregationFields[i];
            if (col.score >= 8 && col.aggregationConfidence >= 0.7) {
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
            
            if (col.score >= 6 && col.uniquenessRatio >= 0.7) {
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


function performFuzzyMatchingForExport(body1, body2, finalHeaders, finalAllCols, isLargeFile, tableHeaders) {
    
    
    const combinedData = [finalHeaders, ...body1, ...body2];
    const keyColumnIndexes = smartDetectKeyColumns(finalHeaders, combinedData);
    
    
    let used2 = new Array(body2.length).fill(false);
    let pairs = [];
    
    function countMatches(rowA, rowB) {
        let matches = 0;
        let keyMatches = 0; 
        let toleranceMatches = 0; 
        let keyToleranceMatches = 0; 
        
        
        const keyColumnsIndexes = keyColumnIndexes;
        
        for (let i = 0; i < finalAllCols; i++) {
            let valueA = (rowA[i] || '').toString();
            let valueB = (rowB[i] || '').toString();
            
            if (toleranceMode) {
                
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
                
                if (valueA.toUpperCase() === valueB.toUpperCase()) {
                    matches++;
                    if (keyColumnsIndexes.includes(i)) {
                        keyMatches++;
                    }
                }
            }
        }
        
        if (toleranceMode) {
            
            
            const adjustedToleranceMatches = toleranceMatches * 0.7;
            const adjustedKeyToleranceMatches = keyToleranceMatches * 0.7;
            
            const totalKeyScore = (keyMatches * 3) + (adjustedKeyToleranceMatches * 3);
            const totalOtherScore = (matches - keyMatches) + (adjustedToleranceMatches - adjustedKeyToleranceMatches);
            
            return totalKeyScore + totalOtherScore;
        } else {
            
            const otherMatches = matches - keyMatches;
            const weightedScore = (keyMatches * 3) + otherMatches;
            return weightedScore;
        }
    }
    
    
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
            
            const keyColumnsCount = keyColumnIndexes.length;
            let minKeyMatches = Math.ceil(keyColumnsCount * 0.6) * 3; 
            let minTotalMatches = Math.ceil(finalAllCols * 0.5); 
            
            
            if (toleranceMode) {
                minKeyMatches = Math.ceil(keyColumnsCount * 0.8) * 3; 
                minTotalMatches = Math.ceil(finalAllCols * 0.7); 
            }
            
            if (bestScore >= minKeyMatches || bestScore >= minTotalMatches) {
                pairs.push({row1: body1[i], row2: body2[bestIdx]});
                used2[bestIdx] = true;
            } else {
                pairs.push({row1: body1[i], row2: null});
            }
        }
        
        
        if (endIndex < body1.length) {
            
            if (body1.length > 1000) {
                const progress = Math.round((endIndex / body1.length) * 100);
                const progressMessage = isLargeFile ? 
                    `🔄 Processing large files for export... ${progress}% complete` :
                    `🔄 Comparing files... ${progress}% complete`;
                document.getElementById('result').innerHTML = `<div style="text-align: center; padding: 20px; font-size: 16px;">${progressMessage}</div>`;
            }
            setTimeout(() => processBatch(endIndex, batchSize), 10);
        } else {
            
            for (let j = 0; j < body2.length; j++) {
                if (!used2[j]) {
                    pairs.push({row1: null, row2: body2[j]});
                }
            }
            
            
            currentPairs = pairs;
            currentFinalHeaders = finalHeaders;
            currentFinalAllCols = finalAllCols;
            
            
            updateSummaryStatistics(tableHeaders, body1.length, body2.length);
            
            if (isLargeFile) {
                
                showLargeFileMessage(Math.max(body1.length, body2.length));
            } else {
                
                renderComparisonTable();
            }
        }
    }
    
    
    const batchSize = body1.length > 5000 ? 100 : body1.length > 1000 ? 250 : 1000;
    processBatch(0, batchSize);
}


function isDateString(str) {
    if (!str || typeof str !== 'string') return false;
    
    
    const cleanStr = str.replace(/['"]/g, '').trim();
    
    
    const datePatterns = [
        /^\d{4}-\d{2}-\d{2}/, 
        /^\d{2}\/\d{2}\/\d{4}/, 
        /^\d{2}\.\d{2}\.\d{4}/, 
        /^\d{1,2}\/\d{1,2}\/\d{4}/, 
        /^\d{1,2}-\d{1,2}-\d{4}/, 
    ];
    
    return datePatterns.some(pattern => pattern.test(cleanStr));
}

function isNumericString(str) {
    if (!str || typeof str !== 'string') return false;
    
    
    const cleanStr = str.replace(/['",$\s]/g, '').trim();
    
    
    return !isNaN(cleanStr) && !isNaN(parseFloat(cleanStr)) && isFinite(cleanStr);
}

function extractDateOnly(str) {
    if (!str) return '';
    
    const cleanStr = str.toString().replace(/['"]/g, '').trim();
    
    
    const dateMatch = cleanStr.match(/(\d{4}-\d{2}-\d{2}|\d{2}\/\d{2}\/\d{4}|\d{2}\.\d{2}\.\d{4}|\d{1,2}\/\d{1,2}\/\d{4}|\d{1,2}-\d{1,2}-\d{4})/);
    
    return dateMatch ? dateMatch[0] : cleanStr;
}

function parseNumber(str) {
    if (!str) return 0;
    
    
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
    
    
    if (str1.toUpperCase() === str2.toUpperCase()) {
        return 'identical';
    }
    
    
    if (isDateString(str1) && isDateString(str2)) {
        const date1 = extractDateOnly(str1);
        const date2 = extractDateOnly(str2);
        
        if (date1 === date2) {
            return 'tolerance'; 
        }
    }
    
    
    if (isNumericString(str1) && isNumericString(str2)) {
        if (isWithinTolerance(str1, str2, 0.015)) {
            return 'tolerance'; 
        }
    }
    
    return 'different';
}


function renderComparisonTable() {
    if (!currentPairs || currentPairs.length === 0) {
        document.querySelector('.diff-table-header thead').innerHTML = '';
        document.querySelector('.filter-row').innerHTML = '';
        document.querySelector('.diff-table-body tbody').innerHTML = '<tr><td colspan="100" style="text-align:center; padding:20px;">No data to display</td></tr>';
        return;
    }
    
    
    const hideSameEl = document.getElementById('hideSameRows');
    const hideDiffEl = document.getElementById('hideDiffColumns');
    const hideNewRows1El = document.getElementById('hideNewRows1');
    const hideNewRows2El = document.getElementById('hideNewRows2');
    
    let hideSame = hideSameEl ? hideSameEl.checked : false;
    const hideDiffRows = hideDiffEl ? hideDiffEl.checked : false;
    const hideNewRows1 = hideNewRows1El ? hideNewRows1El.checked : false;
    const hideNewRows2 = hideNewRows2El ? hideNewRows2El.checked : false;
    
    
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
    
    
    let filterHtml = '<tr><td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>';
    for (let c = 0; c < currentFinalAllCols; c++) {
        filterHtml += `<td><input type="text" placeholder="Filter..." onkeyup="filterTable()"></td>`;
    }
    filterHtml += '</tr>';
    document.querySelector('.filter-row').innerHTML = filterHtml;
    
    
    let visibleRowCount = 0;
    let tempBodyHtml = '';
    
    currentPairs.forEach(pair => {
        let row1 = pair.row1;
        let row2 = pair.row2;
        
        
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

    
    let bodyHtml = '';
    currentPairs.forEach(pair => {
        let row1 = pair.row1;
        let row2 = pair.row2;
        
        
        let row1Upper = row1 ? row1.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        let row2Upper = row2 ? row2.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
        
        let isEmpty = true;
        let hasWarn = false;
        let hasTolerance = false;
        let allSame = true;
        let hasDiff = false;
        
        
        let columnComparisons = [];
        
        for (let c = 0; c < currentFinalAllCols; c++) {
            let v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
            let v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
            
            if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                isEmpty = false;
            }
            
            if (row1 && row2) {
                if (toleranceMode) {
                    
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
        
        
        if (row1 && row2 && (hasWarn || hasTolerance)) {
            
            bodyHtml += `<tr class="warn-row warn-row-group-start">`;
            bodyHtml += `<td class="warn-cell">${getFileDisplayName(fileName1, sheetName1) || 'File 1'}</td>`;
            for (let c = 0; c < currentFinalAllCols; c++) {
                let v1 = row1[c] !== undefined ? row1[c] : '';
                let compResult = columnComparisons[c];
                
                if (compResult === 'identical') {
                    
                    bodyHtml += `<td class="identical" rowspan="2" style="vertical-align: middle; text-align: center;">${v1}</td>`;
                } else if (compResult === 'tolerance') {
                    
                    bodyHtml += `<td class="tolerance-cell">${v1}</td>`;
                } else {
                    
                    bodyHtml += `<td class="warn-cell">${v1}</td>`;
                }
            }
            bodyHtml += '</tr>';
            
            
            bodyHtml += `<tr class="warn-row warn-row-group-end">`;
            bodyHtml += `<td class="warn-cell">${getFileDisplayName(fileName2, sheetName2) || 'File 2'}</td>`;
            for (let c = 0; c < currentFinalAllCols; c++) {
                let v2 = row2[c] !== undefined ? row2[c] : '';
                let compResult = columnComparisons[c];
                
                if (compResult === 'tolerance') {
                    
                    bodyHtml += `<td class="tolerance-cell">${v2}</td>`;
                } else if (compResult === 'different') {
                    
                    bodyHtml += `<td class="warn-cell">${v2}</td>`;
                }
                
            }
            bodyHtml += '</tr>';
        } else {
            
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
                    
                    cellValue = row1[c] !== undefined ? row1[c] : '';
                    cellClass = 'identical';
                }
                
                bodyHtml += `<td class="${cellClass}">${cellValue}</td>`;
            }
            bodyHtml += '</tr>';
        }
    });
    
    document.querySelector('.diff-table-body tbody').innerHTML = bodyHtml;
    
    
    setTimeout(() => {
        
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
            
        }
        
        if (bodyTable) {
            bodyTable.style.borderCollapse = 'separate';
            bodyTable.style.borderSpacing = '0';
            bodyTable.style.tableLayout = 'fixed';
            
        }
        
        
        
        syncColumnWidths();
        forceTableWidthSync();
        syncTableScroll();
    }, 50);
}


function sortTable(column) {
    if (currentSortColumn === column) {
        currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortColumn = column;
        currentSortDirection = 'asc';
    }
    
    
    currentPairs.sort((a, b) => {
        let aVal = '';
        let bVal = '';
        
        
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
    
    
    renderComparisonTable();
}


function filterTable() {
    let filters = [];
    let filterInputs = document.querySelectorAll('.filter-row input[type="text"]');
    filterInputs.forEach(input => {
        filters.push(input.value.toLowerCase());
    });
    
    let rows = document.querySelectorAll('.diff-table-body tbody tr');
    
    
    for (let i = 0; i < rows.length; i++) {
        let row = rows[i];
        let cells = row.querySelectorAll('td');
        
        
        let hasRowspan = Array.from(cells).some(cell => cell.hasAttribute('rowspan'));
        
        if (hasRowspan && i + 1 < rows.length) {
            
            let nextRow = rows[i + 1];
            let show = false;
            
            
            show = checkRowMatchesFilters(row, filters) || checkRowMatchesFilters(nextRow, filters, row);
            
            
            row.style.display = show ? '' : 'none';
            nextRow.style.display = show ? '' : 'none';
            
            i++; 
        } else {
            
            let show = checkRowMatchesFilters(row, filters);
            row.style.display = show ? '' : 'none';
        }
    }
    
    
    refreshTableLayout();
}


function checkRowMatchesFilters(row, filters, rowspanRow = null) {
    let cells = row.querySelectorAll('td');
    
    for (let j = 0; j < filters.length; j++) {
        if (filters[j] && filters[j].trim() !== '') {
            let cellText = '';
            
            if (j < cells.length) {
                cellText = cells[j].textContent.toLowerCase();
            } else if (rowspanRow) {
                
                let rowspanCells = rowspanRow.querySelectorAll('td');
                
                
                let adjustedIndex = j;
                let currentCellIndex = 0;
                
                for (let k = 0; k <= j && currentCellIndex < rowspanCells.length; k++) {
                    if (k === j) {
                        if (rowspanCells[currentCellIndex] && rowspanCells[currentCellIndex].hasAttribute('rowspan')) {
                            
                            cellText = rowspanCells[currentCellIndex].textContent.toLowerCase();
                        }
                        break;
                    }
                    
                    
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


function refreshTableLayout() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    const headerContainer = document.querySelector('.table-header-fixed');
    const bodyContainer = document.querySelector('.table-body-scrollable');
    
    if (!headerTable || !bodyTable) return;
    
    
    headerTable.style.tableLayout = 'auto';
    bodyTable.style.tableLayout = 'auto';
    
    
    headerTable.offsetHeight;
    bodyTable.offsetHeight;
    
    
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    
    setTimeout(() => {
        syncColumnWidths();
        forceTableWidthSync();
        
        
        if (headerContainer && bodyContainer) {
            headerContainer.scrollLeft = bodyContainer.scrollLeft;
        }
    }, 10);
}


function adjustTableWidth() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    
    if (headerTable && bodyTable) {
        
        const headerCells = headerTable.querySelectorAll('th');
        const numColumns = headerCells.length;
        
        if (numColumns > 0) {
            
            const sourceColumnWidth = 240;
            const dataColumnWidth = 180;
            const totalWidth = sourceColumnWidth + ((numColumns - 1) * dataColumnWidth);
            
            
            const widthStyle = totalWidth + 'px';
            headerTable.style.width = widthStyle;
            bodyTable.style.width = widthStyle;
            headerTable.style.minWidth = widthStyle;
            bodyTable.style.minWidth = widthStyle;
            
            
            setTimeout(() => {
                syncColumnWidths();
            }, 10);
        }
    }
}


function getMemoryUsage() {
    if (performance && performance.memory) {
        return {
            used: Math.round(performance.memory.usedJSHeapSize / 1048576), 
            total: Math.round(performance.memory.totalJSHeapSize / 1048576), 
            limit: Math.round(performance.memory.jsHeapSizeLimit / 1048576) 
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


function parseCSVChunked(csvText, chunkSize = 1000) {
    const lines = csvText.split(/\r?\n/);
    const result = [];
    const delimiter = detectCSVDelimiter(csvText);
    
    
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
            
            setTimeout(() => processChunk(endIndex), 10);
        }
    }
    
    return new Promise((resolve) => {
        processChunk(0);
        
        
        resolve(parseCSV(csvText));
    });
}


window.addEventListener('load', function() {
    syncTableScroll();
    showPlaceholderMessage(); 
});


document.addEventListener('DOMContentLoaded', function() {
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

function syncColumnWidths() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    
    if (!headerTable || !bodyTable) return;
    
    const headerCells = headerTable.querySelectorAll('th');
    const allBodyRows = bodyTable.querySelectorAll('tbody tr');
    
    if (!headerCells.length) return;
    
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    const numColumns = headerCells.length;
    
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    const sourceColumnWidth = 240;
    const dataColumnWidth = 180;
    const calculatedTotalWidth = sourceColumnWidth + ((numColumns - 1) * dataColumnWidth);
    
    
    headerTable.style.width = calculatedTotalWidth + 'px';
    bodyTable.style.width = calculatedTotalWidth + 'px';
    headerTable.style.minWidth = calculatedTotalWidth + 'px';
    bodyTable.style.minWidth = calculatedTotalWidth + 'px';
}

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

function forceTableWidthSync() {
    const headerTable = document.querySelector('.diff-table-header');
    const bodyTable = document.querySelector('.diff-table-body');
    
    if (!headerTable || !bodyTable) return;
    
    const headerCells = headerTable.querySelectorAll('th');
    const allBodyRows = bodyTable.querySelectorAll('tbody tr');
    
    if (headerCells.length === 0) return;
    
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    const numColumns = headerCells.length;
    
    headerTable.style.tableLayout = 'fixed';
    bodyTable.style.tableLayout = 'fixed';
    
    const sourceColumnWidth = 240;
    const dataColumnWidth = 180;
    const calculatedTotalWidth = sourceColumnWidth + ((numColumns - 1) * dataColumnWidth);
    
    headerTable.style.width = calculatedTotalWidth + 'px';
    bodyTable.style.width = calculatedTotalWidth + 'px';
    headerTable.style.minWidth = calculatedTotalWidth + 'px';
    bodyTable.style.minWidth = calculatedTotalWidth + 'px';
}


function showLargeFileMessage(totalRows) {
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
    
    
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
    
    
    const exportBtn = document.getElementById('exportExcelBtn');
    const buttonsContainer = document.querySelector('.buttons-container');
    const exportButtonHalf = exportBtn ? exportBtn.closest('.button-half') : null;
    
    if (exportBtn && buttonsContainer) {
        exportBtn.style.display = 'inline-block';
        exportBtn.style.fontSize = '16px';
        exportBtn.style.padding = '12px 24px';
        exportBtn.style.backgroundColor = '#28a745';
        exportBtn.style.fontWeight = 'bold';
        
        
        if (exportButtonHalf) {
            exportButtonHalf.classList.remove('export-hidden');
        }
        buttonsContainer.classList.remove('export-hidden');
    }
}
