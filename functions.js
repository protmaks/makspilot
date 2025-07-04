// Excel and CSV File Comparison Functions
let data1 = [], data2 = [];
let fileName1 = '', fileName2 = '';

// Global variables for sorting
let currentSortColumn = -1;
let currentSortDirection = 'asc';
let currentPairs = [];
let currentFinalHeaders = [];
let currentFinalAllCols = 0;

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
                // Check if number could be an Excel date (1900-2100 range)
                if (value >= 36000 && value <= 73050 && value === Math.floor(value)) {
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
           (potentialDateRatio > 0.7 && numberCount > 0) || 
           combinedDateRatio > 0.6;
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
            if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
    }
    
    // Check if value is a number that could be an Excel date serial number
    if (typeof value === 'number' && value > 1 && value < 100000) {
        // Only convert numbers to dates if we're in a date column context
        // and the number is in a reasonable Excel date range
        if (isInDateColumn && 
            value >= 36000 && // Around 1998 - more reasonable starting point for dates
            value <= 73050 && // Approximately year 2100
            value === Math.floor(value)) { // Only whole numbers (dates don't have fractional parts)
            
            // Excel epoch starts from January 1, 1900 (but Excel incorrectly treats 1900 as a leap year)
            // Use UTC to avoid timezone issues
            const excelEpochUTC = Date.UTC(1899, 11, 30); // December 30, 1899 UTC
            const dateUTC = new Date(excelEpochUTC + value * 24 * 60 * 60 * 1000);
            
            // Ensure we're working with UTC values
            const year = dateUTC.getUTCFullYear();
            const month = dateUTC.getUTCMonth() + 1;
            const day = dateUTC.getUTCDate();
            
            if (year >= 1998 && year <= 2100) {
                // Format as YYYY-MM-DD
                return year + '-' + 
                       String(month).padStart(2, '0') + '-' + 
                       String(day).padStart(2, '0');
            }
        }
    }
    
    // Check if value is already a date string in various formats
    if (typeof value === 'string') {
        // First check if it's already in YYYY-MM-DD format - keep it as is
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
        
        // Check if it's a Date object converted to string that might have timezone issues
        // This can happen when CSV parsing creates Date objects that get stringified
        if (value.includes('T') || value.includes('GMT') || value.includes('UTC')) {
            try {
                // For ISO strings, extract date part directly to avoid timezone conversion
                if (value.includes('T')) {
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
                
                // Fallback: Parse as Date and extract just the date part in local timezone
                const dateObj = new Date(value);
                if (!isNaN(dateObj.getTime())) {
                    // Use local date components to avoid timezone shifts
                    const year = dateObj.getFullYear();
                    const month = dateObj.getMonth() + 1;
                    const day = dateObj.getDate();
                    if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
                        return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
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
        
        // DD/MM/YYYY format (slashes as separators)
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const month = parseInt(dateMatch[2]);
            const year = parseInt(dateMatch[3]);
            if (day <= 31 && month <= 12 && year >= 1900 && year <= 2100) {
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            }
        }
        
        // MM/DD/YYYY format (American format)
        dateMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dateMatch) {
            const first = parseInt(dateMatch[1]);
            const second = parseInt(dateMatch[2]);
            const year = parseInt(dateMatch[3]);
            
            // Smart interpretation: if first > 12, it's likely DD/MM, not MM/DD
            if (first > 12 && second <= 12 && year >= 1900 && year <= 2100) {
                // Definitely DD/MM format
                const day = first;
                const month = second;
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            } else if (first <= 12 && second <= 12 && year >= 1900 && year <= 2100) {
                // Ambiguous - could be either MM/DD or DD/MM
                // For European context, assume DD/MM is more likely
                const day = first;
                const month = second;
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            } else {
                // Standard MM/DD interpretation
                const month = first;
                const day = second;
                if (day <= 31 && month <= 12 && year >= 1900 && year <= 2100) {
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
            
            // Smart interpretation: if first > 12, it's likely DD/MM, not MM/DD
            if (first > 12 && second <= 12) {
                // Definitely DD/MM format
                const day = first;
                const month = second;
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            } else if (first <= 12 && second <= 12) {
                // Ambiguous - could be either MM/DD or DD/MM
                // For European context, assume DD/MM is more likely
                const day = first;
                const month = second;
                return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
            } else {
                // Standard MM/DD interpretation
                const month = first;
                const day = second;
                if (day <= 31 && month <= 12) {
                    return year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
                }
            }
        }
        
        // DD.MM.YY format (short year)
        dateMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2})$/);
        if (dateMatch) {
            const day = parseInt(dateMatch[1]);
            const month = parseInt(dateMatch[2]);
            let year = parseInt(dateMatch[3]);
            // Assume years 00-30 are 2000-2030, years 31-99 are 1931-1999
            year = year <= 30 ? 2000 + year : 1900 + year;
            if (day <= 31 && month <= 12) {
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
    let maxCols = Math.max(...data.map(row => row ? row.length : 0));
    if (maxCols === 0) return data;
    
    // Find columns that are completely empty
    let columnsToKeep = [];
    for (let col = 0; col < maxCols; col++) {
        let hasContent = false;
        for (let row = 0; row < data.length; row++) {
            if (data[row] && col < data[row].length) {
                let cellValue = data[row][col];
                if (cellValue !== null && cellValue !== undefined && cellValue.toString().trim() !== '') {
                    hasContent = true;
                    break;
                }
            }
        }
        if (hasContent) {
            columnsToKeep.push(col);
        }
    }
    
    // Create new rows with only non-empty columns, and normalize dates intelligently
    let result = data.map(row => {
        if (!row) return [];
        return columnsToKeep.map(colIndex => 
            colIndex < row.length ? row[colIndex] : ''
        );
    });
    
    // Now normalize dates column by column
    if (result.length > 0) {
        const numCols = result[0].length;
        
        for (let col = 0; col < numCols; col++) {
            // Extract column values for date detection
            const columnValues = result.map(row => row[col]);
            const columnHeader = result.length > 0 ? result[0][col] : ''; // Use first row as header
            const isDateCol = isDateColumn(columnValues, columnHeader);
            
            // Normalize each value in the column
            for (let row = 0; row < result.length; row++) {
                if (result[row][col] !== null && result[row][col] !== undefined && result[row][col] !== '') {
                    result[row][col] = convertExcelDate(result[row][col], isDateCol);
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
    const maxCols = Math.max(...data.map(row => row ? row.length : 0));
    
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
    if (value.match(/^\d{4}-\d{1,2}-\d{1,2}$/) || 
        value.match(/^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{4}$/) ||
        value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/) ||
        value.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/)) {
        return value; // Keep dates as strings
    }
    
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
    if (num === 1) {
        fileName1 = file.name;
    } else {
        fileName2 = file.name;
    }
    
    // Check if it's a CSV file
    const isCSV = file.name.toLowerCase().endsWith('.csv');
    
    if (isCSV) {
        // Handle CSV files with text reading to avoid timezone issues
        const reader = new FileReader();
        reader.onload = function(e) {
            const csvText = e.target.result;
            const json = parseCSV(csvText);
            
            // Normalize row lengths first
            const normalizedJson = normalizeRowLengths(json);
            
            // Remove empty columns and normalize dates intelligently
            const cleanedJson = removeEmptyColumns(normalizedJson);
            
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
        };
        reader.readAsText(file, 'UTF-8');
    } else {
        // Handle Excel files with XLSX library
        const reader = new FileReader();
        reader.onload = function(e) {
            let data = new Uint8Array(e.target.result);
            // Use cellDates option to preserve date formatting and avoid timezone issues
            let workbook = XLSX.read(data, {
                type: 'array',
                cellDates: true,
                UTC: false  // Use local timezone to avoid shifts
            });
            let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            let json = XLSX.utils.sheet_to_json(firstSheet, {
                header: 1, 
                defval: '',
                raw: false,  // Don't use raw values, process them
                dateNF: 'yyyy-mm-dd'  // Prefer this date format
            });
            
            // Remove completely empty rows
            json = json.filter(row => Array.isArray(row) && row.some(cell => (cell !== null && cell !== undefined && cell.toString().trim() !== '')));
            
            // Remove completely empty columns and normalize dates intelligently
            json = removeEmptyColumns(json);
            
            // Round decimal numbers
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
        };
        reader.readAsArrayBuffer(file);
    }
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
        // First 3 data rows
        for (let i = 1; i < Math.min(4, data.length); i++) {
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
    document.getElementById('diffTable').innerHTML = `
        <div class="placeholder-message">
            <div class="placeholder-icon">📊</div>
            <div class="placeholder-text">Choose files on your computer and click Compare</div>
        </div>
    `;
    
    // Hide filter controls when showing placeholder
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'none';
    }
}

function restoreTableStructure() {
    document.getElementById('diffTable').innerHTML = `
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

function compareTables() {
    if (!data1.length || !data2.length) {
        document.getElementById('result').innerText = 'Please, load both files.';
        document.getElementById('summary').innerHTML = '';
        showPlaceholderMessage();
        return;
    }
    
    // Restore table structure for comparison results
    restoreTableStructure();
    
    // Get excluded columns
    const excludedColumns = getExcludedColumns();
    
    // Get filter options
    const hideDiffRows = document.getElementById('hideDiffColumns').checked; // Hide rows with differences
    const hideNewRows1 = document.getElementById('hideNewRows1').checked; // Hide rows only in file 1
    const hideNewRows2 = document.getElementById('hideNewRows2').checked; // Hide rows only in file 2
    
    // --- Statistics ---
    let header1 = data1[0] || [];
    let header2 = data2[0] || [];
    
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
    let body1 = data1.slice(1).map(row => filterRowExcludingColumns(row, excludedIndexes1));
    let body2 = data2.slice(1).map(row => filterRowExcludingColumns(row, excludedIndexes2));
    
    // Find columns that exist in one file but not in the other (from original filtered headers, before hide filters)
    let header1Set = new Set(filteredHeader1.map(h => h ? h.toString().trim() : ''));
    let header2Set = new Set(filteredHeader2.map(h => h ? h.toString().trim() : ''));
    
    let onlyInFile1 = filteredHeader1.filter(h => h && h.toString().trim() !== '' && !header2Set.has(h.toString().trim()));
    let onlyInFile2 = filteredHeader2.filter(h => h && h.toString().trim() !== '' && !header1Set.has(h.toString().trim()));
    
    let diffColumns1 = onlyInFile1.length > 0 ? onlyInFile1.join(', ') : '-';
    let diffColumns2 = onlyInFile2.length > 0 ? onlyInFile2.join(', ') : '-';
    
    // Apply red highlighting if there are diff columns
    let diffColumns1Html = onlyInFile1.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns1}</span>` : diffColumns1;
    let diffColumns2Html = onlyInFile2.length > 0 ? `<span style="color: red; font-weight: bold;">${diffColumns2}</span>` : diffColumns2;
    
    // Form keys for exact match with case-insensitive comparison
    function rowKey(row) { 
        return JSON.stringify(row.map(x => (x !== undefined ? x.toString().toUpperCase() : ''))); 
    }
    let set1 = new Set(body1.map(rowKey));
    let set2 = new Set(body2.map(rowKey));
    let only1 = 0, only2 = 0, both = 0;
    set1.forEach(k => { if (set2.has(k)) both++; else only1++; });
    set2.forEach(k => { if (!set1.has(k)) only2++; });
    // --- Top table ---
    let totalRows = Math.max(data1.length, data2.length) - 1;
    let maxRows = Math.max(body1.length, body2.length);
    let percentDiff = maxRows > 0 ? (((only1 + only2) / maxRows) * 100).toFixed(2) : '0.00';
    let percentClass = 'percent-low';
    if (parseFloat(percentDiff) > 30) percentClass = 'percent-high';
    else if (parseFloat(percentDiff) > 10) percentClass = 'percent-medium';
    
    // Add excluded columns info
    let excludedInfo = '';
    if (excludedColumns.length > 0) {
        excludedInfo = `<div class="excluded-info">
            <strong>Excluded from comparison:</strong> ${excludedColumns.join(', ')}
        </div>`;
    }
    let htmlSummary = `
        ${excludedInfo}
        <table style="margin-bottom:20px; border: 1px solid #ccc;">
            <tr><th>File</th><th>Row Count</th><th>Rows only in this file</th><th>Identical rows</th><th>% Difference</th><th>Diff columns</th></tr>
            <tr><td>${fileName1 || 'File 1'}</td><td>${body1.length}</td><td>${only1}</td><td rowspan="2">${both}</td><td rowspan="2" class="percent-cell ${percentClass}">${percentDiff}%</td><td>${diffColumns1Html}</td></tr>
            <tr><td>${fileName2 || 'File 2'}</td><td>${body2.length}</td><td>${only2}</td><td>${diffColumns2Html}</td></tr>
        </table>
    `;
    document.getElementById('summary').innerHTML = htmlSummary;
    
    // Show filter controls after comparison
    const filterControls = document.querySelector('.filter-controls');
    if (filterControls) {
        filterControls.style.display = 'block';
    }
    
    // --- Fuzzy matching for bottom table ---
    let used2 = new Array(body2.length).fill(false);
    let pairs = [];
    function countMatches(rowA, rowB) {
        let matches = 0;
        for (let i = 0; i < finalAllCols; i++) {
            // Convert both values to uppercase for case-insensitive comparison (cache the conversion)
            let valueA = (rowA[i] || '').toString();
            let valueB = (rowB[i] || '').toString();
            if (valueA.toUpperCase() === valueB.toUpperCase()) matches++;
        }
        return matches;
    }
    for (let i = 0; i < body1.length; i++) {
        let bestIdx = -1, bestScore = -1;
        for (let j = 0; j < body2.length; j++) {
            if (used2[j]) continue;
            let score = countMatches(body1[i], body2[j]);
            if (score > bestScore) {
                bestScore = score;
                bestIdx = j;
            }
        }
        if (bestScore >= Math.ceil(finalAllCols / 2)) {
            pairs.push({row1: body1[i], row2: body2[bestIdx]});
            used2[bestIdx] = true;
        } else {
            pairs.push({row1: body1[i], row2: null});
        }
    }
    for (let j = 0; j < body2.length; j++) {
        if (!used2[j]) {
            pairs.push({row1: null, row2: body2[j]});
        }
    }
    // Save pairs and headers for sorting
    currentPairs = pairs;
    currentFinalHeaders = finalHeaders;
    currentFinalAllCols = finalAllCols;
    // --- Bottom table with same rows filtering ---
    let hideSame = document.getElementById('hideSameRows').checked;
    let html = '';
    if (pairs.length > 0) {
        // Table headers
        let headerHtml = '<tr><th title="Source file">Source</th>';
        for (let c = 0; c < finalAllCols; c++) {
            let headerText = finalHeaders[c] !== undefined ? finalHeaders[c] : '';
            let titleAttr = headerText ? ` title="${headerText.toString().replace(/"/g, '&quot;')}"` : '';
            headerHtml += `<th class="sortable" onclick="sortTable(${c})"${titleAttr}>${headerText}</th>`;
        }
        headerHtml += '</tr>';
        document.querySelector('.diff-table-header thead').innerHTML = headerHtml;
        
        // Filter row
        let filterHtml = '<tr><td><input type="text" placeholder="Filter..." onkeyup="filterTable()" style="width:100%; padding:4px; border:1px solid #ccc; border-radius:3px;"></td>';
        for (let c = 0; c < finalAllCols; c++) {
            filterHtml += `<td><input type="text" placeholder="Filter..." onkeyup="filterTable()" style="width:100%; padding:4px; border:1px solid #ccc; border-radius:3px;"></td>`;
        }
        filterHtml += '</tr>';
        document.querySelector('.filter-row').innerHTML = filterHtml;
        
        // Table body
        let bodyHtml = '';
        pairs.forEach(pair => {
            let row1 = pair.row1;
            let row2 = pair.row2;
            
            // Pre-convert values to uppercase for performance
            let row1Upper = row1 ? row1.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
            let row2Upper = row2 ? row2.map(val => (val !== undefined ? val.toString().toUpperCase() : '')) : null;
            
            // Skip completely empty rows
            let isEmpty = true;
            let hasWarn = false;
            let allSame = true;
            for (let c = 0; c < finalAllCols; c++) {
                let v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
                let v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
                if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                    isEmpty = false;
                }
                if (row1 && row2) {
                    // Use pre-converted uppercase values
                    if (row1Upper[c] !== row2Upper[c]) {
                        hasWarn = true;
                        allSame = false;
                    }
                }
            }
            if (isEmpty) return;
            if (hideSame && row1 && row2 && allSame) return;
            
            // Apply additional row filters after computing hasWarn
            if (hideNewRows1 && row1 && !row2) return; // Skip rows only in file 1
            if (hideNewRows2 && !row1 && row2) return; // Skip rows only in file 2
            if (hideDiffRows && row1 && row2 && hasWarn) return; // Skip rows with differences (orange rows)
            
            let source = '';
            if (row1 && row2) {
                source = 'Both files';
            } else if (row1) {
                source = fileName1 || 'File 1';
            } else {
                source = fileName2 || 'File 2';
            }
            
            // Enhanced row class logic
            let rowClass = '';
            let hasDiff = !row1 || !row2;
            if (hasDiff) rowClass = 'row-diff';
            else if (hasWarn) rowClass = 'row-warn';
            else if (allSame) rowClass = 'row-identical';
            else rowClass = 'row-default';
            
            bodyHtml += `<tr class="${rowClass}">`;
            // Format source column similar to diff cells when there are differences
            if (row1 && row2 && (hasWarn || !allSame)) {
                bodyHtml += `<td><div>${fileName1 || 'File 1'}</div><div style='border-top:1px solid #eee;color:#555;font-size:90%'>${fileName2 || 'File 2'}</div></td>`;
            } else {
                bodyHtml += `<td>${source}</td>`;
            }
            for (let c = 0; c < finalAllCols; c++) {
                let v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
                let v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
                let cellClass = '';
                if (!row1 || !row2) cellClass = 'diff';
                else {
                    // Use pre-converted uppercase values for comparison
                    if (row1Upper[c] !== row2Upper[c]) cellClass = 'warn';
                    else cellClass = 'identical';
                }
                if (!row1 && row2) {
                    bodyHtml += `<td class="${cellClass}"><div>${v2}</div></td>`;
                } else if (row1 && !row2) {
                    bodyHtml += `<td class="${cellClass}"><div>${v1}</div></td>`;
                } else {
                    // Compare using pre-converted values but display original values
                    if (row1Upper[c] !== row2Upper[c]) {
                        bodyHtml += `<td class="${cellClass}"><div>${v1}</div><div style='border-top:1px solid #eee;color:#555;font-size:90%'>${v2}</div></td>`;
                    } else {
                        bodyHtml += `<td class="${cellClass}"><div>${v1}</div></td>`;
                    }
                }
            }
            bodyHtml += '</tr>';
        });
        document.querySelector('.diff-table-body tbody').innerHTML = bodyHtml;
        
        // Sync scroll after table generation
        setTimeout(() => {
            adjustTableWidth();
            syncTableScroll();
        }, 50);
    } else {
        document.querySelector('.diff-table-header thead').innerHTML = '';
        document.querySelector('.filter-row').innerHTML = '';
        document.querySelector('.diff-table-body tbody').innerHTML = '<tr><td colspan="100" style="text-align:center; padding:20px;">No different rows found</td></tr>';
    }
}

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
    
    // Update display
    renderSortedTable();
}

function renderSortedTable() {
    let hideSame = document.getElementById('hideSameRows').checked;
    const hideDiffRows = document.getElementById('hideDiffColumns').checked;
    const hideNewRows1 = document.getElementById('hideNewRows1').checked;
    const hideNewRows2 = document.getElementById('hideNewRows2').checked;
    
    if (currentPairs.length > 0) {
        // Table headers
        let headerHtml = '<tr><th title="Source file">Source</th>';
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
        
        // Filter row
        let filterHtml = '<tr><td><input type="text" placeholder="Filter..." onkeyup="filterTable()" style="width:100%; padding:4px; border:1px solid #ccc; border-radius:3px;"></td>';
        for (let c = 0; c < currentFinalAllCols; c++) {
            filterHtml += `<td><input type="text" placeholder="Filter..." onkeyup="filterTable()" style="width:100%; padding:4px; border:1px solid #ccc; border-radius:3px;"></td>`;
        }
        filterHtml += '</tr>';
        document.querySelector('.filter-row').innerHTML = filterHtml;
        
        // Table body
        let bodyHtml = '';
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
            let hasIdentical = false;
            
            for (let c = 0; c < currentFinalAllCols; c++) {
                let v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
                let v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
                if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                    isEmpty = false;
                }
                if (row1 && row2) {
                    // Use pre-converted uppercase values
                    if (row1Upper[c] !== row2Upper[c]) {
                        hasWarn = true;
                        allSame = false;
                    } else {
                        hasIdentical = true;
                    }
                } else {
                    hasDiff = true;
                    allSame = false;
                }
            }
            
            if (isEmpty) return;
            if (hideSame && row1 && row2 && allSame) return;
            
            // Apply additional row filters after computing hasWarn
            if (hideNewRows1 && row1 && !row2) return; // Skip rows only in file 1
            if (hideNewRows2 && !row1 && row2) return; // Skip rows only in file 2
            if (hideDiffRows && row1 && row2 && hasWarn) return; // Skip rows with differences (orange rows)
            
            let source = '';
            if (row1 && row2) {
                source = 'Both files';
            } else if (row1) {
                source = fileName1 || 'File 1';
            } else {
                source = fileName2 || 'File 2';
            }
            
            let rowClass = '';
            if (hasDiff) rowClass = 'row-diff';
            else if (hasWarn) rowClass = 'row-warn';
            else if (hasIdentical && allSame) rowClass = 'row-identical';
            else rowClass = 'row-default'; // Add default class
            
            bodyHtml += `<tr class="${rowClass}">`;
            // Format source column similar to diff cells when there are differences
            if (row1 && row2 && (hasWarn || !allSame)) {
                bodyHtml += `<td><div>${fileName1 || 'File 1'}</div><div style='border-top:1px solid #eee;color:#555;font-size:90%'>${fileName2 || 'File 2'}</div></td>`;
            } else {
                bodyHtml += `<td>${source}</td>`;
            }
            for (let c = 0; c < currentFinalAllCols; c++) {
                let v1 = row1 ? (row1[c] !== undefined ? row1[c] : '') : '';
                let v2 = row2 ? (row2[c] !== undefined ? row2[c] : '') : '';
                let cellClass = '';
                if (!row1 || !row2) cellClass = 'diff';
                else {
                    // Use pre-converted uppercase values for comparison
                    if (row1Upper[c] !== row2Upper[c]) cellClass = 'warn';
                    else cellClass = 'identical';
                }
                if (!row1 && row2) {
                    bodyHtml += `<td class="${cellClass}"><div>${v2}</div></td>`;
                } else if (row1 && !row2) {
                    bodyHtml += `<td class="${cellClass}"><div>${v1}</div></td>`;
                } else {
                    // Compare using pre-converted values but display original values
                    if (row1Upper[c] !== row2Upper[c]) {
                        bodyHtml += `<td class="${cellClass}"><div>${v1}</div><div style='border-top:1px solid #eee;color:#555;font-size:90%'>${v2}</div></td>`;
                    } else {
                        bodyHtml += `<td class="${cellClass}"><div>${v1}</div></td>`;
                    }
                }
            }
            bodyHtml += '</tr>';
        });
        document.querySelector('.diff-table-body tbody').innerHTML = bodyHtml;
        
        // Only sync scroll after rendering (don't adjust width on filtering)
        setTimeout(() => {
            syncTableScroll();
        }, 50);
    } else {
        document.querySelector('.diff-table-header thead').innerHTML = '';
        document.querySelector('.filter-row').innerHTML = '';
        document.querySelector('.diff-table-body tbody').innerHTML = '<tr><td colspan="100" style="text-align:center; padding:20px;">No different rows found</td></tr>';
    }
}

// Table filtering function
function filterTable() {
    let filters = [];
    let filterInputs = document.querySelectorAll('.filter-row input[type="text"]');
    filterInputs.forEach(input => {
        filters.push(input.value.toLowerCase());
    });
    
    let rows = document.querySelectorAll('.diff-table-body tbody tr');
    // Skip headers and filter row
    for (let i = 0; i < rows.length; i++) {
        let row = rows[i];
        let cells = row.querySelectorAll('td');
        let show = true;
        
        for (let j = 0; j < filters.length && j < cells.length; j++) {
            if (filters[j] && filters[j].trim() !== '') {
                let cellText = cells[j].textContent.toLowerCase();
                if (!cellText.includes(filters[j])) {
                    show = false;
                    break;
                }
            }
        }
        
        row.style.display = show ? '' : 'none';
    }
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
            // Calculate fixed total width: Source column (120px) + data columns (150px each)
            const sourceColumnWidth = 120;
            const dataColumnWidth = 150;
            const totalWidth = sourceColumnWidth + ((numColumns - 1) * dataColumnWidth);
            
            // Set consistent table widths
            const widthStyle = totalWidth + 'px';
            headerTable.style.width = widthStyle;
            bodyTable.style.width = widthStyle;
            headerTable.style.minWidth = widthStyle;
            bodyTable.style.minWidth = widthStyle;
        }
    }
}

// Synchronize header and body table scrolling
function syncTableScroll() {
    const headerContainer = document.querySelector('.table-header-fixed');
    const bodyContainer = document.querySelector('.table-body-scrollable');
    
    if (headerContainer && bodyContainer) {
        // Remove existing listeners to avoid duplicates
        headerContainer.removeEventListener('scroll', syncHeaderToBody);
        bodyContainer.removeEventListener('scroll', syncBodyToHeader);
        
        // Add new listeners
        bodyContainer.addEventListener('scroll', syncBodyToHeader);
        headerContainer.addEventListener('scroll', syncHeaderToBody);
    }
}

function syncBodyToHeader() {
    const headerContainer = document.querySelector('.table-header-fixed');
    const bodyContainer = document.querySelector('.table-body-scrollable');
    if (headerContainer && bodyContainer) {
        headerContainer.scrollLeft = bodyContainer.scrollLeft;
    }
}

function syncHeaderToBody() {
    const headerContainer = document.querySelector('.table-header-fixed');
    const bodyContainer = document.querySelector('.table-body-scrollable');
    if (headerContainer && bodyContainer) {
        bodyContainer.scrollLeft = headerContainer.scrollLeft;
    }
}

// Initialize synchronization and placeholder on page load
window.addEventListener('load', function() {
    syncTableScroll();
    showPlaceholderMessage(); // Show placeholder message on page load
});

// Initialize file handlers
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('file1').addEventListener('change', function(e) {
        handleFile(e.target.files[0], 1);
    });
    document.getElementById('file2').addEventListener('change', function(e) {
        handleFile(e.target.files[0], 2);
    });
    
    // Add event listeners for filter checkboxes
    document.getElementById('hideSameRows').addEventListener('change', function() {
        if (currentPairs && currentPairs.length > 0) {
            renderSortedTable();
        }
    });
    
    document.getElementById('hideDiffColumns').addEventListener('change', function() {
        if (currentPairs && currentPairs.length > 0) {
            renderSortedTable();
        }
    });
    
    document.getElementById('hideNewRows1').addEventListener('change', function() {
        if (currentPairs && currentPairs.length > 0) {
            renderSortedTable();
        }
    });
    
    document.getElementById('hideNewRows2').addEventListener('change', function() {
        if (currentPairs && currentPairs.length > 0) {
            renderSortedTable();
        }
    });
    
    // Initialize table with empty values on page load
    let htmlSummary = `
        <table style="margin-bottom:20px; border: 1px solid #ccc;">
            <tr><th>File</th><th>Row Count</th><th>Rows only in this file</th><th>Identical rows</th><th>% Difference</th><th>Diff columns</th></tr>
            <tr><td>-</td><td>-</td><td>-</td><td rowspan="2">-</td><td rowspan="2" class="percent-cell">-</td><td>-</td></tr>
            <tr><td>-</td><td>-</td><td>-</td><td>-</td></tr>
        </table>
    `;
    document.getElementById('summary').innerHTML = htmlSummary;
});
