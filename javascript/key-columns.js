// Key columns functionality for comparison

// Debounce timer for updateKeyColumnsOptions
let updateKeyColumnsTimer = null;

// Get column type based on data analysis
function getColumnType(columnValues, columnHeader = '') {
    if (!columnValues || columnValues.length === 0) return 'text';
    
    const nonEmptyValues = columnValues.filter(val => 
        val !== null && val !== undefined && val.toString().trim() !== ''
    );
    
    if (nonEmptyValues.length === 0) return 'text';
    
    let numberCount = 0;
    let dateCount = 0;
    let textCount = 0;
    
    const headerLower = columnHeader.toString().toLowerCase();
    const dateKeywords = ['date', 'time', 'created', 'modified', 'updated', 'birth', 'дата', 'время', 'создан', 'изменен', 'обновлен'];
    const numberKeywords = ['id', 'count', 'amount', 'price', 'cost', 'sum', 'total', 'number', 'номер', 'количество', 'сумма', 'цена', 'стоимость'];
    
    for (let value of nonEmptyValues) {
        const strValue = value.toString().trim().toLowerCase();
        
        // Check for numbers
        if (!isNaN(value) && !isNaN(parseFloat(value)) && isFinite(value)) {
            numberCount++;
            continue;
        }
        
        // Check for dates
        if (typeof value === 'string') {
            if (value.match(/^\d{4}-\d{1,2}-\d{1,2}$/) || 
                value.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/) || 
                value.match(/^\d{1,2}\.\d{1,2}\.\d{4}$/) ||
                value.match(/^\d{1,2}-\d{1,2}-\d{4}$/) ||
                value.includes('T') || value.includes('GMT') || value.includes('UTC')) {
                dateCount++;
                continue;
            }
        } else if (value instanceof Date) {
            dateCount++;
            continue;
        } else if (typeof value === 'number' && value >= 29221 && value <= 219146) {
            // Potential Excel serial date
            dateCount++;
            continue;
        }
        
        textCount++;
    }
    
    const total = nonEmptyValues.length;
    const numberRatio = numberCount / total;
    const dateRatio = dateCount / total;
    
    // Header-based hints
    const headerSuggestsDate = dateKeywords.some(keyword => headerLower.includes(keyword));
    const headerSuggestsNumber = numberKeywords.some(keyword => headerLower.includes(keyword));
    
    // Determine type based on ratios and header hints
    if (dateRatio > 0.6 || (headerSuggestsDate && dateRatio > 0.3)) {
        return 'date';
    }
    
    if (numberRatio > 0.7 || (headerSuggestsNumber && numberRatio > 0.5)) {
        return 'number';
    }
    
    return 'text';
}

// Helper function to extract data from table DOM element
function extractDataFromTableElement(tableElement) {
    try {
        if (!tableElement) return null;
        
        const table = tableElement.querySelector('table');
        if (!table) return null;
        
        const rows = table.querySelectorAll('tr');
        if (rows.length === 0) return null;
        
        const data = [];
        rows.forEach(row => {
            const cells = row.querySelectorAll('th, td');
            const rowData = Array.from(cells).map(cell => cell.textContent.trim());
            if (rowData.length > 0) {
                data.push(rowData);
            }
        });
        
        return data.length > 0 ? data : null;
    } catch (error) {
        console.warn('⚠️ Error extracting data from table element:', error);
        return null;
    }
}

// Debounced wrapper for updateKeyColumnsOptions
function debouncedUpdateKeyColumnsOptions(forceUpdate = false, delay = 150) {
    // Clear existing timer
    if (updateKeyColumnsTimer) {
        clearTimeout(updateKeyColumnsTimer);
    }
    
    // Set new timer
    updateKeyColumnsTimer = setTimeout(() => {
        updateKeyColumnsOptionsInternal(forceUpdate);
        updateKeyColumnsTimer = null;
    }, delay);
}

// Update key columns dropdown when files are loaded
function updateKeyColumnsOptionsInternal(forceUpdate = false) {
    const dropdownContent = document.getElementById('keyColumnsDropdownContent');
    const dropdownButton = document.getElementById('keyColumnsDropdownButton');
    
    // Reduced logging - only log errors
    if (!dropdownContent || !dropdownButton) {
        console.error('❌ Required DOM elements not found for key columns dropdown');
        return;
    }
    
    // Check if we already have checkboxes and data hasn't changed (unless force update)
    const existingCheckboxes = dropdownContent.querySelectorAll('input[name="keyColumns"]');
    if (!forceUpdate && existingCheckboxes.length > 0) {
        return; // Skip silently if not forced
    }
    
    // Save currently selected checkboxes before clearing
    const currentlySelected = [];
    const existingSelectedCheckboxes = dropdownContent.querySelectorAll('input[name="keyColumns"]:checked');
    existingSelectedCheckboxes.forEach(checkbox => {
        currentlySelected.push(checkbox.value);
    });
    
    // Try to load cached selections if no current selections or if forced update
    if (currentlySelected.length === 0 || forceUpdate) {
        const cachedSelections = loadKeyColumnsFromCache();
        if (cachedSelections && cachedSelections.length > 0) {
            // Only use cache if no current selections, or if forced update and cache is different
            if (currentlySelected.length === 0 || 
                (forceUpdate && JSON.stringify(currentlySelected.sort()) !== JSON.stringify(cachedSelections.sort()))) {
                currentlySelected.length = 0; // Clear current
                currentlySelected.push(...cachedSelections);
                //console.log('🔄 Restored key columns from cache:', cachedSelections);
            }
        } else if (forceUpdate && currentlySelected.length === 0) {
            // No cache found and no current selections - try auto-detection
            //console.log('🤖 No cache found, attempting auto-detection...');
        }
    }
        
    // Clear existing checkboxes
    dropdownContent.innerHTML = '';
    
    // Get headers from both files
    let allHeaders = new Set();
    
    // Try multiple ways to access the data
    let dataTable1 = null;
    let dataTable2 = null;
    
    // Method 1: Check global window variables
    if (window.data1 && window.data1.length > 0) {
        dataTable1 = window.data1;
    }
    if (window.data2 && window.data2.length > 0) {
        dataTable2 = window.data2;
    }
    
    // Method 2: Check local variables (if available)
    if (!dataTable1 && typeof data1 !== 'undefined' && data1 && data1.length > 0) {
        dataTable1 = data1;
    }
    if (!dataTable2 && typeof data2 !== 'undefined' && data2 && data2.length > 0) {
        dataTable2 = data2;
    }
    
    // Method 3: Try to find data in DOM tables (fallback)
    if (!dataTable1) {
        const table1Element = document.getElementById('table1');
        if (table1Element) {
            const table1Data = extractDataFromTableElement(table1Element);
            if (table1Data && table1Data.length > 0) {
                dataTable1 = table1Data;
            }
        }
    }
    
    if (!dataTable2) {
        const table2Element = document.getElementById('table2');
        if (table2Element) {
            const table2Data = extractDataFromTableElement(table2Element);
            if (table2Data && table2Data.length > 0) {
                dataTable2 = table2Data;
            }
        }
    }
    
    if (dataTable1 && dataTable1.length > 0 && dataTable1[0]) {
        dataTable1[0].forEach((header, index) => {
            if (header && header.toString().trim() !== '') {
                allHeaders.add(header.toString());
            }
        });
    }
    
    if (dataTable2 && dataTable2.length > 0 && dataTable2[0]) {
        dataTable2[0].forEach((header, index) => {
            if (header && header.toString().trim() !== '') {
                allHeaders.add(header.toString());
            }
        });
    }
    
    // Create checkboxes for each header
    if (allHeaders.size > 0) {
        // Add control buttons
        const controlsWrapper = document.createElement('div');
        controlsWrapper.className = 'key-columns-controls';
        controlsWrapper.style.cssText = 'padding: 8px; border-bottom: 1px solid #eee; display: flex; gap: 8px; justify-content: space-between;';
        
        // Get current language for button text
        const currentLang = window.location.pathname.includes('/ru/') ? 'ru' : 
                           window.location.pathname.includes('/pl/') ? 'pl' :
                           window.location.pathname.includes('/es/') ? 'es' :
                           window.location.pathname.includes('/de/') ? 'de' :
                           window.location.pathname.includes('/ja/') ? 'ja' :
                           window.location.pathname.includes('/pt/') ? 'pt' :
                           window.location.pathname.includes('/zh/') ? 'zh' :
                           window.location.pathname.includes('/ar/') ? 'ar' : 'en';
        
        const selectAllTexts = {
            'ru': 'Выбрать все',
            'pl': 'Zaznacz wszystkie',
            'es': 'Seleccionar todo',
            'de': 'Alle auswählen',
            'ja': 'すべて選択',
            'pt': 'Selecionar tudo',
            'zh': '全选',
            'ar': 'تحديد الكل',
            'en': 'Select All'
        };
        
        const clearAllTexts = {
            'ru': 'Очистить все',
            'pl': 'Wyczyść wszystkie',
            'es': 'Limpiar todo',
            'de': 'Alle löschen',
            'ja': 'すべてクリア',
            'pt': 'Limpar tudo',
            'zh': '清除全部',
            'ar': 'مسح الكل',
            'en': 'Clear All'
        };
        
        const selectAllBtn = document.createElement('button');
        selectAllBtn.textContent = selectAllTexts[currentLang];
        selectAllBtn.style.cssText = 'padding: 4px 8px; border: 1px solid #007bff; border-radius: 3px; background: #007bff; color: white; cursor: pointer; font-size: 12px;';
        selectAllBtn.onclick = () => {
            const checkboxes = dropdownContent.querySelectorAll('input[name="keyColumns"]');
            checkboxes.forEach(cb => {
                cb.checked = true;
                const wrapper = cb.closest('.key-column-checkbox');
                if (wrapper) updateCheckboxStyle(wrapper, true);
            });
            updateDropdownButtonText();
        };
        
        const clearAllBtn = document.createElement('button');
        clearAllBtn.textContent = clearAllTexts[currentLang];
        clearAllBtn.style.cssText = 'padding: 4px 8px; border: 1px solid #6c757d; border-radius: 3px; background: #6c757d; color: white; cursor: pointer; font-size: 12px;';
        clearAllBtn.onclick = () => {
            const checkboxes = dropdownContent.querySelectorAll('input[name="keyColumns"]');
            checkboxes.forEach(cb => {
                cb.checked = false;
                const wrapper = cb.closest('.key-column-checkbox');
                if (wrapper) updateCheckboxStyle(wrapper, false);
            });
            updateDropdownButtonText();
        };
        
        controlsWrapper.appendChild(selectAllBtn);
        controlsWrapper.appendChild(clearAllBtn);
        dropdownContent.appendChild(controlsWrapper);
        
        Array.from(allHeaders).sort().forEach(header => {
            const checkboxWrapper = document.createElement('div');
            checkboxWrapper.className = 'key-column-checkbox';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `key-col-${header.replace(/[^a-zA-Z0-9]/g, '_')}`;
            checkbox.value = header;
            checkbox.name = 'keyColumns';
            
            // Restore previous selection state if it was selected
            if (currentlySelected.includes(header)) {
                checkbox.checked = true;
            }
            
            // Get column data for type detection
            const headerIndex = dataTable1?.[0]?.indexOf(header) ?? 
                               dataTable2?.[0]?.indexOf(header) ?? -1;
            
            let columnValues = [];
            if (headerIndex !== -1) {
                // Get values from both tables
                if (dataTable1 && dataTable1.length > 1) {
                    for (let i = 1; i < dataTable1.length; i++) {
                        if (dataTable1[i] && dataTable1[i][headerIndex] !== undefined) {
                            columnValues.push(dataTable1[i][headerIndex]);
                        }
                    }
                }
                if (dataTable2 && dataTable2.length > 1) {
                    for (let i = 1; i < dataTable2.length; i++) {
                        if (dataTable2[i] && dataTable2[i][headerIndex] !== undefined) {
                            columnValues.push(dataTable2[i][headerIndex]);
                        }
                    }
                }
            }
            
            // Determine column type
            const columnType = getColumnType(columnValues, header);
            
            // Create type icon
            const typeIcon = document.createElement('span');
            typeIcon.className = `column-type-icon column-type-${columnType}`;
            typeIcon.title = `Column type: ${columnType}`;
            
            const label = document.createElement('label');
            label.htmlFor = checkbox.id;
            label.textContent = header;
            
            // Add click handler for the wrapper
            checkboxWrapper.addEventListener('click', function(e) {
                if (e.target !== checkbox) {
                    checkbox.checked = !checkbox.checked;
                }
                updateCheckboxStyle(checkboxWrapper, checkbox.checked);
                updateDropdownButtonText();
                
                // Check if export button should be re-enabled
                setTimeout(() => {
                    if (typeof checkAndEnableExportButton === 'function') {
                        checkAndEnableExportButton();
                    }
                }, 50);
            });
            
            // Add change handler for the checkbox
            checkbox.addEventListener('change', function() {
                updateCheckboxStyle(checkboxWrapper, this.checked);
                updateDropdownButtonText();
                
                // Check if export button should be re-enabled
                setTimeout(() => {
                    if (typeof checkAndEnableExportButton === 'function') {
                        checkAndEnableExportButton();
                    }
                }, 50);
            });
            
            checkboxWrapper.appendChild(checkbox);
            checkboxWrapper.appendChild(typeIcon);
            checkboxWrapper.appendChild(label);
            dropdownContent.appendChild(checkboxWrapper);
            
            // Initialize checkbox style based on its state
            updateCheckboxStyle(checkboxWrapper, checkbox.checked);
        });
        
        // Enable dropdown functionality
        setupDropdownToggle();
        updateDropdownButtonText();
        
        // Auto-detect key columns if no selections and force update (new file loaded)
        if (forceUpdate && currentlySelected.length === 0) {
            setTimeout(() => {
                tryAutoDetectKeyColumns(allHeaders, dataTable1, dataTable2);
            }, 100);
        }
    } else {
        // Show placeholder with manual refresh button
        dropdownContent.innerHTML = `
            <div class="key-columns-placeholder">
                Load files to see available columns
                <br><br>
                <button onclick="window.debouncedUpdateKeyColumnsOptions(true)" style="padding: 4px 8px; margin-top: 5px; border: 1px solid #ccc; border-radius: 3px; background: #f8f9fa; cursor: pointer;">
                    <i data-lucide="refresh-cw"></i> Refresh
                </button>
            </div>
        `;
        updateDropdownButtonText();
    }
}

// Setup dropdown toggle functionality
function setupDropdownToggle() {
    
    const dropdown = document.querySelector('.key-columns-dropdown');
    const dropdownButton = document.getElementById('keyColumnsDropdownButton');
    
    if (!dropdown || !dropdownButton) {
        console.error('❌ Required elements for dropdown toggle not found');
        return;
    }
    
    // Remove existing listeners to prevent duplicates
    dropdownButton.removeEventListener('click', toggleDropdown);
    document.removeEventListener('click', closeDropdownOnClickOutside);
    
    // Add click handler for dropdown button
    dropdownButton.addEventListener('click', toggleDropdown);
    
    // Close dropdown when clicking outside
    document.addEventListener('click', closeDropdownOnClickOutside);
}

// Toggle dropdown open/close
function toggleDropdown(e) {
    e.stopPropagation();
    const dropdown = document.querySelector('.key-columns-dropdown');
    const dropdownButton = document.getElementById('keyColumnsDropdownButton');
    
    if (!dropdown || !dropdownButton) {
        console.error('❌ Dropdown elements not found in toggle function');
        return;
    }
    
    if (dropdown.classList.contains('open')) {
        dropdown.classList.remove('open');
        dropdownButton.classList.remove('active');
    } else {        
        dropdown.classList.add('open');
        dropdownButton.classList.add('active');
    }
}

// Close dropdown when clicking outside
function closeDropdownOnClickOutside(e) {
    const dropdown = document.querySelector('.key-columns-dropdown');
    const dropdownButton = document.getElementById('keyColumnsDropdownButton');
    
    if (!dropdown.contains(e.target)) {
        dropdown.classList.remove('open');
        dropdownButton.classList.remove('active');
    }
}

// Update dropdown button text based on selected items
function updateDropdownButtonText() {
    const dropdownText = document.querySelector('.dropdown-text');
    const selectedColumns = getSelectedKeyColumns();
    
    if (!dropdownText) return;
    
    // Get current language for text
    const currentLang = window.location.pathname.includes('/ru/') ? 'ru' : 
                       window.location.pathname.includes('/pl/') ? 'pl' :
                       window.location.pathname.includes('/es/') ? 'es' :
                       window.location.pathname.includes('/de/') ? 'de' :
                       window.location.pathname.includes('/ja/') ? 'ja' :
                       window.location.pathname.includes('/pt/') ? 'pt' :
                       window.location.pathname.includes('/zh/') ? 'zh' :
                       window.location.pathname.includes('/ar/') ? 'ar' : 'en';
    
    const selectTexts = {
        'ru': 'Выберите ключевые колонки...',
        'pl': 'Wybierz kolumny kluczowe...',
        'es': 'Seleccionar columnas clave...',
        'de': 'Schlüsselspalten auswählen...',
        'ja': 'キー列を選択...',
        'pt': 'Selecionar colunas-chave...',
        'zh': '选择关键列...',
        'ar': 'اختر الأعمدة الرئيسية...',
        'en': 'Select key columns...'
    };
    
    const columnsSelectedTexts = {
        'ru': 'колонок выбрано',
        'pl': 'kolumn wybrane',
        'es': 'columnas seleccionadas',
        'de': 'Spalten ausgewählt',
        'ja': '列が選択されました',
        'pt': 'colunas selecionadas',
        'zh': '列已选择',
        'ar': 'أعمدة محددة',
        'en': 'columns selected'
    };
    
    if (selectedColumns.length === 0) {
        dropdownText.textContent = selectTexts[currentLang];
        dropdownText.style.color = '#6c757d';
    } else if (selectedColumns.length === 1) {
        dropdownText.textContent = selectedColumns[0];
        dropdownText.style.color = '#495057';
    } else {
        dropdownText.textContent = `${selectedColumns.length} ${columnsSelectedTexts[currentLang]}`;
        dropdownText.style.color = '#495057';
    }
}

// Update checkbox visual style based on checked state
function updateCheckboxStyle(wrapper, isChecked) {
    if (isChecked) {
        wrapper.classList.add('selected');
    } else {
        wrapper.classList.remove('selected');
    }
}

// Get selected key columns from the checkboxes
function getSelectedKeyColumns() {
    const checkboxes = document.querySelectorAll('input[name="keyColumns"]:checked');
    const selectedColumns = Array.from(checkboxes).map(checkbox => checkbox.value);

    return selectedColumns;
}

// Set selected key columns in the UI (mark checkboxes as checked)
function setSelectedKeyColumns(columnNames) {
    
    if (!columnNames || !Array.isArray(columnNames) || columnNames.length === 0) {
        console.log('❌ No column names provided to set as selected');
        return;
    }
    
    const dropdownContent = document.getElementById('keyColumnsDropdownContent');
    if (!dropdownContent) {
        console.error('❌ Dropdown content not found for setting selected columns');
        return;
    }
    
    // Check if dropdown has checkboxes
    const existingCheckboxes = dropdownContent.querySelectorAll('input[name="keyColumns"]');
    if (existingCheckboxes.length === 0) {
        console.warn('⚠️ No checkboxes found in dropdown, updating dropdown first...');
        // Try to update dropdown and retry after delay
        if (typeof debouncedUpdateKeyColumnsOptions === 'function') {
            debouncedUpdateKeyColumnsOptions(true);
        }
        setTimeout(() => setSelectedKeyColumns(columnNames), 500); // Increased delay
        return;
    }
    
    // Convert indexes to column names if needed
    let columnsToSelect = columnNames;
    if (typeof columnNames[0] === 'number') {
        // If we got indexes, convert them to column names
        const headers = (window.data1 && window.data1.length > 0) ? window.data1[0] : 
                       (typeof data1 !== 'undefined' && data1.length > 0) ? data1[0] : [];
        columnsToSelect = columnNames.map(index => headers[index]).filter(name => name);
    }
    
    //console.log('🔑 Setting selected key columns:', columnsToSelect);
    
    // Find and check the corresponding checkboxes
    let checkedCount = 0;
    columnsToSelect.forEach(columnName => {
        const checkbox = dropdownContent.querySelector(`input[name="keyColumns"][value="${columnName}"]`);
        if (checkbox) {
            checkbox.checked = true;
            checkedCount++;
            
            // Update visual style
            const wrapper = checkbox.closest('.key-column-checkbox');
            if (wrapper) {
                updateCheckboxStyle(wrapper, true);
            }
            
            //console.log(`✅ Selected column: ${columnName}`);
        } else {
            console.warn('⚠️ Checkbox not found for column:', columnName);
        }
    });
    
    //console.log(`🎯 Successfully selected ${checkedCount} out of ${columnsToSelect.length} key columns`);
    
    // Update dropdown button text
    updateDropdownButtonText();
}

// Get indexes of key columns for comparison
function getKeyColumnIndexes(headers, selectedKeyColumns) {

    if (!selectedKeyColumns || selectedKeyColumns.length === 0) {
        // If no key columns selected, use automatic detection
        if (typeof smartDetectKeyColumns === 'function') {
            // smartDetectKeyColumns expects (headers, data) where data includes headers as first row
            const mockData = [headers]; // Create minimal data structure for key detection
            const autoDetectedIndexes = smartDetectKeyColumns(headers, mockData);
            return autoDetectedIndexes || [0]; // Fallback to first column if nothing detected
        } else {
            return [0]; // Fallback to first column
        }
    }
    
    // Return indexes of selected key columns
    const indexes = [];
    selectedKeyColumns.forEach(keyColumn => {
        const index = headers.indexOf(keyColumn);
        if (index !== -1) {
            indexes.push(index);
        }
    });
    
    //('🔑 Selected key column indexes:', indexes);
    return indexes.length > 0 ? indexes : [0]; // Fallback to first column if no valid indexes found
}

// Initialize key columns functionality when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    
    // Setup dropdown toggle immediately
    setTimeout(setupDropdownToggle, 100);
    
    // Update key columns options when files are changed
    const file1Input = document.getElementById('file1');
    const file2Input = document.getElementById('file2');
    
    if (file1Input) {
        file1Input.addEventListener('change', function() {
            setTimeout(() => {
                debouncedUpdateKeyColumnsOptions(true); // Force update when file changes
            }, 3000); // Increased timeout to ensure file processing and cache loading
        });
    }
    
    if (file2Input) {
        file2Input.addEventListener('change', function() {
            setTimeout(() => {
                debouncedUpdateKeyColumnsOptions(true); // Force update when file changes
            }, 3000); // Increased timeout to ensure file processing and cache loading
        });
    }
    
    // Initial update only - no periodic updates to avoid resetting user selections
    setTimeout(() => debouncedUpdateKeyColumnsOptions(true), 500);
    
    // Cleanup old cache entries occasionally (run once per session)
    setTimeout(() => {
        cleanupKeyColumnsCache();
    }, 1000);
});

// Public function that uses debouncing
function updateKeyColumnsOptions(forceUpdate = false) {
    return debouncedUpdateKeyColumnsOptions(forceUpdate);
}

// Key columns cache functionality
const KEY_COLUMNS_CACHE_KEY = 'keyColumnsCache';
const MAX_CACHE_ENTRIES = 50; // Limit number of cached files

// Generate cache key for file(s)
function generateFileCacheKey() {
    const fileName1 = window.fileName1 || '';
    const fileName2 = window.fileName2 || '';
    const sheetName1 = window.sheetName1 || '';
    const sheetName2 = window.sheetName2 || '';
    
    // Create a unique key based on files and sheets
    let key = '';
    if (fileName1) {
        key += fileName1;
        if (sheetName1) key += ':' + sheetName1;
    }
    if (fileName2) {
        if (key) key += '|';
        key += fileName2;
        if (sheetName2) key += ':' + sheetName2;
    }
    
    return key || 'default';
}

// Save key columns selection to cache
function saveKeyColumnsToCache(selectedColumns) {
    try {
        const cacheKey = generateFileCacheKey();
        if (!cacheKey || cacheKey === 'default') return;
        
        let cache = {};
        const stored = localStorage.getItem(KEY_COLUMNS_CACHE_KEY);
        if (stored) {
            cache = JSON.parse(stored);
        }
        
        // Add timestamp for cleanup
        cache[cacheKey] = {
            columns: selectedColumns,
            timestamp: Date.now()
        };
        
        // Cleanup old entries if cache gets too large
        const entries = Object.entries(cache);
        if (entries.length > MAX_CACHE_ENTRIES) {
            // Sort by timestamp and keep only the most recent entries
            entries.sort((a, b) => b[1].timestamp - a[1].timestamp);
            cache = {};
            entries.slice(0, MAX_CACHE_ENTRIES).forEach(([key, value]) => {
                cache[key] = value;
            });
        }
        
        localStorage.setItem(KEY_COLUMNS_CACHE_KEY, JSON.stringify(cache));
        //console.log('💾 Saved key columns to cache for:', cacheKey, selectedColumns);
    } catch (error) {
        //console.warn('⚠️ Failed to save key columns to cache:', error);
    }
}

// Load key columns selection from cache
function loadKeyColumnsFromCache() {
    try {
        const cacheKey = generateFileCacheKey();
        if (!cacheKey || cacheKey === 'default') return null;
        
        const stored = localStorage.getItem(KEY_COLUMNS_CACHE_KEY);
        if (!stored) return null;
        
        const cache = JSON.parse(stored);
        const cachedData = cache[cacheKey];
        
        if (cachedData && cachedData.columns) {
            //console.log('📥 Loaded key columns from cache for:', cacheKey, cachedData.columns);
            return cachedData.columns;
        }
    } catch (error) {
        //console.warn('⚠️ Failed to load key columns from cache:', error);
    }
    return null;
}

// Clear old cache entries (older than 30 days)
function cleanupKeyColumnsCache() {
    try {
        const stored = localStorage.getItem(KEY_COLUMNS_CACHE_KEY);
        if (!stored) return;
        
        const cache = JSON.parse(stored);
        const thirtyDaysAgo = Date.now() - (30 * 24 * 60 * 60 * 1000);
        
        let cleaned = false;
        Object.keys(cache).forEach(key => {
            if (cache[key].timestamp < thirtyDaysAgo) {
                delete cache[key];
                cleaned = true;
            }
        });
        
        if (cleaned) {
            localStorage.setItem(KEY_COLUMNS_CACHE_KEY, JSON.stringify(cache));
            //console.log('🧹 Cleaned up old key columns cache entries');
        }
    } catch (error) {
        //console.warn('⚠️ Failed to cleanup key columns cache:', error);
    }
}

// Auto-detect and set key columns when no cache is available
function tryAutoDetectKeyColumns(allHeaders, dataTable1, dataTable2) {
    try {
        // Prepare data for auto-detection
        let headers = [];
        let combinedData = [];
        
        if (dataTable1 && dataTable1.length > 0) {
            headers = dataTable1[0];
            combinedData = [...dataTable1];
        }
        
        if (dataTable2 && dataTable2.length > 0) {
            if (headers.length === 0) {
                headers = dataTable2[0];
                combinedData = [...dataTable2];
            } else {
                // Add data from second table
                combinedData.push(...dataTable2.slice(1));
            }
        }
        
        if (headers.length === 0) return;
        
        // Use existing smart detection function if available
        let autoDetectedIndexes = [];
        if (typeof smartDetectKeyColumns === 'function') {
            autoDetectedIndexes = smartDetectKeyColumns(headers, combinedData);
        } else {
            // Fallback simple detection - look for ID-like columns
            for (let i = 0; i < headers.length; i++) {
                const header = headers[i].toString().toLowerCase();
                if (header.includes('id') || header.includes('key') || header.includes('код') || 
                    header.includes('номер') || header.includes('артикул')) {
                    autoDetectedIndexes.push(i);
                    break; // Take first ID-like column
                }
            }
            // If no ID-like columns, use first column
            if (autoDetectedIndexes.length === 0) {
                autoDetectedIndexes = [0];
            }
        }
        
        // Convert indexes to column names
        const autoDetectedColumnNames = autoDetectedIndexes
            .map(index => headers[index])
            .filter(name => name && allHeaders.has(name.toString()));
        
        if (autoDetectedColumnNames.length > 0) {
            //console.log('🤖 Auto-detected key columns:', autoDetectedColumnNames);
            
            // Set these columns as selected in the UI
            setTimeout(() => {
                setSelectedKeyColumns(autoDetectedColumnNames);
            }, 50);
        }
        
    } catch (error) {
        //console.warn('⚠️ Error in auto-detection:', error);
    }
}

// Save current key columns selection to cache (called when comparison starts)
function saveCurrentKeyColumnsSelection() {
    const selectedColumns = getSelectedKeyColumns();
    if (selectedColumns.length > 0) {
        saveKeyColumnsToCache(selectedColumns);
        //console.log('💾 Saved current key columns selection to cache:', selectedColumns);
    }
}

// Clear key columns cache completely
function clearKeyColumnsCache() {
    try {
        localStorage.removeItem(KEY_COLUMNS_CACHE_KEY);
        //console.log('🗑️ Cleared all key columns cache');
        return true;
    } catch (error) {
        //console.warn('⚠️ Failed to clear key columns cache:', error);
        return false;
    }
}

// Make functions globally available
window.updateKeyColumnsOptions = updateKeyColumnsOptions;
window.debouncedUpdateKeyColumnsOptions = debouncedUpdateKeyColumnsOptions;
window.getSelectedKeyColumns = getSelectedKeyColumns;
window.setSelectedKeyColumns = setSelectedKeyColumns;
window.getKeyColumnIndexes = getKeyColumnIndexes;
window.saveKeyColumnsToCache = saveKeyColumnsToCache;
window.loadKeyColumnsFromCache = loadKeyColumnsFromCache;
window.cleanupKeyColumnsCache = cleanupKeyColumnsCache;
window.clearKeyColumnsCache = clearKeyColumnsCache;
window.saveCurrentKeyColumnsSelection = saveCurrentKeyColumnsSelection;
window.tryAutoDetectKeyColumns = tryAutoDetectKeyColumns;