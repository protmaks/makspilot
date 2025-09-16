// Key columns functionality for comparison

// Get column type based on data analysis
function getColumnType(columnValues, columnHeader = '') {
    if (!columnValues || columnValues.length === 0) return 'text';
    
    const nonEmptyValues = columnValues.filter(val => 
        val !== null && val !== undefined && val.toString().trim() !== ''
    );
    
    if (nonEmptyValues.length === 0) return 'text';
    
    let numberCount = 0;
    let dateCount = 0;
    let booleanCount = 0;
    let textCount = 0;
    
    const headerLower = columnHeader.toString().toLowerCase();
    const dateKeywords = ['date', 'time', 'created', 'modified', 'updated', 'birth', 'дата', 'время', 'создан', 'изменен', 'обновлен'];
    const booleanKeywords = ['active', 'enabled', 'disabled', 'visible', 'hidden', 'deleted', 'verified', 'confirmed', 'активен', 'включен', 'отключен', 'видимый', 'скрытый', 'удален', 'подтвержден'];
    const numberKeywords = ['id', 'count', 'amount', 'price', 'cost', 'sum', 'total', 'number', 'номер', 'количество', 'сумма', 'цена', 'стоимость'];
    
    for (let value of nonEmptyValues) {
        const strValue = value.toString().trim().toLowerCase();
        
        // Check for boolean values
        if (['true', 'false', '1', '0', 'yes', 'no', 'да', 'нет', 'истина', 'ложь'].includes(strValue)) {
            booleanCount++;
            continue;
        }
        
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
    const booleanRatio = booleanCount / total;
    
    // Header-based hints
    const headerSuggestsDate = dateKeywords.some(keyword => headerLower.includes(keyword));
    const headerSuggestsBoolean = booleanKeywords.some(keyword => headerLower.includes(keyword));
    const headerSuggestsNumber = numberKeywords.some(keyword => headerLower.includes(keyword));
    
    // Determine type based on ratios and header hints
    if (booleanRatio > 0.7 || (headerSuggestsBoolean && booleanRatio > 0.4)) {
        return 'boolean';
    }
    
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

// Update key columns dropdown when files are loaded
function updateKeyColumnsOptions(forceUpdate = false) {
    console.log('🔄 updateKeyColumnsOptions called', { forceUpdate });
    
    const dropdownContent = document.getElementById('keyColumnsDropdownContent');
    const dropdownButton = document.getElementById('keyColumnsDropdownButton');
    
    console.log('🔍 DOM elements found:', { 
        dropdownContent: !!dropdownContent, 
        dropdownButton: !!dropdownButton 
    });
    
    if (!dropdownContent || !dropdownButton) {
        console.error('❌ Required DOM elements not found for key columns dropdown');
        return;
    }
    
    // Check if we already have checkboxes and data hasn't changed (unless force update)
    const existingCheckboxes = dropdownContent.querySelectorAll('input[name="keyColumns"]');
    if (!forceUpdate && existingCheckboxes.length > 0) {
        console.log('🔧 Skipping update - checkboxes already exist and no force update requested');
        return;
    }
    
    // Save currently selected checkboxes before clearing
    const currentlySelected = [];
    const existingSelectedCheckboxes = dropdownContent.querySelectorAll('input[name="keyColumns"]:checked');
    existingSelectedCheckboxes.forEach(checkbox => {
        currentlySelected.push(checkbox.value);
    });
    
    console.log('💾 Saving currently selected columns:', currentlySelected);
    
    // Clear existing checkboxes
    dropdownContent.innerHTML = '';
    
    // Get headers from both files
    let allHeaders = new Set();
    
    //console.log('🔍 Checking data availability:', {  data1Available: !!(window.data1 || data1),  data1Length: (window.data1 || data1)?.length, data2Available: !!(window.data2 || data2),  data2Length: (window.data2 || data2)?.length, globalData1: !!window.data1, globalData2: !!window.data2, localData1: typeof data1 !== 'undefined' ? !!data1 : false, localData2: typeof data2 !== 'undefined' ? !!data2 : false });
    
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
                //console.log('📋 Extracted data from table1 DOM element');
            }
        }
    }
    
    if (!dataTable2) {
        const table2Element = document.getElementById('table2');
        if (table2Element) {
            const table2Data = extractDataFromTableElement(table2Element);
            if (table2Data && table2Data.length > 0) {
                dataTable2 = table2Data;
                //console.log('📋 Extracted data from table2 DOM element');
            }
        }
    }
    
    //console.log('🔍 Final data status:', { dataTable1Available: !!dataTable1, dataTable1Length: dataTable1?.length, dataTable2Available: !!dataTable2, dataTable2Length: dataTable2?.length });
    
    if (dataTable1 && dataTable1.length > 0 && dataTable1[0]) {
        //console.log('📋 Table 1 headers:', dataTable1[0]);
        dataTable1[0].forEach((header, index) => {
            if (header && header.toString().trim() !== '') {
                allHeaders.add(header.toString());
            }
        });
    }
    
    if (dataTable2 && dataTable2.length > 0 && dataTable2[0]) {
        //console.log('📋 Table 2 headers:', dataTable2[0]);
        dataTable2[0].forEach((header, index) => {
            if (header && header.toString().trim() !== '') {
                allHeaders.add(header.toString());
            }
        });
    }
    
    //console.log('🔍 All unique headers found:', Array.from(allHeaders));
    
    // Create checkboxes for each header
    if (allHeaders.size > 0) {
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
                console.log('🔧 Restoring selection for column:', header);
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
            });
            
            // Add change handler for the checkbox
            checkbox.addEventListener('change', function() {
                updateCheckboxStyle(checkboxWrapper, this.checked);
                updateDropdownButtonText();
            });
            
            checkboxWrapper.appendChild(checkbox);
            checkboxWrapper.appendChild(typeIcon);
            checkboxWrapper.appendChild(label);
            dropdownContent.appendChild(checkboxWrapper);
            
            // Initialize checkbox style based on its state
            updateCheckboxStyle(checkboxWrapper, checkbox.checked);
        });
        
        console.log('✅ Created checkboxes for', allHeaders.size, 'columns');
        
        // TEST: Automatically mark the first column (NAME) for testing
        console.log('🧪 TEST: Attempting to auto-mark NAME column');
        setTimeout(() => {
            const nameCheckbox = dropdownContent.querySelector('input[name="keyColumns"][value="NAME"]');
            if (nameCheckbox) {
                nameCheckbox.checked = true;
                const wrapper = nameCheckbox.closest('.key-column-checkbox');
                if (wrapper) {
                    updateCheckboxStyle(wrapper, true);
                }
                updateDropdownButtonText();
                console.log('✅ TEST: Successfully marked NAME column as selected');
            } else {
                console.log('❌ TEST: NAME checkbox not found');
            }
        }, 200);
        
        // Enable dropdown functionality
        setupDropdownToggle();
        updateDropdownButtonText();
    } else {
        //console.log('⚠️ No headers found, showing placeholder');
        // Show placeholder with manual refresh button
        dropdownContent.innerHTML = `
            <div class="key-columns-placeholder">
                Load files to see available columns
                <br><br>
                <button onclick="window.updateKeyColumnsOptions(true)" style="padding: 4px 8px; margin-top: 5px; border: 1px solid #ccc; border-radius: 3px; background: #f8f9fa; cursor: pointer;">
                    🔄 Refresh
                </button>
            </div>
        `;
        updateDropdownButtonText();
    }
}

// Setup dropdown toggle functionality
function setupDropdownToggle() {
    //console.log('🔧 Setting up dropdown toggle functionality');
    
    const dropdown = document.querySelector('.key-columns-dropdown');
    const dropdownButton = document.getElementById('keyColumnsDropdownButton');
    
    //console.log('🔍 Toggle setup elements:', {  dropdown: !!dropdown,  dropdownButton: !!dropdownButton  });
    
    if (!dropdown || !dropdownButton) {
        console.error('❌ Required elements for dropdown toggle not found');
        return;
    }
    
    // Remove existing listeners to prevent duplicates
    dropdownButton.removeEventListener('click', toggleDropdown);
    document.removeEventListener('click', closeDropdownOnClickOutside);
    
    // Add click handler for dropdown button
    dropdownButton.addEventListener('click', toggleDropdown);
    //console.log('✅ Click handler added to dropdown button');
    
    // Close dropdown when clicking outside
    document.addEventListener('click', closeDropdownOnClickOutside);
    //console.log('✅ Outside click handler added');
}

// Toggle dropdown open/close
function toggleDropdown(e) {
    //console.log('🔽 toggleDropdown called');
    e.stopPropagation();
    const dropdown = document.querySelector('.key-columns-dropdown');
    const dropdownButton = document.getElementById('keyColumnsDropdownButton');
    
    if (!dropdown || !dropdownButton) {
        console.error('❌ Dropdown elements not found in toggle function');
        return;
    }
    
    if (dropdown.classList.contains('open')) {
        //console.log('📤 Closing dropdown');
        dropdown.classList.remove('open');
        dropdownButton.classList.remove('active');
    } else {
        //console.log('📥 Opening dropdown');
        
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
    
    if (selectedColumns.length === 0) {
        dropdownText.textContent = 'Select key columns...';
        dropdownText.style.color = '#6c757d';
    } else if (selectedColumns.length === 1) {
        dropdownText.textContent = selectedColumns[0];
        dropdownText.style.color = '#495057';
    } else {
        dropdownText.textContent = `${selectedColumns.length} columns selected`;
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
    
    //console.log('🔑 getSelectedKeyColumns called - found checkboxes:', { totalCheckboxes: document.querySelectorAll('input[name="keyColumns"]').length, checkedCheckboxes: checkboxes.length, selectedColumns: selectedColumns });
    
    return selectedColumns;
}

// Set selected key columns in the UI (mark checkboxes as checked)
function setSelectedKeyColumns(columnNames) {
    console.log('🔧 setSelectedKeyColumns called with:', columnNames);
    console.log('🔍 Current DOM state check:', {
        dropdownContentExists: !!document.getElementById('keyColumnsDropdownContent'),
        totalCheckboxes: document.querySelectorAll('input[name="keyColumns"]').length,
        data1Available: !!(window.data1 || (typeof data1 !== 'undefined' && data1)),
        data2Available: !!(window.data2 || (typeof data2 !== 'undefined' && data2))
    });
    
    if (!columnNames || !Array.isArray(columnNames) || columnNames.length === 0) {
        console.log('❌ No column names provided to set as selected');
        return;
    }
    
    const dropdownContent = document.getElementById('keyColumnsDropdownContent');
    if (!dropdownContent) {
        console.error('❌ Dropdown content not found for setting selected columns');
        return;
    }
    
    // Convert indexes to column names if needed
    let columnsToSelect = columnNames;
    if (typeof columnNames[0] === 'number') {
        // If we got indexes, convert them to column names
        const headers = (window.data1 && window.data1.length > 0) ? window.data1[0] : 
                       (typeof data1 !== 'undefined' && data1.length > 0) ? data1[0] : [];
        columnsToSelect = columnNames.map(index => headers[index]).filter(name => name);
        console.log('🔄 Converted indexes to column names:', { indexes: columnNames, columnNames: columnsToSelect });
    }
    
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
            
            console.log('✅ Marked column as selected:', columnName);
        } else {
            console.warn('⚠️ Checkbox not found for column:', columnName);
        }
    });
    
    console.log(`🔧 Successfully marked ${checkedCount} of ${columnsToSelect.length} columns as selected`);
    
    // Update dropdown button text
    updateDropdownButtonText();
}

// Get indexes of key columns for comparison
function getKeyColumnIndexes(headers, selectedKeyColumns) {
    //console.log('🔑 getKeyColumnIndexes called with:', {  headersLength: headers ? headers.length : 0, selectedKeyColumns,  selectedLength: selectedKeyColumns ? selectedKeyColumns.length : 0  });
    
    if (!selectedKeyColumns || selectedKeyColumns.length === 0) {
        // If no key columns selected, use automatic detection
        //console.log('🔑 No key columns selected, using automatic detection');
        if (typeof smartDetectKeyColumns === 'function') {
            // smartDetectKeyColumns expects (headers, data) where data includes headers as first row
            const mockData = [headers]; // Create minimal data structure for key detection
            const autoDetectedIndexes = smartDetectKeyColumns(headers, mockData);
            //console.log('🔑 Auto-detected key column indexes:', autoDetectedIndexes);
            return autoDetectedIndexes || [0]; // Fallback to first column if nothing detected
        } else {
            //console.log('🔑 smartDetectKeyColumns not available, using first column');
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
    //console.log('🚀 Key columns functionality initialized');
    
    // Setup dropdown toggle immediately
    setTimeout(setupDropdownToggle, 100);
    
    // Update key columns options when files are changed
    const file1Input = document.getElementById('file1');
    const file2Input = document.getElementById('file2');
    
    if (file1Input) {
        file1Input.addEventListener('change', function() {
            //console.log('📁 File 1 changed, will update key columns in 2 seconds');
            setTimeout(() => {
                //console.log('⏰ Updating key columns after file 1 load');
                updateKeyColumnsOptions(true); // Force update when file changes
            }, 2000); // Increase timeout to ensure file processing is complete
        });
    }
    
    if (file2Input) {
        file2Input.addEventListener('change', function() {
            //console.log('📁 File 2 changed, will update key columns in 2 seconds');
            setTimeout(() => {
                //console.log('⏰ Updating key columns after file 2 load');
                updateKeyColumnsOptions(true); // Force update when file changes
            }, 2000); // Increase timeout to ensure file processing is complete
        });
    }
    
    // Initial update only - no periodic updates to avoid resetting user selections
    setTimeout(() => updateKeyColumnsOptions(true), 500);
});

// Make functions globally available
window.updateKeyColumnsOptions = updateKeyColumnsOptions;
window.getSelectedKeyColumns = getSelectedKeyColumns;
window.setSelectedKeyColumns = setSelectedKeyColumns;
window.getKeyColumnIndexes = getKeyColumnIndexes;