if (typeof window.MaxPilotDuckDB === 'undefined') {
class FastTableComparator {
    constructor() {
        this.initialized = false;
        this.tables = new Map();
        this.mode = 'local';
    }

    async initialize() {
        try {
            console.log('🔧 Initializing fast comparison engine...');
            
            if (typeof WebAssembly !== 'undefined') {
                try {
                    console.log('🚀 Attempting real DuckDB WASM initialization...');
                    await this.initializeWASM();
                    this.mode = 'wasm';
                    console.log('✅ Real DuckDB WASM mode activated!');
                    this.initialized = true;
                    return true;
                } catch (wasmError) {
                    console.log('📝 DuckDB WASM not available, falling back to optimized local mode');
                }
            }
            
            // Fallback to local mode
            await this.initializeLocal();
            this.mode = 'local';
            this.initialized = true;
            
            console.log('✅ Fast comparison engine ready - optimized local mode');
            return true;
            
        } catch (error) {
            console.error('❌ Failed to initialize comparison engine:', error);
            this.initialized = false;
            return false;
        }
    }

    async initializeWASM() {
        try {
            console.log('🔧 Attempting to initialize real DuckDB WASM...');
            
            // Загружаем настоящий DuckDB
            const script = document.createElement('script');
            script.src = '/javascript/duckdb/duckdb-real.js';
            script.type = 'text/javascript';
            
            await new Promise((resolve, reject) => {
                script.onload = resolve;
                script.onerror = reject;
                document.head.appendChild(script);
            });
            
            // Ждем загрузки и инициализируем
            await new Promise(resolve => setTimeout(resolve, 100));
            
            if (!window.duckdbLoader) {
                throw new Error('DuckDB loader not available');
            }
            
            console.log('🔧 Initializing DuckDB WASM engine...');
            this.db = await window.duckdbLoader.initialize();
            
            console.log('✅ Real DuckDB WASM initialized successfully');
            return true;
            
        } catch (error) {
            console.error('❌ Real DuckDB WASM initialization failed:', error);
            throw new Error('Real DuckDB WASM not available');
        }
    }

    async initializeLocal() {
        this.tables.clear();
        return true;
    }

    async createTableFromData(tableName, data, headers = null) {
        if (!this.initialized) {
            throw new Error('Comparator not initialized');
        }

        try {
            const processedData = this.processTableData(data, headers);
            this.tables.set(tableName, processedData);
            
            return true;
            
        } catch (error) {
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

    async compareTablesFast(data1, data2, excludeColumns = [], useTolerance = false) {
        if (this.mode === 'wasm') {
            console.log('� Using DuckDB WASM with original matching logic');
            return await this.compareTablesWithOriginalLogic(data1, data2, excludeColumns, useTolerance);
        } else {
            return await this.compareTablesLocal(data1, data2, excludeColumns, useTolerance);
        }
    }

    async compareTablesWithDuckDB(table1Name, table2Name, excludeColumns = []) {
        console.log('🚀 Using DuckDB WASM for table comparison...');
        const startTime = performance.now();

        try {
            const table1 = this.tables.get(table1Name);
            const table2 = this.tables.get(table2Name);
            
            if (!table1 || !table2) {
                throw new Error('One or both tables not found');
            }

            // Создаем временные таблицы в DuckDB
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

            // Создаем SELECT для общих колонок
            const columnList = commonColumns.map(col => `"${col}"`).join(', ');
            const whereClause = commonColumns.map(col => `t1."${col}" = t2."${col}"`).join(' AND ');

            // SQL для поиска одинаковых строк
            const identicalQuery = `
                SELECT t1.rowid as row1, t2.rowid as row2, 'IDENTICAL' as status
                FROM temp_table1 t1
                INNER JOIN temp_table2 t2 ON ${whereClause}
            `;

            // SQL для строк только в таблице 1
            const onlyInTable1Query = `
                SELECT t1.rowid as row1, NULL as row2, 'ONLY_IN_TABLE1' as status
                FROM temp_table1 t1
                LEFT JOIN temp_table2 t2 ON ${whereClause}
                WHERE t2.rowid IS NULL
            `;

            // SQL для строк только в таблице 2
            const onlyInTable2Query = `
                SELECT NULL as row1, t2.rowid as row2, 'ONLY_IN_TABLE2' as status
                FROM temp_table2 t2
                LEFT JOIN temp_table1 t1 ON ${whereClause}
                WHERE t1.rowid IS NULL
            `;

            // Выполняем запросы параллельно
            const [identicalResult, onlyTable1Result, onlyTable2Result] = await Promise.all([
                window.duckdbLoader.query(identicalQuery),
                window.duckdbLoader.query(onlyInTable1Query),
                window.duckdbLoader.query(onlyInTable2Query)
            ]);

            // Преобразуем результаты
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

            console.log(`🚀 DuckDB WASM comparison completed in ${duration.toFixed(2)}ms`);

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
            console.error('❌ DuckDB comparison failed, falling back to local mode:', error);
            return await this.compareTablesLocal(table1Name, table2Name, excludeColumns);
        }
    }

    async compareTablesWithOriginalLogic(data1, data2, excludeColumns = [], useTolerance = false) {
        console.log('🚀 Using DuckDB WASM with multi-stage comparison logic...');
        console.log('🔧 compareTablesWithOriginalLogic - excludeColumns debug:', {
            excludeColumns: excludeColumns,
            excludeColumnsType: typeof excludeColumns,
            excludeColumnsIsArray: Array.isArray(excludeColumns),
            excludeColumnsLength: excludeColumns?.length || 0,
            excludeColumnsStringified: JSON.stringify(excludeColumns)
        });
        const startTime = performance.now();

        try {
            // Детальная отладка входящих данных
            console.log('🔍 Detailed input analysis:', {
                data1Type: typeof data1,
                data1IsArray: Array.isArray(data1),
                data1Length: data1?.length,
                data1Sample: data1?.slice ? data1.slice(0, 3) : 'No slice method',
                data2Type: typeof data2,
                data2IsArray: Array.isArray(data2),
                data2Length: data2?.length,
                data2Sample: data2?.slice ? data2.slice(0, 3) : 'No slice method',
                excludeColumns,
                useTolerance
            });

            if (!Array.isArray(data1) || !Array.isArray(data2) || data1.length === 0 || data2.length === 0) {
                console.error('❌ Data validation failed:', {
                    data1IsArray: Array.isArray(data1),
                    data2IsArray: Array.isArray(data2),
                    data1Length: data1?.length,
                    data2Length: data2?.length
                });
                throw new Error('Invalid input data');
            }

            const headers1 = data1[0] || [];
            const headers2 = data2[0] || [];
            
            console.log('📝 Step 1: Creating tables with SQL...');
            // Функция для очистки имен колонок для SQL
            const sanitizeColumnName = (name, index) => {
                if (!name || typeof name !== 'string') {
                    return `col_${index}`;
                }
                // Удаляем специальные символы и заменяем пробелы на подчеркивания
                let sanitized = name.replace(/[^a-zA-Z0-9а-яА-Я_]/g, '_')
                                   .replace(/\s+/g, '_')
                                   .replace(/_{2,}/g, '_')
                                   .replace(/^_|_$/g, '');
                
                // Если имя пустое или начинается с цифры, добавляем префикс
                if (!sanitized || /^\d/.test(sanitized)) {
                    sanitized = `col_${index}_${sanitized}`;
                }
                
                return sanitized || `col_${index}`;
            };

            // Создаем массивы очищенных имен колонок для использования в SQL
            const sanitizedHeaders1 = headers1.map((h, i) => sanitizeColumnName(h, i));
            const sanitizedHeaders2 = headers2.map((h, i) => sanitizeColumnName(h, i));

            // Создаем SQL для создания таблиц с реальными именами колонок
            const createTable1SQL = `CREATE OR REPLACE TABLE table1 (
                rowid INTEGER,
                ${sanitizedHeaders1.map(h => `"${h}" VARCHAR`).join(', ')}
            )`;

            const createTable2SQL = `CREATE OR REPLACE TABLE table2 (
                rowid INTEGER,
                ${sanitizedHeaders2.map(h => `"${h}" VARCHAR`).join(', ')}
            )`;

            await window.duckdbLoader.query(createTable1SQL);
            await window.duckdbLoader.query(createTable2SQL);

            console.log('📊 Step 2: Inserting data in batches...');
            // Вставляем данные пакетами
            const insertBatch = async (tableName, data, headers) => {
                const BATCH_SIZE = 1000;
                for (let i = 1; i < data.length; i += BATCH_SIZE) {
                    const batchEnd = Math.min(i + BATCH_SIZE, data.length);
                    const batchData = data.slice(i, batchEnd);
                    
                    const values = batchData.map((row, idx) => {
                        const rowId = i + idx - 1; // 0-based row index
                        const cleanRow = headers.map((_, colIdx) => {
                            const val = row[colIdx] || '';
                            return `'${val.toString().replace(/'/g, "''")}'`;
                        }).join(', ');
                        return `(${rowId}, ${cleanRow})`;
                    }).join(', ');

                    if (values) {
                        const insertSQL = `INSERT INTO ${tableName} VALUES ${values}`;
                        await window.duckdbLoader.query(insertSQL);
                    }
                }
            };

            await insertBatch('table1', data1, headers1);
            await insertBatch('table2', data2, headers2);

            console.log('🔍 Step 3: Filtering columns and detecting key columns...');
            
            // Определяем колонки для сравнения (исключаем указанные)
            const comparisonColumns = [];
            headers1.forEach((header, index) => {
                // Проверяем, нужно ли исключить эту колонку
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
            
            console.log('🔍 Column filtering results:', {
                totalColumns: headers1.length,
                excludeColumns: excludeColumns,
                comparisonColumns: comparisonColumns,
                comparisonColumnNames: comparisonColumns.map(idx => headers1[idx]),
                excludedColumnNames: excludeColumns
            });
            
            if (comparisonColumns.length === 0) {
                throw new Error('All columns are excluded from comparison');
            }
            
            // Определяем ключевые поля на основе заголовков (используем оригинальный smartDetectKeyColumns)
            const allKeyColumns = this.detectKeyColumnsSQL(headers1);
            // Фильтруем ключевые колонки, исключая те, что исключены из сравнения
            const keyColumns = allKeyColumns.filter(keyCol => comparisonColumns.includes(keyCol));
            
            console.log('🔑 Key columns detected:', keyColumns);
            console.log('🔑 Headers analysis:', {
                headers1: headers1,
                headers2: headers2,
                headers1Length: headers1.length,
                headers2Length: headers2.length,
                allKeyColumns: allKeyColumns,
                filteredKeyColumns: keyColumns
            });

            console.log('🎯 Step 4: Finding identical rows...');
            
            // Проверим сначала общее количество строк в каждой таблице
            const table1CountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM table1');
            const table2CountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM table2');
            const table1Count = Number(table1CountResult.toArray()[0]?.count || 0);
            const table2Count = Number(table2CountResult.toArray()[0]?.count || 0);
            console.log('📊 Table counts:', { table1Count, table2Count });
            
            // Сначала находим полностью идентичные строки (используем только колонки для сравнения)
            const identicalSQL = `
                CREATE OR REPLACE TABLE identical_pairs AS
                SELECT 
                    t1.rowid as row1_id,
                    t2.rowid as row2_id,
                    'IDENTICAL' as match_type
                FROM table1 t1
                INNER JOIN table2 t2 ON (
                    ${comparisonColumns.map(colIdx => 
                        useTolerance 
                            ? `t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}"`
                            : `UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}"))`
                    ).join(' AND ')}
                )
            `;
            
            console.log('🔍 Identical SQL query sample conditions:', {
                useTolerance,
                comparisonColumns: comparisonColumns,
                firstCondition: useTolerance 
                    ? `t1."${sanitizedHeaders1[comparisonColumns[0]]}" = t2."${sanitizedHeaders2[comparisonColumns[0]]}"`
                    : `UPPER(TRIM(t1."${sanitizedHeaders1[comparisonColumns[0]]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[comparisonColumns[0]]}"))`,
                firstColumnName: headers1[comparisonColumns[0]],
                totalConditions: comparisonColumns.length,
                excludedColumns: excludeColumns
            });
            
            console.log('🔍 Identical SQL query:', identicalSQL);
            await window.duckdbLoader.query(identicalSQL);
            
            // Проверяем количество найденных идентичных строк
            const identicalCountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM identical_pairs');
            const identicalCount = Number(identicalCountResult.toArray()[0]?.count || 0);
            console.log('📊 Found identical pairs:', identicalCount);

            console.log('🔍 Step 5: Finding similar rows by key columns...');
            // Затем ищем похожие строки по ключевым полям (исключая уже найденные идентичные)
            const keyColumnChecks = keyColumns.map(colIdx => 
                useTolerance 
                    ? `t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}"`
                    : `UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}"))`
            ).join(' AND ');

            // Более строгие требования для сопоставления (основанные на колонках для сравнения)
            const minKeyMatchesRequired = Math.max(1, Math.ceil(keyColumns.length * (useTolerance ? 0.8 : 0.8))); // Минимум 80% ключевых полей
            const minTotalMatchesRequired = Math.max(2, Math.ceil(comparisonColumns.length * (useTolerance ? 0.6 : 0.7))); // Минимум 60-70% колонок для сравнения

            console.log('🔑 Key column condition:', keyColumnChecks);
            console.log('🔑 Matching requirements updated:', {
                keyColumns: keyColumns.length,
                minKeyMatches: Math.ceil(keyColumns.length * (useTolerance ? 0.8 : 0.8)),
                comparisonColumns: comparisonColumns.length,
                minTotalMatches: Math.ceil(comparisonColumns.length * (useTolerance ? 0.7 : 0.8)),
                maxTotalMatches: comparisonColumns.length - 1,
                strategy: 'strict_similar_matching_with_exclusions'
            });

            // Настраиваемый лимит для SIMILAR пар в зависимости от размера данных
            const similarLimit = Math.max(1000, Math.min(10000, table1Count + table2Count));
            console.log('🔧 SIMILAR pairs limit calculated:', similarLimit);

            const similarSQL = `
                CREATE OR REPLACE TABLE similar_pairs AS
                WITH key_matches AS (
                    SELECT 
                        t1.rowid as row1_id,
                        t2.rowid as row2_id,
                        ${comparisonColumns.map(colIdx => 
                            useTolerance 
                                ? `CASE WHEN t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}" THEN 1 ELSE 0 END`
                                : `CASE WHEN UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}")) THEN 1 ELSE 0 END`
                        ).join(' + ')} as total_matches,
                        ${keyColumns.map(colIdx => 
                            useTolerance 
                                ? `CASE WHEN t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}" THEN 1 ELSE 0 END`
                                : `CASE WHEN UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}")) THEN 1 ELSE 0 END`
                        ).join(' + ')} as key_matches
                    FROM table1 t1
                    CROSS JOIN table2 t2
                    WHERE NOT EXISTS (
                        SELECT 1 FROM identical_pairs ip 
                        WHERE ip.row1_id = t1.rowid AND ip.row2_id = t2.rowid
                    )
                )
                SELECT 
                    row1_id, row2_id, 'SIMILAR' as match_type, total_matches, key_matches
                FROM key_matches  
                WHERE key_matches >= ${Math.ceil(keyColumns.length * (useTolerance ? 0.8 : 0.8))}
                  AND total_matches >= ${Math.ceil(comparisonColumns.length * (useTolerance ? 0.7 : 0.8))}
                  AND total_matches < ${comparisonColumns.length}
                ORDER BY key_matches DESC, total_matches DESC
                LIMIT ${similarLimit}  -- Динамический лимит на основе размера данных
            `;
            
            console.log('🔍 Similar SQL query (first part):', similarSQL.substring(0, 500) + '...');
            console.log('🔍 SIMILAR matching criteria:', {
                keyColumns: keyColumns.length,
                comparisonColumns: comparisonColumns.length,
                minKeyMatches: Math.ceil(keyColumns.length * (useTolerance ? 0.8 : 0.8)),
                minTotalMatches: Math.ceil(comparisonColumns.length * (useTolerance ? 0.7 : 0.8)),
                maxTotalMatches: comparisonColumns.length - 1,
                useTolerance: useTolerance,
                excludedColumns: excludeColumns
            });
            await window.duckdbLoader.query(similarSQL);
            
            // Проверяем количество найденных похожих строк
            const similarCountResult = await window.duckdbLoader.query('SELECT COUNT(*) as count FROM similar_pairs');
            const similarCount = Number(similarCountResult.toArray()[0]?.count || 0);
            console.log('📊 Found similar pairs:', similarCount);
            
            // Отладочная информация: сколько всего было кандидатов на SIMILAR
            const candidatesCountResult = await window.duckdbLoader.query(`
                SELECT COUNT(*) as count FROM (
                    SELECT 
                        t1.rowid as row1_id,
                        t2.rowid as row2_id,
                        ${comparisonColumns.map(colIdx => 
                            useTolerance 
                                ? `CASE WHEN t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}" THEN 1 ELSE 0 END`
                                : `CASE WHEN UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}")) THEN 1 ELSE 0 END`
                        ).join(' + ')} as total_matches,
                        ${keyColumns.map(colIdx => 
                            useTolerance 
                                ? `CASE WHEN t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}" THEN 1 ELSE 0 END`
                                : `CASE WHEN UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}")) THEN 1 ELSE 0 END`
                        ).join(' + ')} as key_matches
                    FROM table1 t1
                    CROSS JOIN table2 t2
                    WHERE NOT EXISTS (
                        SELECT 1 FROM identical_pairs ip 
                        WHERE ip.row1_id = t1.rowid AND ip.row2_id = t2.rowid
                    )
                ) candidates
            `);
            const candidatesCount = Number(candidatesCountResult.toArray()[0]?.count || 0);
            console.log('🔍 Total SIMILAR candidates (before filtering):', candidatesCount);
            
            // Проверим сколько кандидатов прошло каждый фильтр
            const filterStatsResult = await window.duckdbLoader.query(`
                SELECT 
                    COUNT(*) as total_candidates,
                    COUNT(CASE WHEN key_matches >= ${Math.ceil(keyColumns.length * (useTolerance ? 0.8 : 0.8))} THEN 1 END) as passed_key_filter,
                    COUNT(CASE WHEN total_matches >= ${Math.ceil(comparisonColumns.length * (useTolerance ? 0.7 : 0.8))} THEN 1 END) as passed_total_filter,
                    COUNT(CASE WHEN total_matches < ${comparisonColumns.length} THEN 1 END) as passed_not_identical_filter,
                    COUNT(CASE WHEN key_matches >= ${Math.ceil(keyColumns.length * (useTolerance ? 0.8 : 0.8))} 
                               AND total_matches >= ${Math.ceil(comparisonColumns.length * (useTolerance ? 0.7 : 0.8))}
                               AND total_matches < ${comparisonColumns.length} THEN 1 END) as passed_all_filters,
                    AVG(total_matches) as avg_total_matches,
                    AVG(key_matches) as avg_key_matches,
                    MIN(total_matches) as min_total_matches,
                    MAX(total_matches) as max_total_matches
                FROM (
                    SELECT 
                        ${comparisonColumns.map(colIdx => 
                            useTolerance 
                                ? `CASE WHEN t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}" THEN 1 ELSE 0 END`
                                : `CASE WHEN UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}")) THEN 1 ELSE 0 END`
                        ).join(' + ')} as total_matches,
                        ${keyColumns.map(colIdx => 
                            useTolerance 
                                ? `CASE WHEN t1."${sanitizedHeaders1[colIdx]}" = t2."${sanitizedHeaders2[colIdx]}" THEN 1 ELSE 0 END`
                                : `CASE WHEN UPPER(TRIM(t1."${sanitizedHeaders1[colIdx]}")) = UPPER(TRIM(t2."${sanitizedHeaders2[colIdx]}")) THEN 1 ELSE 0 END`
                        ).join(' + ')} as key_matches
                    FROM table1 t1
                    CROSS JOIN table2 t2
                    WHERE NOT EXISTS (
                        SELECT 1 FROM identical_pairs ip 
                        WHERE ip.row1_id = t1.rowid AND ip.row2_id = t2.rowid
                    )
                ) stats
            `);
            const filterStats = filterStatsResult.toArray()[0];
            console.log('🔍 SIMILAR filter statistics:', filterStats);

            console.log('📋 Step 6: Collecting final results...');
            // Получаем все результаты одним запросом
            const finalResultsSQL = `
                -- Идентичные пары
                SELECT 'IDENTICAL' as type, row1_id, row2_id, ${headers1.length} as matches
                FROM identical_pairs
                
                UNION ALL
                
                -- Похожие пары
                SELECT 'SIMILAR' as type, row1_id, row2_id, total_matches as matches
                FROM similar_pairs
                
                UNION ALL
                
                -- Строки только в таблице 1
                SELECT 'ONLY_IN_TABLE1' as type, t1.rowid as row1_id, NULL as row2_id, 0 as matches
                FROM table1 t1
                WHERE NOT EXISTS (SELECT 1 FROM identical_pairs ip WHERE ip.row1_id = t1.rowid)
                  AND NOT EXISTS (SELECT 1 FROM similar_pairs sp WHERE sp.row1_id = t1.rowid)
                
                UNION ALL
                
                -- Строки только в таблице 2
                SELECT 'ONLY_IN_TABLE2' as type, NULL as row1_id, t2.rowid as row2_id, 0 as matches
                FROM table2 t2
                WHERE NOT EXISTS (SELECT 1 FROM identical_pairs ip WHERE ip.row2_id = t2.rowid)
                  AND NOT EXISTS (SELECT 1 FROM similar_pairs sp WHERE sp.row2_id = t2.rowid)
                
                ORDER BY type, row1_id, row2_id
            `;

            console.log('⚡ Step 7: Executing final comparison query...');
            const result = await window.duckdbLoader.query(finalResultsSQL);
            const allResults = result.toArray();

            // Разделяем результаты по типам
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
            console.log(`🚀 Multi-stage DuckDB comparison completed in ${duration.toFixed(2)}ms`);
            console.log(`📊 Results: ${identical.length} identical, ${similar.length} similar, ${onlyInTable1.length} only in table1, ${onlyInTable2.length} only in table2`);

            return {
                identical: identical,
                similar: similar, // Сохраняем separate для правильного отображения
                onlyInTable1: onlyInTable1,
                onlyInTable2: onlyInTable2,
                table1Count: data1.length - 1,
                table2Count: data2.length - 1,
                commonColumns: headers1, // Все колонки (для отображения таблицы)
                comparisonColumns: comparisonColumns, // Колонки, участвующие в сравнении
                excludedColumns: excludeColumns, // Исключенные колонки
                keyColumns: keyColumns,
                performance: {
                    duration: duration,
                    rowsPerSecond: Math.round(((data1.length + data2.length) / duration) * 1000)
                }
            };

        } catch (error) {
            console.error('❌ Multi-stage DuckDB comparison failed:', error);
            throw error;
        }
    }

    // Функция определения ключевых колонок (полная копия smartDetectKeyColumns из functions.js)
    detectKeyColumnsSQL(headers, data = null) {
        console.log('🔍 Using original smartDetectKeyColumns algorithm from functions.js');
        
        if (!headers || headers.length === 0) {
            return [0]; 
        }
        
        const columnCount = headers.length;
        
        // Если нет данных для анализа, используем только заголовки
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
            
            // Анализ заголовка
            let headerScore = 0;
            
            // Поля агрегации имеют самый высокий приоритет
            if (keyIndicators.aggregation.some(keyword => header.includes(keyword))) {
                // Особые случаи для YYYYMM и период
                if (header.match(/^y{4}m{2}$|yyyymm|year.*month|месяц.*год|period|период/i)) {
                    headerScore = 12; 
                } else {
                    headerScore = 9; 
                }
            }
            // Высокоприоритетные ключевые поля (ID, UID, ключи)
            else if (keyIndicators.high.some(keyword => header.includes(keyword))) {
                headerScore = 10;
            }
            // Среднеприоритетные поля (имена, названия)
            else if (keyIndicators.medium.some(keyword => header.includes(keyword))) {
                headerScore = 6;
            }
            // Низкоприоритетные поля (даты)
            else if (keyIndicators.low.some(keyword => header.includes(keyword))) {
                headerScore = 3;
            }
            // Обычные поля
            else {
                headerScore = 1;
            }
            
            score += headerScore * 0.4;
            
            // Базовый анализ уникальности (без данных используем позицию)
            let uniquenessScore = 5; // Средняя оценка по умолчанию
            score += uniquenessScore * 0.4;
            
            // Позиционный бонус (первые колонки важнее)
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
        
        // Сортируем по убыванию оценки
        columnScores.sort((a, b) => b.score - a.score);
        
        console.log('🔍 Column scoring results:', columnScores);
        
        let keyColumns = [];
        
        // Приоритет полям агрегации
        const aggregationFields = columnScores.filter(col => col.isAggregationField);
        if (aggregationFields.length > 0) {
            console.log('📅 Found aggregation fields:', aggregationFields);
            keyColumns.push(aggregationFields[0].index);
            
            // Добавляем дополнительные поля агрегации если они есть
            for (let i = 1; i < Math.min(3, aggregationFields.length); i++) {
                const col = aggregationFields[i];
                if (col.score >= 8) {
                    keyColumns.push(col.index);
                }
            }
        }
        
        // Если нет полей агрегации, берем лучшую по оценке
        if (keyColumns.length === 0) {
            if (columnScores.length > 0) {
                keyColumns.push(columnScores[0].index);
            }
        }
        
        // Добавляем дополнительные ключевые поля до 3 максимум
        for (let i = 0; i < Math.min(3, columnScores.length) && keyColumns.length < 3; i++) {
            const col = columnScores[i];
            if (!keyColumns.includes(col.index)) {
                // Добавляем только поля с высокой оценкой
                if (col.score >= 6) {
                    keyColumns.push(col.index);
                }
            }
        }
        
        // Гарантируем, что есть хотя бы одна ключевая колонка
        if (keyColumns.length === 0) {
            keyColumns = [0];
        }
        
        // Сортируем по позиции
        keyColumns.sort((a, b) => a - b);
        
        console.log('🔑 Final key columns selection:', {
            selectedColumns: keyColumns,
            columnNames: keyColumns.map(idx => headers[idx] || `Column ${idx}`),
            strategy: 'original_smartDetectKeyColumns_algorithm'
        });
        
        return keyColumns;
    }

    async createDuckDBTable(tableName, tableData) {
        // Создаем таблицу с данными в DuckDB
        const columns = tableData.columns.map(col => `"${col}" VARCHAR`).join(', ');
        const createTableSQL = `CREATE OR REPLACE TABLE ${tableName} (rowid INTEGER, ${columns})`;
        
        await window.duckdbLoader.query(createTableSQL);

        // Вставляем данные батчами для производительности
        const BATCH_SIZE = 1000;
        for (let i = 0; i < tableData.rows.length; i += BATCH_SIZE) {
            const batch = tableData.rows.slice(i, i + BATCH_SIZE);
            const values = batch.map(row => {
                const rowData = row.data.map(cell => `'${String(cell || '').replace(/'/g, "''")}'`).join(', ');
                return `(${row.index + 1}, ${rowData})`;
            }).join(', ');
            
            const insertSQL = `INSERT INTO ${tableName} VALUES ${values}`;
            await window.duckdbLoader.query(insertSQL);
        }
    }

    async compareTablesLocal(table1Name, table2Name, excludeColumns = []) {
        if (!this.initialized) {
            throw new Error('Comparator not initialized');
        }

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
        return fastComparator;
    }

    try {
        fastComparator = new FastTableComparator();
        const initialized = await fastComparator.initialize();
        
        if (initialized) {
            showFastModeStatus(true, fastComparator.mode);
            window.duckDBManager = fastComparator;
            window.duckDBAvailable = true;
            
            return fastComparator;
        } else {
            showFastModeStatus(false);
            return null;
        }
    } catch (error) {
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

async function compareTablesWithFastComparator(data1, data2, excludeColumns = [], useTolerance = false, tolerance = 1.5) {
    try {
        console.log('🔧 compareTablesWithFastComparator called with parameters:', {
            excludeColumns: excludeColumns,
            excludeColumnsType: typeof excludeColumns,
            excludeColumnsLength: excludeColumns?.length || 0,
            excludeColumnsArray: Array.isArray(excludeColumns),
            useTolerance: useTolerance,
            tolerance: tolerance
        });
        
        if (!fastComparator || !fastComparator.initialized) {
            console.log('❌ Fast comparator not initialized');
            return null;
        }

        console.log(`🔧 Starting comparison using ${fastComparator.mode} mode`);
        
        // Отладка входящих данных
        console.log('🔍 FastComparator input debug:', {
            data1Type: typeof data1,
            data1IsArray: Array.isArray(data1),
            data1Length: data1?.length,
            data1FirstElement: data1?.[0],
            data2Type: typeof data2,
            data2IsArray: Array.isArray(data2),
            data2Length: data2?.length,
            data2FirstElement: data2?.[0],
            excludeColumns,
            useTolerance
        });
        
        const startTime = performance.now();

        if (fastComparator.mode === 'wasm') {
            // Для WASM режима передаем данные напрямую в compareTablesFast
            console.log('📊 Using DuckDB WASM mode - passing data directly');
            
            const result = await fastComparator.compareTablesFast(data1, data2, excludeColumns, useTolerance);
            
            if (!result) {
                console.log('⚠️ DuckDB WASM returned empty result');
                return null;
            }
            
            const duration = performance.now() - startTime;
            console.log(`✅ Fast comparison completed in ${duration.toFixed(2)}ms`);
            
            return result;
            
        } else {
            // Для локального режима используем prepareDataForComparison
            console.log('📊 Using local mode - preparing data alignment');
            
            const { data1: alignedData1, data2: alignedData2, columnInfo } = prepareDataForComparison(data1, data2);
            
            const headers1 = alignedData1[0] || [];
            const headers2 = alignedData2[0] || [];
            const dataRows1 = alignedData1.slice(1);
            const dataRows2 = alignedData2.slice(1);
            
            console.log(`📊 Table 1: ${dataRows1.length} rows, Table 2: ${dataRows2.length} rows`);
            
            await fastComparator.createTableFromData('table1', dataRows1, headers1);
            await fastComparator.createTableFromData('table2', dataRows2, headers2);
            
            const comparisonResult = await fastComparator.compareTablesFast(
                'table1', 'table2', excludeColumns
            );

            const totalTime = performance.now() - startTime;
            console.log(`⚡ Comparison completed in ${totalTime.toFixed(2)}ms using ${fastComparator.mode} mode`);
            
            return comparisonResult;
        }
        
    } catch (error) {
        console.error('❌ Fast comparison failed:', error);
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
    console.log('🎯 ============ COMPARE TABLES ENHANCED STARTED ============');
    console.log('🎯 compareTablesEnhanced called with useTolerance:', useTolerance);
    
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

    console.log('✅ Data check passed:', { data1Length: data1.length, data2Length: data2.length });

    const totalRows = Math.max(data1.length, data2.length);
    console.log('📊 Total rows check:', { totalRows, limit: MAX_ROWS_LIMIT });
    
    if (totalRows > MAX_ROWS_LIMIT) {
        console.log('❌ Row limit exceeded');
        resultDiv.innerHTML = generateLimitErrorMessage(
            'rows', data1.length, MAX_ROWS_LIMIT, '', 
            'columns', Math.max(data1[0]?.length || 0, data2[0]?.length || 0), MAX_COLS_LIMIT
        );
        summaryDiv.innerHTML = '';
        return;
    }

    console.log('✅ All checks passed, proceeding to comparison logic');

    try {
        console.log('🔍 Checking fastComparator status:', {
            fastComparator: !!fastComparator,
            initialized: fastComparator?.initialized,
            mode: fastComparator?.mode
        });
        
        if (fastComparator && fastComparator.initialized) {
            console.log('🚀 Using fast comparison engine with mode:', fastComparator.mode);
            resultDiv.innerHTML = '<div class="comparison-loading-enhanced">⚡ Using fast comparison engine...</div>';
            summaryDiv.innerHTML = '<div style="text-align: center; padding: 10px;">Processing with enhanced performance...</div>';
            
            setTimeout(async () => {
                try {
                    console.log('📊 Starting DuckDB WASM comparison...');
                    resultDiv.innerHTML = '<div style="text-align: center; padding: 20px; font-size: 16px;">⚡ Fast mode - processing comparison...</div>';
                    
                    const excludedColumns = getExcludedColumns ? getExcludedColumns() : [];
                    const tolerance = window.currentTolerance || 1.5;
                    
                    console.log('🔧 Fast comparison parameters:', {
                        excludedColumns: excludedColumns,
                        excludedColumnsType: typeof excludedColumns,
                        excludedColumnsLength: excludedColumns?.length || 0,
                        useTolerance: useTolerance,
                        tolerance: tolerance,
                        getExcludedColumnsExists: typeof getExcludedColumns !== 'undefined'
                    });
                    
                    console.log('🔧 Calling compareTablesWithFastComparator...');
                    const fastResult = await compareTablesWithFastComparator(
                        data1, data2, excludedColumns, useTolerance, tolerance
                    );
                    
                    if (fastResult) {
                        console.log('✅ DuckDB WASM comparison completed successfully');
                        await processFastComparisonResults(fastResult, useTolerance);
                    } else {
                        console.log('⚠️ DuckDB WASM returned empty result, falling back');
                        await performComparison();
                    }
                } catch (error) {
                    console.error('❌ DuckDB WASM comparison failed:', error);
                    await performComparison();
                }
                
                // Double check and clear result div after comparison completes
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

        console.log('⚠️ Fast comparison not available, using standard mode');
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
    const { identical, similar, onlyInTable1, onlyInTable2, table1Count, table2Count, commonColumns, performance } = fastResult;
    
    // Store fast result for export
    window.currentFastResult = fastResult;
    
    // Handle performance data with fallback
    const perfData = performance || { duration: 0, rowsPerSecond: 0 };
    console.log('📊 Performance data:', perfData);
    
    // Детальная отладочная информация о результатах
    console.log('🔍 Detailed comparison results breakdown:', {
        identicalRows: identical?.length || 0,
        identicalSample: identical?.slice(0, 3),
        similarRows: similar?.length || 0,
        similarSample: similar?.slice(0, 3),
        onlyInTable1Rows: onlyInTable1?.length || 0,
        onlyInTable1Sample: onlyInTable1?.slice(0, 3),
        onlyInTable2Rows: onlyInTable2?.length || 0,
        onlyInTable2Sample: onlyInTable2?.slice(0, 3),
        table1Count,
        table2Count
    });
    
    // Clear any loading messages immediately
    const resultDiv = document.getElementById('result');
    if (resultDiv) {
        resultDiv.innerHTML = '';
        resultDiv.style.display = 'none';
    }
    
    // Подсчитываем количество совпадений раздельно
    const identicalCount = identical?.length || 0;
    const similarCount = similar?.length || 0;
    const totalMatches = identicalCount; // Теперь считаем только identical
    const similarity = table1Count > 0 ? ((identicalCount / Math.max(table1Count, table2Count)) * 100).toFixed(1) : 0;
    
    // Определяем класс для раскраски процента сходства (как в functions.js)
    let percentClass = 'percent-high';
    if (parseFloat(similarity) < 30) percentClass = 'percent-low';
    else if (parseFloat(similarity) < 70) percentClass = 'percent-medium';
    else percentClass = 'percent-high';
    
    // Используем getFileDisplayName для отображения имени файла с листом
    const file1Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName1 || 'File 1', window.sheetName1 || '')
        : (window.fileName1 || 'File 1');
    const file2Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName2 || 'File 2', window.sheetName2 || '')
        : (window.fileName2 || 'File 2');

    const tableHeaders = getSummaryTableHeaders();
    
    // Создаём детализированную сводку с разбивкой по типам совпадений
    const summaryHTML = `
        <div style="overflow-x: auto; margin: 20px 0;">
            <div style="text-align: center; margin-bottom: 15px; padding: 10px; background: rgba(40,167,69,0.1); border-radius: 6px;">
                ⚡ Fast Mode: ${perfData.rowsPerSecond.toLocaleString()} rows/sec | ${perfData.duration.toFixed(2)}ms total
            </div>
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>${tableHeaders.file}</th>
                        <th>${tableHeaders.rowCount}</th>
                        <th>Identical Rows</th>
                        <th>${tableHeaders.rowsOnlyInFile}</th>
                        <th>${tableHeaders.similarity}</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>${file1Name}</strong></td>
                        <td>${table1Count.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 16px; color: #28a745;">${identicalCount.toLocaleString()}</td>
                        <td>${onlyInTable1.length.toLocaleString()}</td>
                        <td rowspan="2" style="vertical-align: middle; font-weight: bold; font-size: 18px;" class="percent-cell ${percentClass}">${similarity}%</td>
                    </tr>
                    <tr>
                        <td><strong>${file2Name}</strong></td>
                        <td>${table2Count.toLocaleString()}</td>
                        <td>${onlyInTable2.length.toLocaleString()}</td>
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
    if (exportBtn) {
        exportBtn.style.display = 'inline-block';
        exportBtn.classList.remove('export-btn-hidden');
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
    
    console.log('🔍 Processing detailed table results:', {
        identicalCount: identical?.length || 0,
        similarCount: similar?.length || 0,
        onlyInTable1Count: onlyInTable1?.length || 0,
        onlyInTable2Count: onlyInTable2?.length || 0,
        workingData1Length: workingData1?.length,
        workingData2Length: workingData2?.length,
        workingData1Sample: workingData1?.slice(0, 3),
        workingData2Sample: workingData2?.slice(0, 3)
    });
    
    // Обрабатываем точные совпадения (identical)
    const maxIdenticalToShow = (onlyInTable1?.length || 0) === 0 && (onlyInTable2?.length || 0) === 0 ? 1000 : 100;
    (identical || []).slice(0, maxIdenticalToShow).forEach(identicalPair => {
        const row1Index = identicalPair.row1; // rowid из DuckDB (0-based для данных)
        const row2Index = identicalPair.row2;
        
        // Безопасное получение данных с проверкой границ
        const row1 = (row1Index >= 0 && row1Index + 1 < workingData1.length) ? workingData1[row1Index + 1] : null;
        const row2 = (row2Index >= 0 && row2Index + 1 < workingData2.length) ? workingData2[row2Index + 1] : null;
        
        console.log('🔍 IDENTICAL pair debug:', {
            originalRow1: identicalPair.row1,
            originalRow2: identicalPair.row2,
            calculatedRow1Index: row1Index,
            calculatedRow2Index: row2Index,
            row1Data: row1,
            row2Data: row2,
            row1IsHeaders: row1 === workingData1[0],
            row2IsHeaders: row2 === workingData2[0],
            workingData1Length: workingData1?.length,
            workingData2Length: workingData2?.length
        });
        
        // Пропускаем пары, где данные не найдены или это заголовки
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
    
    // Обрабатываем похожие совпадения (similar)
    const maxSimilarToShow = 100;
    (similar || []).slice(0, maxSimilarToShow).forEach(similarPair => {
        const row1Index = similarPair.row1; // rowid из DuckDB
        const row2Index = similarPair.row2;
        
        // Безопасное получение данных с проверкой границ
        const row1 = (row1Index >= 0 && row1Index + 1 < workingData1.length) ? workingData1[row1Index + 1] : null;
        const row2 = (row2Index >= 0 && row2Index + 1 < workingData2.length) ? workingData2[row2Index + 1] : null;
        
        console.log('🔍 SIMILAR pair debug:', {
            originalRow1: similarPair.row1,
            originalRow2: similarPair.row2,
            calculatedRow1Index: row1Index,
            calculatedRow2Index: row2Index,
            row1Data: row1,
            row2Data: row2,
            row1IsHeaders: row1 === workingData1[0],
            row2IsHeaders: row2 === workingData2[0]
        });
        
        // Пропускаем пары, где данные не найдены или это заголовки
        if (!row1 || !row2 || row1 === workingData1[0] || row2 === workingData2[0]) {
            console.warn('⚠️ Skipping invalid SIMILAR pair:', { row1Index, row2Index, row1, row2 });
            return;
        }
        
        pairs.push({
            row1: row1,
            row2: row2,
            index1: row1Index,
            index2: row2Index,
            isDifferent: true, // Similar pairs показываем как different для выделения
            onlyIn: null,
            matchType: 'SIMILAR',
            matches: similarPair.matches || 0
        });
    });
    
    // Обрабатываем строки только в таблице 1
    (onlyInTable1 || []).forEach(diff => {
        const rowIndex = diff.row1; // rowid из DuckDB
        
        // Безопасное получение данных с проверкой границ
        const row1 = (rowIndex >= 0 && rowIndex + 1 < workingData1.length) ? workingData1[rowIndex + 1] : null;
        
        console.log('🔍 ONLY_IN_TABLE1 debug:', {
            originalRow1: diff.row1,
            calculatedRowIndex: rowIndex,
            row1Data: row1,
            row1IsHeaders: row1 === workingData1[0]
        });
        
        // Пропускаем строки, где данные не найдены или это заголовки
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
    
    // Обрабатываем строки только в таблице 2
    (onlyInTable2 || []).forEach(diff => {
        const rowIndex = diff.row2; // rowid из DuckDB
        
        // Безопасное получение данных с проверкой границ
        const row2 = (rowIndex >= 0 && rowIndex + 1 < workingData2.length) ? workingData2[rowIndex + 1] : null;
        
        console.log('🔍 ONLY_IN_TABLE2 debug:', {
            originalRow2: diff.row2,
            calculatedRowIndex: rowIndex,
            row2Data: row2,
            row2IsHeaders: row2 === workingData2[0]
        });
        
        // Пропускаем строки, где данные не найдены или это заголовки
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
    
    console.log('📊 Generated pairs for table:', {
        totalPairs: pairs.length,
        identicalPairs: pairs.filter(p => p.matchType === 'IDENTICAL').length,
        similarPairs: pairs.filter(p => p.matchType === 'SIMILAR').length,
        onlyTable1: pairs.filter(p => p.matchType === 'ONLY_IN_TABLE1').length,
        onlyTable2: pairs.filter(p => p.matchType === 'ONLY_IN_TABLE2').length
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
    
    let headerHtml = '<tr><th title="Source - shows which file the data comes from" class="source-column">Source</th>';
    realHeaders.forEach((header, index) => {
        const headerText = header || `Column ${index + 1}`;
        headerHtml += `<th class="sortable" onclick="sortTable(${index})" title="${headerText}">${headerText}</th>`;
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
    
    // Получаем состояние фильтров чекбоксов
    const hideSameEl = document.getElementById('hideSameRows');
    const hideDiffEl = document.getElementById('hideDiffColumns');
    const hideNewRows1El = document.getElementById('hideNewRows1');
    const hideNewRows2El = document.getElementById('hideNewRows2');
    
    const hideSame = hideSameEl ? hideSameEl.checked : false;
    const hideDiffRows = hideDiffEl ? hideDiffEl.checked : false;
    const hideNewRows1 = hideNewRows1El ? hideNewRows1El.checked : false;
    const hideNewRows2 = hideNewRows2El ? hideNewRows2El.checked : false;
    
    // Используем getFileDisplayName для отображения имени файла с листом в колонке Source
    const file1Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName1 || 'File 1', window.sheetName1 || '')
        : (window.fileName1 || 'File 1');
    const file2Name = window.getFileDisplayName 
        ? window.getFileDisplayName(window.fileName2 || 'File 2', window.sheetName2 || '')
        : (window.fileName2 || 'File 2');
    
    pairs.slice(0, 1000).forEach((pair, index) => {
        // Проверяем фильтры чекбоксов
        const row1 = pair.row1;
        const row2 = pair.row2;
        
        // Определяем характеристики строки для фильтрации
        let allSame = true;
        let hasWarn = false;
        let isEmpty = true;
        
        if (row1 && row2) {
            // Сравниваем строки для определения allSame и hasWarn
            for (let c = 0; c < realHeaders.length; c++) {
                const v1 = row1[c] !== undefined ? row1[c] : '';
                const v2 = row2[c] !== undefined ? row2[c] : '';
                
                if ((v1 && v1.toString().trim() !== '') || (v2 && v2.toString().trim() !== '')) {
                    isEmpty = false;
                }
                
                if (v1.toString().toUpperCase().trim() !== v2.toString().toUpperCase().trim()) {
                    allSame = false;
                    hasWarn = true;
                }
            }
        } else {
            allSame = false;
            hasWarn = false;
            
            // Проверяем на пустоту
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
        
        // Применяем фильтры
        if (isEmpty) return;
        if (hideSame && row1 && row2 && allSame) return;
        if (hideNewRows1 && row1 && !row2) return;
        if (hideNewRows2 && !row1 && row2) return;
        if (hideDiffRows && row1 && row2 && hasWarn) return;
        
        // Для SIMILAR пар создаем группированные строки (как в functions.js)
        if (pair.matchType === 'SIMILAR' && row1 && row2) {
            // Анализируем колонки на предмет различий
            const columnComparisons = [];
            let hasAnyDifference = false;
            
            realHeaders.forEach((header, colIndex) => {
                const v1 = row1[colIndex] !== undefined ? row1[colIndex] : '';
                const v2 = row2[colIndex] !== undefined ? row2[colIndex] : '';
                
                if (v1 && v2 && v1.toString().toUpperCase().trim() === v2.toString().toUpperCase().trim()) {
                    columnComparisons[colIndex] = 'identical';
                } else {
                    columnComparisons[colIndex] = 'different';
                    hasAnyDifference = true;
                }
            });
            
            if (hasAnyDifference) {
                // Первая строка (File 1)
                bodyHtml += `<tr class="warn-row warn-row-group-start" data-row-index="${pair.index1}">`;
                bodyHtml += `<td class="warn-cell">${file1Name}</td>`;
                
                realHeaders.forEach((header, colIndex) => {
                    const v1 = row1[colIndex] !== undefined ? row1[colIndex] : '';
                    const compResult = columnComparisons[colIndex];
                    
                    if (compResult === 'identical') {
                        // Объединяем одинаковые колонки
                        bodyHtml += `<td class="identical" rowspan="2" style="vertical-align: middle; text-align: center;">${v1}</td>`;
                    } else {
                        // Разные колонки показываем отдельно
                        bodyHtml += `<td class="warn-cell">${v1}</td>`;
                    }
                });
                bodyHtml += '</tr>';
                
                // Вторая строка (File 2)
                bodyHtml += `<tr class="warn-row warn-row-group-end" data-row-index="${pair.index2}">`;
                bodyHtml += `<td class="warn-cell">${file2Name}</td>`;
                
                realHeaders.forEach((header, colIndex) => {
                    const v2 = row2[colIndex] !== undefined ? row2[colIndex] : '';
                    const compResult = columnComparisons[colIndex];
                    
                    if (compResult === 'different') {
                        // Показываем только разные колонки (identical уже показаны с rowspan)
                        bodyHtml += `<td class="warn-cell">${v2}</td>`;
                    }
                    // identical колонки пропускаем, так как они уже отображены с rowspan="2"
                });
                bodyHtml += '</tr>';
            } else {
                // Если нет различий, показываем как identical
                bodyHtml += `<tr class="identical-row" data-row-index="${pair.index1}">`;
                bodyHtml += `<td class="file-both">Both files</td>`;
                
                realHeaders.forEach((header, colIndex) => {
                    const value = row1[colIndex] !== undefined ? row1[colIndex] : '';
                    bodyHtml += `<td class="identical" title="${value}">${value}</td>`;
                });
                bodyHtml += `</tr>`;
            }
        } else {
            // Для других типов (IDENTICAL, ONLY_IN_TABLE1, ONLY_IN_TABLE2) используем старую логику
            const rowData = row1 || row2;
            let fileIndicator = '';
            
            // Определяем индикатор источника в зависимости от типа совпадения
            if (pair.matchType === 'IDENTICAL') {
                fileIndicator = 'Both files';
            } else if (pair.onlyIn === 'table1') {
                fileIndicator = file1Name;
            } else if (pair.onlyIn === 'table2') {
                fileIndicator = file2Name;
            }
            
            let rowClass = 'diff-row';
            let sourceClass = 'file-indicator';
            
            // Устанавливаем стили в зависимости от типа совпадения
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
                // Fallback для других случаев
                rowClass += ' identical-row';
                sourceClass += ' file-both';
            }
            
            bodyHtml += `<tr class="${rowClass}" data-row-index="${pair.index1 >= 0 ? pair.index1 : pair.index2}">`;
            bodyHtml += `<td class="${sourceClass}">${fileIndicator}</td>`;
            
            realHeaders.forEach((header, colIndex) => {
                const value = rowData && rowData[colIndex] !== undefined ? rowData[colIndex] : '';
                
                // Улучшенная логика для определения стиля ячейки
                let cellClass = 'diff-cell';
                if (pair.matchType === 'IDENTICAL') {
                    cellClass += ' identical';
                } else {
                    cellClass += ' different';
                }
                
                bodyHtml += `<td class="${cellClass}" title="${value}">${value}</td>`;
            });
            
            bodyHtml += `</tr>`;
        }
    });
    
    if (pairs.length > 1000) {
        bodyHtml += `<tr class="info-row"><td colspan="${realHeaders.length + 1}" style="text-align:center; padding:10px; background:#f8f9fa; font-style: italic;">⚠️ Showing first 1,000 of ${pairs.length} differences</td></tr>`;
    }
    
    bodyTable.innerHTML = bodyHtml;
    
    const diffTableFinal = document.getElementById('diffTable');
    if (diffTableFinal) {
        diffTableFinal.style.display = 'block';
        diffTableFinal.style.visibility = 'visible';
    }
}

window.MaxPilotDuckDB = {
    FastTableComparator,
    fastComparator,
    initializeFastComparator,
    compareTablesWithFastComparator,
    compareTablesEnhanced,
    processFastComparisonResults,
    prepareDataForExportFast,
    benchmarkExportPerformance,
    runFastComparatorTests,
    testExportPerformance: () => window.testExportPerformance(),
    version: '1.2.0'
};


document.addEventListener('DOMContentLoaded', async () => {
    await initializeFastComparator();
    
    if (typeof window.compareTables === 'function') {
        window.compareTablesOriginal = window.compareTables;
    }
    // Включаем обратно DuckDB WASM с улучшенной логикой
    window.compareTables = compareTablesEnhanced;
    
    // Добавляем обработчики событий для фильтров чекбоксов
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
        console.log('ℹ️ No fast comparison result available for benchmarking');
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
    
    console.log('📊 Export Performance Analysis:', {
        totalInputRows: totalDataRows.toLocaleString(),
        rowsToExport: exportRows.toLocaleString(),
        breakdown: {
            identical: identical.length.toLocaleString(),
            onlyInFile1: onlyInTable1.length.toLocaleString(),  
            onlyInFile2: onlyInTable2.length.toLocaleString()
        },
        fastModeActive: !!(fastComparator && fastComparator.initialized),
        processingMode: fastComparator?.mode || 'standard',
        expectedSpeedup: expectedSpeedup
    });
    
    if (exportRows > 20000) {
        console.log('💡 Large export detected - fast processing will show significant improvement');
    } else if (exportRows > 5000) {
        console.log('💡 Medium export - expect moderate performance improvement');  
    }
    
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
    console.log('🧪 Starting export performance test...');
    
    if (!window.currentFastResult) {
        console.log('❌ No comparison result available. Please run a comparison first.');
        return;
    }
    
    const benchmark = await benchmarkExportPerformance();
    if (!benchmark) return;
    
    console.log('⏱️ Testing data preparation performance...');
    
    const startTime = performance.now();
    const testData = await prepareDataForExportFast(window.currentFastResult, false);
    const endTime = performance.now();
    
    const duration = endTime - startTime;
    const rowsPerSecond = Math.round(benchmark.exportRows / (duration / 1000));
    
    console.log('✅ Performance Test Results:', {
        dataPreparationTime: `${duration.toFixed(2)}ms`,
        rowsPerSecond: rowsPerSecond.toLocaleString(),
        exportRows: benchmark.exportRows.toLocaleString(),
        efficiency: rowsPerSecond > 10000 ? 'Excellent' : 
                   rowsPerSecond > 5000 ? 'Good' : 
                   rowsPerSecond > 1000 ? 'Fair' : 'Needs improvement'
    });
    
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
        console.log('Fast export not available - falling back to standard method');
        return null;
    }

    console.log('⚡ Starting optimized fast export processing...', {
        identical: fastResult.identical.length,
        onlyInTable1: fastResult.onlyInTable1.length,
        onlyInTable2: fastResult.onlyInTable2.length,
        totalRows: fastResult.table1Count + fastResult.table2Count
    });

    const startTime = performance.now();
    const { identical, onlyInTable1, onlyInTable2, commonColumns, alignedData1, alignedData2 } = fastResult;
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
    // Используем getFileDisplayName для отображения имени файла с листом
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
    
    // Process identical rows (limit to prevent huge files)
    const maxIdentical = Math.min(identical.length, identical.length > 10000 ? 500 : 2000);
    
    for (let batchStart = 0; batchStart < maxIdentical; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, maxIdentical);
        const batch = identical.slice(batchStart, batchEnd);
        
        batch.forEach(identicalPair => {
            const row1Index = identicalPair.row1 - 1;
            const row1 = workingData1[row1Index + 1];
            
            if (!row1) return;
            
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
    
    // Process table1-only rows efficiently
    for (let batchStart = 0; batchStart < onlyInTable1.length; batchStart += BATCH_SIZE) {
        const batchEnd = Math.min(batchStart + BATCH_SIZE, onlyInTable1.length);
        const batch = onlyInTable1.slice(batchStart, batchEnd);
        
        batch.forEach(diff => {
            const row1Index = diff.row1 - 1;
            const row1 = workingData1[row1Index + 1];
            
            if (!row1) return;
            
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
            const row2Index = diff.row2 - 1;
            const row2 = workingData2[row2Index + 1];
            
            if (!row2) return;
            
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
    console.log(`⚡ Fast export completed in ${duration.toFixed(2)}ms - prepared ${data.length - 1} rows for export`);
    
    return { data, formatting, colWidths };
}

}