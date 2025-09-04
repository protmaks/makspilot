// Реальная интеграция DuckDB WASM через esm.sh
class DuckDBWASMLoader {
    constructor() {
        this.db = null;
        this.initialized = false;
    }

    async initialize() {
        // Сохраняем оригинальные функции console перед try блоком
        const originalConsoleLog = console.log;
        const originalConsoleWarn = console.warn;
        const originalConsoleInfo = console.info;
        
        try {
            //console.log('🔧 Loading DuckDB WASM from ESM CDN...');
            
            // Перехватываем console.log для подавления DuckDB логов
            console.log = function(...args) {
                // Фильтруем сообщения от DuckDB
                const message = args.join(' ');
                if (message && typeof args[0] === 'object' && 
                    args[0].timestamp && args[0].level !== undefined && 
                    args[0].origin !== undefined && args[0].topic !== undefined) {
                    return; // Подавляем DuckDB логи
                }
                originalConsoleLog.apply(console, args);
            };
            
            console.warn = function(...args) {
                const message = args.join(' ');
                if (message && typeof args[0] === 'object' && 
                    args[0].timestamp && args[0].level !== undefined) {
                    return;
                }
                originalConsoleWarn.apply(console, args);
            };
            
            console.info = function(...args) {
                const message = args.join(' ');
                if (message && typeof args[0] === 'object' && 
                    args[0].timestamp && args[0].level !== undefined) {
                    return;
                }
                originalConsoleInfo.apply(console, args);
            };
            
            // Используем esm.sh который автоматически разрешает зависимости
            const duckdb = await import('https://esm.sh/@duckdb/duckdb-wasm@1.29.0');
            
            //console.log('📦 DuckDB module loaded, initializing...');
            
            // Создаем кастомные бандлы с локальными файлами
            const MANUAL_BUNDLES = {
                eh: {
                    mainModule: '/javascript/duckdb/duckdb-eh.wasm',
                    mainWorker: '/javascript/duckdb/duckdb-browser-eh.worker.js',
                    pthreadWorker: null
                }
            };
            
            const bundle = MANUAL_BUNDLES.eh;
            
            const worker = new Worker(bundle.mainWorker);
            
            // Создаем собственный тихий логгер
            const logger = {
                log: () => {},
                warn: () => {},
                error: () => {},
                info: () => {},
                debug: () => {}
            };
            
            this.db = new duckdb.AsyncDuckDB(logger, worker);
            await this.db.instantiate(bundle.mainModule);
            
            // Отключаем все виды логирования
            try {
                await this.db.query('SET log_query_path = \'\';');
                await this.db.query('SET enable_progress_bar = false;');
                await this.db.query('SET enable_print_progress = false;');
            } catch (e) {
                // Игнорируем ошибки настройки логирования
            }
            
            // Создаем соединение для выполнения запросов
            this.connection = await this.db.connect();
            
            // Восстанавливаем оригинальные функции console
            console.log = originalConsoleLog;
            console.warn = originalConsoleWarn;
            console.info = originalConsoleInfo;
            
            //console.log('✅ DuckDB WASM initialized successfully with local files!');
            this.initialized = true;
            
            return this.db;
            
        } catch (error) {
            // Восстанавливаем оригинальные функции console в случае ошибки
            console.log = originalConsoleLog;
            console.warn = originalConsoleWarn;
            console.info = originalConsoleInfo;
            
            console.error('❌ Failed to initialize DuckDB WASM:', error);
            
            // Пробуем упрощенный подход
            try {
                //console.log('🔄 Trying simplified DuckDB initialization...');
                
                const duckdb = await import('https://cdn.skypack.dev/@duckdb/duckdb-wasm');
                
                // Создаем упрощенную инициализацию
                this.db = await duckdb.default.DuckDBDataProtocol.initialize();
                this.initialized = true;
                
                //console.log('✅ DuckDB WASM initialized (simplified mode)');
                return this.db;
                
            } catch (fallbackError) {
                console.error('❌ Simplified initialization also failed:', fallbackError);
                throw new Error('All DuckDB initialization methods failed');
            }
        }
    }

    async query(sql) {
        if (!this.initialized || !this.connection) {
            throw new Error('DuckDB not initialized or no connection available');
        }
        
        try {
            //console.log('🔍 Executing query:', sql);
            const result = await this.connection.query(sql);
            return result;
        } catch (error) {
            console.error('❌ Query failed:', error);
            throw error;
        }
    }

    async close() {
        if (this.connection) {
            await this.connection.close();
            this.connection = null;
        }
        if (this.db) {
            await this.db.close();
            this.db = null;
        }
        this.initialized = false;
    }
}

// Экспортируем глобально
window.DuckDBWASMLoader = DuckDBWASMLoader;

// Создаем инстанс для немедленного использования
window.duckdbLoader = new DuckDBWASMLoader();

//console.log('✅ DuckDB WASM loader ready');
