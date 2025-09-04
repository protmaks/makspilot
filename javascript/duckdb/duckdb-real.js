// Реальная интеграция DuckDB WASM через esm.sh
class DuckDBWASMLoader {
    constructor() {
        this.db = null;
        this.initialized = false;
    }

    async initialize() {
        try {
            console.log('🔧 Loading DuckDB WASM from ESM CDN...');
            
            // Используем esm.sh который автоматически разрешает зависимости
            const duckdb = await import('https://esm.sh/@duckdb/duckdb-wasm@1.29.0');
            
            console.log('📦 DuckDB module loaded, initializing...');
            
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
            const logger = new duckdb.ConsoleLogger();
            
            this.db = new duckdb.AsyncDuckDB(logger, worker);
            await this.db.instantiate(bundle.mainModule);
            
            // Создаем соединение для выполнения запросов
            this.connection = await this.db.connect();
            
            console.log('✅ DuckDB WASM initialized successfully with local files!');
            this.initialized = true;
            
            return this.db;
            
        } catch (error) {
            console.error('❌ Failed to initialize DuckDB WASM:', error);
            
            // Пробуем упрощенный подход
            try {
                console.log('🔄 Trying simplified DuckDB initialization...');
                
                const duckdb = await import('https://cdn.skypack.dev/@duckdb/duckdb-wasm');
                
                // Создаем упрощенную инициализацию
                this.db = await duckdb.default.DuckDBDataProtocol.initialize();
                this.initialized = true;
                
                console.log('✅ DuckDB WASM initialized (simplified mode)');
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
            console.log('🔍 Executing query:', sql);
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

console.log('✅ DuckDB WASM loader ready');
