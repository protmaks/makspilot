/**
 * Test suite for DuckDB WASM integration
 * Run in browser console to test functionality
 */

class DuckDBTester {
    constructor() {
        this.testResults = [];
    }

    async runAllTests() {
        console.log('🧪 Starting DuckDB WASM Tests...\n');
        
        const tests = [
            { name: 'Browser Support', fn: this.testBrowserSupport },
            { name: 'DuckDB Initialization', fn: this.testInitialization },
            { name: 'Table Creation', fn: this.testTableCreation },
            { name: 'Basic Comparison', fn: this.testBasicComparison },
            { name: 'Large Dataset', fn: this.testLargeDataset },
            { name: 'Tolerance Comparison', fn: this.testToleranceComparison },
            { name: 'Performance Benchmark', fn: this.testPerformance }
        ];

        for (const test of tests) {
            try {
                console.log(`\n🔬 Testing: ${test.name}`);
                const result = await test.fn.call(this);
                this.testResults.push({ name: test.name, status: 'PASS', result });
                console.log(`✅ ${test.name}: PASS`);
            } catch (error) {
                this.testResults.push({ name: test.name, status: 'FAIL', error: error.message });
                console.error(`❌ ${test.name}: FAIL - ${error.message}`);
            }
        }

        this.printSummary();
    }

    async testBrowserSupport() {
        const isSupported = await DuckDBManager.isSupported();
        if (!isSupported) {
            throw new Error('Browser does not support DuckDB WASM');
        }
        return 'Browser supports DuckDB WASM';
    }

    async testInitialization() {
        const manager = new DuckDBManager();
        const initialized = await manager.initialize();
        
        if (!initialized) {
            throw new Error('Failed to initialize DuckDB');
        }
        
        await manager.close();
        return 'DuckDB initialized successfully';
    }

    async testTableCreation() {
        const manager = new DuckDBManager();
        await manager.initialize();

        const testData = [
            ['id', 'name', 'value'],
            [1, 'Alice', 100],
            [2, 'Bob', 200],
            [3, 'Charlie', 300]
        ];

        await manager.createTableFromData('test_table', testData);
        const data = await manager.getTableData('test_table');
        
        if (data.length !== 3) {
            throw new Error(`Expected 3 rows, got ${data.length}`);
        }

        await manager.close();
        return `Created table with ${data.length} rows`;
    }

    async testBasicComparison() {
        const manager = new DuckDBManager();
        await manager.initialize();

        const data1 = [
            ['id', 'name', 'value'],
            [1, 'Alice', 100],
            [2, 'Bob', 200],
            [3, 'Charlie', 300]
        ];

        const data2 = [
            ['id', 'name', 'value'],
            [1, 'Alice', 100],  // identical
            [2, 'Bob', 250],    // different
            [4, 'David', 400]   // only in data2
        ];

        await manager.createTableFromData('table1', data1);
        await manager.createTableFromData('table2', data2);
        
        const result = await manager.compareTablesFast('table1', 'table2');
        
        if (result.identical.length !== 1) {
            throw new Error(`Expected 1 identical row, got ${result.identical.length}`);
        }

        await manager.close();
        return `Comparison found ${result.identical.length} identical, ${result.onlyInTable1.length} unique to table1, ${result.onlyInTable2.length} unique to table2`;
    }

    async testLargeDataset() {
        const manager = new DuckDBManager();
        await manager.initialize();

        // Generate larger test dataset
        const size = 1000;
        const data1 = [['id', 'name', 'value', 'category']];
        const data2 = [['id', 'name', 'value', 'category']];

        for (let i = 1; i <= size; i++) {
            data1.push([i, `User${i}`, i * 10, `Category${i % 5}`]);
            // Make data2 slightly different
            data2.push([i, `User${i}`, i * 10 + (i % 3), `Category${i % 5}`]);
        }

        const startTime = performance.now();
        
        await manager.createTableFromData('large_table1', data1);
        await manager.createTableFromData('large_table2', data2);
        
        const result = await manager.compareTablesFast('large_table1', 'large_table2');
        
        const endTime = performance.now();
        const duration = endTime - startTime;

        await manager.close();
        return `Processed ${size} rows in ${duration.toFixed(2)}ms`;
    }

    async testToleranceComparison() {
        const manager = new DuckDBManager();
        await manager.initialize();

        const data1 = [
            ['price', 'quantity'],
            [100.00, 5],
            [200.50, 10],
            [300.75, 15]
        ];

        const data2 = [
            ['price', 'quantity'],
            [100.50, 5],    // 0.5% difference
            [202.00, 10],   // 0.75% difference  
            [310.00, 15]    // 3.4% difference
        ];

        await manager.createTableFromData('tolerance1', data1);
        await manager.createTableFromData('tolerance2', data2);
        
        const differences = await manager.findRowDifferencesWithTolerance('tolerance1', 'tolerance2', ['price', 'quantity'], 1.5);
        
        // Should find differences where tolerance exceeds 1.5%
        await manager.close();
        return `Tolerance comparison found ${differences.length} rows exceeding 1.5% tolerance`;
    }

    async testPerformance() {
        const sizes = [100, 500, 1000];
        const results = {};

        for (const size of sizes) {
            const manager = new DuckDBManager();
            await manager.initialize();

            // Generate test data
            const data1 = [['id', 'value']];
            const data2 = [['id', 'value']];
            
            for (let i = 1; i <= size; i++) {
                data1.push([i, Math.random() * 1000]);
                data2.push([i, Math.random() * 1000]);
            }

            const startTime = performance.now();
            
            await manager.createTableFromData('perf1', data1);
            await manager.createTableFromData('perf2', data2);
            await manager.compareTablesFast('perf1', 'perf2');
            
            const endTime = performance.now();
            results[size] = endTime - startTime;

            await manager.close();
        }

        return `Performance: ${Object.entries(results).map(([size, time]) => `${size} rows: ${time.toFixed(2)}ms`).join(', ')}`;
    }

    printSummary() {
        console.log('\n📊 Test Summary:');
        console.log('================');
        
        const passed = this.testResults.filter(r => r.status === 'PASS').length;
        const failed = this.testResults.filter(r => r.status === 'FAIL').length;
        
        console.log(`✅ Passed: ${passed}`);
        console.log(`❌ Failed: ${failed}`);
        console.log(`📈 Success Rate: ${((passed / this.testResults.length) * 100).toFixed(1)}%`);
        
        if (failed > 0) {
            console.log('\n🐛 Failed Tests:');
            this.testResults
                .filter(r => r.status === 'FAIL')
                .forEach(test => console.log(`   - ${test.name}: ${test.error}`));
        }

        console.log('\n📝 Detailed Results:');
        this.testResults.forEach(test => {
            const status = test.status === 'PASS' ? '✅' : '❌';
            const details = test.status === 'PASS' ? test.result : test.error;
            console.log(`   ${status} ${test.name}: ${details}`);
        });
    }
}

// Utility function to generate test data
function generateTestData(rows, cols) {
    const data = [];
    const headers = Array.from({length: cols}, (_, i) => `col_${i}`);
    data.push(headers);
    
    for (let i = 0; i < rows; i++) {
        const row = [];
        for (let j = 0; j < cols; j++) {
            row.push(Math.random() * 1000);
        }
        data.push(row);
    }
    
    return data;
}

// Run tests if in browser environment
if (typeof window !== 'undefined') {
    window.DuckDBTester = DuckDBTester;
    window.generateTestData = generateTestData;
    
    // Add test button to UI for easy access
    function addTestButton() {
        const button = document.createElement('button');
        button.textContent = '🧪 Run DuckDB Tests';
        button.style.cssText = `
            position: fixed;
            top: 10px;
            right: 10px;
            z-index: 10000;
            padding: 8px 16px;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        `;
        
        button.onclick = async () => {
            const tester = new DuckDBTester();
            await tester.runAllTests();
        };
        
        document.body.appendChild(button);
    }
    
    // Add test button when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', addTestButton);
    } else {
        addTestButton();
    }
}

// Export for Node.js testing
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { DuckDBTester, generateTestData };
}
