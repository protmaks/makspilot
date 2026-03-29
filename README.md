# MaksPilot - Excel/CSV Comparison Tool

Professional spreadsheet comparison tool. Compare Excel, CSV files online with advanced features. Free, secure, and fast analysis.

## Features
- **Privacy First**: All processing happens in your browser. Data is never sent to a server.
- **Large File Support**: Optimized for large datasets using DuckDB-WASM and Web Workers.
- **Fuzzy Matching**: Intelligent comparison for dates and numbers with customizable tolerance.
- **Multi-language**: Support for English, Russian, Spanish, German, Japanese, Portuguese, Chinese, and Polish.

## Development

### Prerequisites
- Node.js (for running tests and installing dependencies)
- Python 3 (for the local development server)

### Local Setup
To run the project locally without a complex build process:

1. Clone the repository.
2. Start the local development server:
   ```bash
   python3 makspilot-server.py 8124
   ```
3. Open your browser at `http://localhost:8124`.

### Testing
We use [Vitest](https://vitest.dev/) for unit testing our data transformation and comparison logic.

1. Install dependencies:
   ```bash
   npm install
   ```
2. Run tests:
   ```bash
   npm test
   ```

### Project Structure
- `/compare`: Main comparison tool interface.
- `/javascript`: Core application logic.
  - `functions.js`: Data normalization and comparison algorithms.
  - `duckdb/`: DuckDB-WASM integration for high-performance processing.
- `/tests`: Automated test suites.
- `/examples`: Sample Excel and CSV files for testing.

## License
ISC
