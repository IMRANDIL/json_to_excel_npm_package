# @imrandil/JsonToExcel

A simple npm package for generating Excel files (xlsx) from JSON arrays.

## Installation

```bash
npm install @imrandil/JsonToExcel
```

## Usage

```javascript
const JsonToExcel = require('@imrandil/JsonToExcel');

// Example options
const options = {
  jsonFilePath: 'path/to/your/json/file.json',
  outputFileName: 'output.xlsx',
  batchSize: 100,
  headers: ['Header1', 'Header2', 'Header3'],
  sheetName: 'Sheet1'
};

// Create an instance of JsonToExcel with options
const converter = new JsonToExcel(options);

// Convert JSON data to Excel
converter.convert();
```

## Options

- `jsonFilePath`: The file path to the JSON data to be converted.
- `outputFileName`: The name of the Excel file to be generated.
- `batchSize`: The number of JSON data items to process before writing to the Excel file (default: 100).
- `headers`: An array of strings representing the headers for the Excel file.
- `sheetName`: The name of the sheet in the Excel file (default: 'Data').

## License

MIT

## Author

Ali Imran Adil

## Repository

[GitHub Repository](https://github.com/IMRANDIL/json_to_excel_npm_package)

## Bugs

[Issue Tracker](https://github.com/IMRANDIL/json_to_excel_npm_package/issues)