const XLSX = require('xlsx');
const fs = require('fs');
const JSONStream = require('JSONStream');
const { Transform, Writable } = require('stream');

class JsonToExcel {
  constructor(options = {}) {
    // Check if jsonFilePath is provided and it's a string with a .json extension
    if (!options.jsonFilePath || typeof options.jsonFilePath !== 'string' || !options.jsonFilePath.trim().toLowerCase().endsWith('.json')) {
      throw new Error('Invalid or missing JSON file path.');
    }
    this.jsonFilePath = options.jsonFilePath.trim();
    
    // Check if outputFileName is provided and it's a string with a .xlsx extension
    if (!options.outputFileName || typeof options.outputFileName !== 'string' || !options.outputFileName.trim().toLowerCase().endsWith('.xlsx')) {
      throw new Error('Invalid or missing output file name.');
    }
    this.outputFileName = options.outputFileName.trim() || 'output.xlsx';

    // Check if batchSize is provided and it's a positive number
    if (typeof options.batchSize !== 'number' || options.batchSize <= 0) {
      throw new Error('Invalid or missing batch size.');
    }
    this.batchSize = options.batchSize || 100;

    // Check if headers is provided and it's an array
    if (!Array.isArray(options.headers) || options.headers.length === 0) {
      throw new Error('Invalid or missing headers.');
    }
    this.headers = options.headers;

    // Check if headers is provided and it's an array
    if (typeof options.sheetName !== 'string' || options.sheetName === '') {
        throw new Error('Invalid sheetName, it should be string');
      }
    // Set the sheet name, default to 'Data' if not provided
    this.sheetName = options.sheetName || 'Data';
  }

  convert() {
    const jsonStream = fs.createReadStream(this.jsonFilePath);
    const excelWriter = new ExcelWriter({
      headers: this.headers,
      outputFileName: this.outputFileName,
      sheetName: this.sheetName // Pass the sheet name option
    });

    const excelTransform = new JsonToExcelTransform({ batchSize: this.batchSize });

    jsonStream.on('error', (err) => {
      console.error('Error reading JSON file:', err);
    });

    jsonStream.pipe(JSONStream.parse('*')).pipe(excelTransform).pipe(excelWriter);

    excelWriter.on('finish', () => {
      console.log('Excel stream created successfully.');
    });

    excelWriter.on('error', (err) => {
      console.error('Error writing Excel stream:', err);
    });
  }
}

class ExcelWriter extends Writable {
  constructor(options) {
    super({ objectMode: true });
    this.wb = XLSX.utils.book_new();
    this.ws = XLSX.utils.aoa_to_sheet([options.headers]);
    XLSX.utils.book_append_sheet(this.wb, this.ws, options.sheetName || 'Data'); // Use provided sheet name or default to 'Data'
    this.outputFileName = options.outputFileName ;
  }

  _write(chunk, encoding, callback) {
    XLSX.utils.sheet_add_aoa(this.ws, chunk, { origin: -1 });
    callback();
  }

  _final(callback) {
    XLSX.writeFile(this.wb, this.outputFileName);
    callback();
  }
}

class JsonToExcelTransform extends Transform {
  constructor(options) {
    super({ objectMode: true });
    this.buffer = [];
    this.batchSize = options.batchSize;
    this.dataCount = 0;
  }

  _transform(data, encoding, callback) {
    const rowData = Object.values(data);
    this.buffer.push(rowData);
    this.dataCount++;

    if (this.buffer.length >= this.batchSize) {
      this.push([...this.buffer]);
      this.buffer = [];
      console.log(`Added ${this.batchSize} data items to Excel sheet.`);
    }

    callback();
  }

  _flush(callback) {
    if (this.buffer.length > 0) {
      this.push([...this.buffer]);
      console.log(`Added ${this.buffer.length} data items to Excel sheet.`);
    }
    callback();
  }
}

module.exports = JsonToExcel;
