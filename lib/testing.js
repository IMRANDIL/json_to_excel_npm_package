const JsonToExcel = require('./index')

// Define the headers
const headers = [
    "Time", "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8", "V9",
    "V10", "V11", "V12", "V13", "V14", "V15", "V16", "V17", "V18",
    "V19", "V20", "V21", "V22", "V23", "V24", "V25", "V26", "V27",
    "V28", "Amount", "Class"
  ];

const options = {
  jsonFilePath: 'temp2.json',
  outputFileName: 'customOutput.xlsx',
  batchSize: 500,
  headers: headers,
  sheetName: 'test'
};

const converter = new JsonToExcel(options);
converter.convert();


