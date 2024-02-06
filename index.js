const XLSX = require('xlsx');
const process = require('process');

const args = process.argv.slice(2);
const [filePath, sheetName, filterExpression, calculation] = args;
if (!filePath || !sheetName) {
  console.log('Usage: node index.js <filePath> <sheetName> [filterExpression] [calculation]');
  process.exit(1);
}

const workbook = XLSX.readFile(filePath);
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet);

const filteredData = data.filter((row) => {
  // filter as filterExpression, like:  "row['c1'] === 1 && row['c2'] === 'yikf'"
  if (!filterExpression) {
    return true;
  } else {
    return eval(filterExpression);
  }
});

function printAsTable(data) {
    console.table(data);
}

if (calculation) {
  const match = calculation.match(/(\w+)\((\w+)\)/);
  if (match) {
    const [, func, column] = match;
    let result;
    switch (func.toLowerCase()) {
      case 'sum':
        result = filteredData.reduce((sum, row) => sum + (Number(row[column]) || 0), 0);
        printAsTable(result);
        break;
      default:
        console.log(`Unsupported calculation: ${calculation}`);
    }
  } else {
    console.log(`Invalid calculation expression: ${calculation}`);
  }
} else {
  printAsTable(filteredData);
}