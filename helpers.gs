

function getLastColumnForRow(row) {
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  let lastColumn = 1;
  
  for (let i = rowData.length - 1; i >= 0; i--) {
    if (rowData[i] !== "") {
      lastColumn = i + 1;
      break;
    }
  }
  
  return lastColumn;
}

function getHeaders() {
  var lastCol = getLastColumnForRow(headerRow)
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const prompts = sheet.getRange(promptRow, 1, 1, lastCol).getValues()[0];
  const colLetters = getColLetters(headers.length);
  const headerDict = {};
  const headerIndices = {};
  
  for (let i = 0; i < headers.length; i++) {
    headerDict[i+1] = [headers[i], prompts[i]];
    headerIndices[headers[i]] = i+1
  }

  return [headerDict, headerIndices];
}

function getColLetters(numCols) {
  const letters = [];
  for (let i = 0; i < numCols; i++) {
    let letter = '';
    let n = i;
    while (n >= 0) {
      letter = String.fromCharCode(65 + (n % 26)) + letter;
      n = Math.floor(n / 26) - 1;
    }
    letters.push(letter);
  }
  return letters;
}

function getRow(rowIndex) {
  const headers = sheet.getRange(5, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowDict = {};

  for (let i = 0; i < headers.length; i++) {
    rowDict[headers[i].toLowerCase()] = data[i];
  }
  return rowDict;
}
