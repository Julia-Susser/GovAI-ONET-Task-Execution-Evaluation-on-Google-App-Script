const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var headerRow = 5
var promptRow = 4
var x = getHeaders();
const headers = x[0]
const headerIndices = x[1]
const apiKey = 'sk-proj-igFEJZzOcItDWNpkGwdMT3BlbkFJtVhZkiy6PzWV71LT5GSE'; // Global variable


function fixPrompt(prompt, rowDict) {
  return prompt.replace(/\$(\w+)/g, (match, p1) => rowDict[p1] || match);
}

function setCell(col, row, response) {
  if (response && response.length > 0 && response[0].length > 0) {
    if (response.length > 1) {
      sheet.insertRowsAfter(row, response.length - 1);
      var numColumns = sheet.getMaxColumns(); 
      var originalValues = sheet.getRange(row, 1, 1, numColumns).getValues();
      for (var i = 1; i <= response.length - 1; i++) {
        sheet.getRange(row + i, 1, 1, numColumns).setValues(originalValues);
      }
    } 
  }
  console.log(row, col, response.length, response[0].length)
  sheet.getRange(row, col, response.length, response[0].length).setValues(response);

}

// function mockRun(col, row){
//   var colLetter = "E"
//   var row = 6
//   range = sheet.getRange(colLetter + row.toString());
//   var col = range.getColumn()
//   var response = runCell(col, row)
//   setCell(col, row, response)
// }

function runCell(col,row) {
  const header = headers[col]
  const type = header[0];
  const promptTemplate = header[1];
  const rowDict = getRow(row);

  if (type === "Evaluation Rating") {
    const prompt = sheet.getRange("B2").getValue(); // Get prompt from B2
  } else if (["Occupation", "Task ID", "Task", "Subtask index", "Subtask", "Evaluation Reasoning"].includes(header)) {
    return;
  } 
  const prompt = fixPrompt(promptTemplate, rowDict);

  console.log(rowDict)
  Logger.log(prompt)
  const response = callOpenAI(prompt);
  Logger.log(response)
  if (type === "Subtask decomposition") {
    var subtasks = response.split(/\d+\./).map(s => s.trim()).filter(Boolean);
    var responseList = subtasks.map((subtask, index) => [response, index+1, subtask]);
    return responseList
  } else if (type === "Evaluation Rating"){
    return [[response, response]]
  } else {
    return [[response]]
  } 
}





function runCompletion(minRow, maxRow) {
  minCol = headerIndices["Results format"]
  maxCol = minCol + 1 //Subtask completion
  for (let row = minRow; row <= maxRow; row++){
    for (let col = minCol; col <= maxCol; col++) {
      const cell = sheet.getRange(row, col);
      col = cell.getColumn()
      row = cell.getRow();
      var response = runCell(col, row)
      setCell(col, row, response)
    }
  }
}

function runEvaluation(minRow, maxRow) {
  minCol = headerIndices["Evaluation Rating"]
  maxCol = minCol //evaluation reasoning does not need to be run separately
  for (let row = minRow; row <= maxRow; row++){
    for (let col = minCol; col <= maxCol; col++) {
      const cell = sheet.getRange(row, col);
      col = cell.getColumn()
      row = cell.getRow();
      var response = runCell(col, row)
      setCell(col, row, response)
    }
  }
}




function runContextAndSubtaskDecomposition(row){
  //context 
  col = headerIndices["Context"]
  var response = runCell(col, row)
  setCell(col, row, response)


  //subtask decompositon
  col = headerIndices["Subtask decomposition"]
  var response = runCell(col, row)
  setCell(col, row, response)
  maxRow = row+response.length-1
  minCol = col+response[0].length
  return [minCol, maxRow]
}


function resetTaskID(row){
  var col = headerIndices["Task ID"]; 
  var taskID = sheet.getRange(row, col).getValue();
  if (taskID == ""){
    return;
  }
  var startRow = Math.max(1, row - 10);
  var endRow = Math.min(sheet.getLastRow(), row + 10); 
  var data = sheet.getRange(startRow, col, endRow - startRow + 1).getValues();
  var firstRowIndex = -1;
  var lastRowIndex = -1;
  data.forEach(function(value, index) {
    if (value[0] == taskID) {
      if (firstRowIndex === -1) {
        firstRowIndex = startRow + index; 
      }
      lastRowIndex = startRow + index; 
    }
  });

  if (lastRowIndex > firstRowIndex) {
    sheet.deleteRows(firstRowIndex + 1, lastRowIndex - firstRowIndex); 
  }
  var col = headerIndices["Task"] + 1;
  sheet.getRange(firstRowIndex, col, firstRowIndex, sheet.getMaxColumns() - col + 1).clearContent();
  return firstRowIndex;
}

function run(){
  var cell = sheet.getActiveRange();
  var row = cell.getRow(); 
  var row = 6
  resetTaskID(row)
  const maxCol = getLastColumnForRow(headerRow); 
  var x = runContextAndSubtaskDecomposition(row)
  minCol = x[0]
  maxRow = x[1]
  minRow = row
  runCompletion(minRow, maxRow)
  runEvaluation(minRow, maxRow)

}


function runOneCell() {
  const cell = sheet.getActiveCell();
  var col = cell.getColumn()
  var row = cell.getRow();
  
  const header = headers[col]
  const type = header[0];
  if (type == "Subtask Decomposition"){
    return; //can't create new rows when run for one cell
  }

  var response = runCell(col, row)
  if (response.length > 1){
    return; //can't create new rows when run for one cell
  }

  setCell(col, row, response)
}



function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('OpenAI')
    .addItem('Run row', 'run')
    .addItem('Run cell', 'runOneCell')
    .addToUi();
}
