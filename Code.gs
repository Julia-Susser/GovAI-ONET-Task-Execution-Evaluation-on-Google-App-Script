const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var headerRow = 5
var promptRow = 4
var x = getHeaders();
const headers = x[0]
const headerIndices = x[1]
const prompts = x[2]
const apiKey = 'sk-proj-NAyMH2u9z5b5nTaqrOfNT3BlbkFJywgPok9bPiUJ055xGHCk'; // Global variable


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
  sheet.getRange(row, col, response.length, response[0].length).setValues(response);

}



function runEvaluationCell(row){
  var rowDict = getRow(row)
  const type = "Evaluation rating";
  const promptTemplate = sheet.getRange("B2").getValue();
  const prompt = fixPrompt(promptTemplate, rowDict);
  const response = callOpenAI(prompt);
  return [[response, response]]
}


function runSimpleCell(row, type){
  var rowDict = getRow(row)
  const promptTemplate = prompts[type];
  const prompt = fixPrompt(promptTemplate, rowDict);
  const response = callOpenAI(prompt);
  return [[response]]
}


function runSubtaskDecomposition(row){
  var rowDict = getRow(row)
  const type = "Subtask decomposition";
  const promptTemplate = prompts[type];
  const prompt = fixPrompt(promptTemplate, rowDict);
  const response = callOpenAI(prompt);
  var subtasks = response.split(/\d+\./).map(s => s.trim()).filter(Boolean);
  var responseList = subtasks.map((subtask, index) => [response, index+1, subtask]);
  return responseList
}



function runCompletion(minRow, maxRow) {
  for (let row = minRow; row <= maxRow; row++){
    var type = "Results format"
    col = headerIndices[type]
    var response = runSimpleCell(row, type)
    setCell(col, row, response)

    var type = "Subtask completion"
    col = headerIndices[type]
    var response = runSimpleCell(row, type)
    setCell(col, row, response)
    
  }
}

function runEvaluation(minRow, maxRow) {
  minCol = headerIndices["Evaluation rating"]
  maxCol = minCol //evaluation reasoning does not need to be run separately
  for (let row = minRow; row <= maxRow; row++){
    for (let col = minCol; col <= maxCol; col++) {
      const cell = sheet.getRange(row, col);
      col = cell.getColumn()
      row = cell.getRow();
      var response = runEvaluationCell(row)
      setCell(col, row, response)
    }
  }
}




function runContextAndSubtaskDecomposition(row){
  //context 
  var type = "Context"
  col = headerIndices[type]
  var response = runSimpleCell(row, type)
  setCell(col, row, response)


  //subtask decompositon
  col = headerIndices["Subtask decomposition"]
  var response = runSubtaskDecomposition(row)
  setCell(col, row, response)
  maxRow = row+response.length-1
  return maxRow
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
  sheet.getRange(firstRowIndex, col, 1, sheet.getMaxColumns() - col + 1).clearContent();
  return [firstRowIndex, lastRowIndex];
}


function run(){
  var cell = sheet.getActiveRange();
  var row = cell.getRow(); 
  var row = 6
  row = resetTaskID(row)[0]
  const maxCol = getLastColumnForRow(headerRow); 
  var maxRow = runContextAndSubtaskDecomposition(row)
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
  if (type == "Evaluation rating"){
    var response = runEvaluationCell(row)
  }else if(type in ["Context", "Results format", "Subtask completion"]){
    var response = runSimpleCell(row,type)
  }else{
    return
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
