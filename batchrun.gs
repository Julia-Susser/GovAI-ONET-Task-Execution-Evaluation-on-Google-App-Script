function setBatchCells(col, row, responses) {
  sheet.getRange(row, col, responses.length, responses[0].length).setValues(responses);
}

function batchRunContextAndSubtaskDecomposition(minRow, maxRow){
  let rows = Array.from({ length: maxRow - minRow + 1 }, (_, i) => minRow + i);
  
  // context 
  var type = "Context"
  col = headerIndices[type]
  var batchPrompts = rows.map(row => {
    var rowDict = getRow(row)
    const promptTemplate = prompts[type];
    const prompt = fixPrompt(promptTemplate, rowDict);
    return prompt
  })
  console.log(batchPrompts)
  responses = batchOpenAIRequest(batchPrompts)
  setBatchCells(col, minRow, responses)


  //subtask decompositon
  var type = "Subtask decomposition"
  col = headerIndices[type]
  var batchPrompts = rows.map(row => {
    var rowDict = getRow(row)
    const promptTemplate = prompts[type];
    const prompt = fixPrompt(promptTemplate, rowDict);
    return prompt
  })
  console.log(batchPrompts)
  var responses = batchOpenAIRequest(batchPrompts)
  var subtasksBatch = responses.map(response => response[0].split(/\d+\./).map(s => s.trim()).filter(Boolean));
  var responseList = subtasksBatch.map((subtasks, i) => subtasks.map((subtask, index) => [responses[i][0], index+1, subtask]));
  taskRows = []
  row = minRow
  console.log(responseList)
  for (let i = 0; i < responseList.length; i++) {
    const response = responseList[i];
    const rowCount = response.length;
    
    setCell(col, row, response);
    taskRows.push([row, row + rowCount - 1]);
    row += rowCount;
  }
  return taskRows
}


function batchRunCompletion(minRow, maxRow){
  let rows = Array.from({ length: maxRow - minRow + 1 }, (_, i) => minRow + i);
  
  // results format
  var type = "Results format"
  col = headerIndices[type]
  var batchPrompts = rows.map(row => {
    var rowDict = getRow(row)
    const promptTemplate = prompts[type];
    const prompt = fixPrompt(promptTemplate, rowDict);
    return prompt
  })
  console.log(batchPrompts)
  responses = batchOpenAIRequest(batchPrompts)
  setBatchCells(col, minRow, responses)


  // subtask completion
  var type = "Subtask completion"
  col = headerIndices[type]
  var batchPrompts = rows.map(row => {
    var rowDict = getRow(row)
    const promptTemplate = prompts[type];
    const prompt = fixPrompt(promptTemplate, rowDict);
    return prompt
  })
  console.log(batchPrompts)
  responses = batchOpenAIRequest(batchPrompts)
  setBatchCells(col, minRow, responses)
}


function batchRunEvaluation(minRow, maxRow){
  let rows = Array.from({ length: maxRow - minRow + 1 }, (_, i) => minRow + i);
  
  // evaluation rating
  var type = "Evaluation rating"
  col = headerIndices[type]
  var batchPrompts = rows.map(row => {
    var rowDict = getRow(row)
    const promptTemplate = sheet.getRange("B2").getValue();
    const prompt = fixPrompt(promptTemplate, rowDict);
    return prompt
  })
  responses = batchOpenAIRequest(batchPrompts)
  response = responses.map(response => [response, response])
  setBatchCells(col, minRow, response)

}

function resetRows(startRow,endRow){
  for (let row = startRow; row <= endRow; row++){
    var reset = resetTaskID(row)
    if (reset == undefined){
      endRow = row -1
    }
    if (reset[0] < startRow){
      startRow = reset[0]
    }
    row = reset[0]
    var change = reset[1]-reset[0]
    if (endRow - change > startRow){
          endRow = endRow-change
    }else{
      endRow = row
    }
    console.log(row,startRow, endRow)
  }
  return [startRow,endRow]
}

function runBatch(){
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var endRow = range.getLastRow();
  
  var startRow = 6
  var n = 10
  var endRow = startRow + n -1
  
  var reset = resetRows(startRow,endRow)
  startRow = reset[0]
  endRow = reset[1]

  var taskRows = batchRunContextAndSubtaskDecomposition(startRow, endRow)
  for (let i=0; i< taskRows.length; i++){
    minRow = taskRows[i][0]
    maxRow = taskRows[i][1]
    batchRunCompletion(minRow, maxRow)
    batchRunEvaluation(minRow, maxRow)
  }
}
