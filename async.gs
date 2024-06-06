//   var properties = PropertiesService.getScriptProperties();
//   if (!properties.getProperty('number')) {
//     properties.setProperty('number', 0);
//   }




// async function runAsyncCompletionCell(row,col, type){
//    let response = runSimpleCell(row, type); 
//    setCell(col, row, response);
// }

// async function runAsyncCompletion(minRow, maxRow) {
//   let rows = Array.from({ length: maxRow - minRow + 1 }, (_, i) => minRow + i);
  
//   let promises = rows.map(async (row) => {
//     console.log(row);

//     let type = "Results format";
//     let col = headerIndices[type];
//     await runAsyncCompletionCell(row, col, type);

//     type = "Subtask completion";
//     col = headerIndices[type];
//     await runAsyncCompletionCell(row, col, type);
//   });

//   await Promise.all(promises);
// }


// async function mockRun(){
//   runAsyncCompletion(6,10)
// }

// async function runAsyncEvaluation(minRow, maxRow) {
//   let minCol = headerIndices["Evaluation Rating"];
//   let maxCol = minCol; // Evaluation reasoning does not need to be run separately
//   for (let row = minRow; row <= maxRow; row++) {
//     for (let col = minCol; col <= maxCol; col++) {
//       const cell = sheet.getRange(row, col);
//       col = cell.getColumn();
//       row = cell.getRow();
//       let response = await runEvaluationCell(row);
//       await setCell(col, row, response);
//     }
//   }
// }
// async function runRows() {
//   let cell = sheet.getActiveRange();
//   let row = cell.getRow();
//   row = 6;
//   await resetTaskID(row);
//   const maxCol = getLastColumnForRow(headerRow);
//   let x = await runContextAndSubtaskDecomposition(row);
//   let minCol = x[0];
//   let maxRow = x[1];
//   let minRow = row;
//   await runAsyncCompletion(minRow, maxRow);
//   await runAsyncEvaluation(minRow, maxRow);
// }
