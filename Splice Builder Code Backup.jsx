function executeModules() {
  manageSheets();
  copyDataToOutputSheet();
  sortOutputTable();
  createCalculationsSheet();
  populateCalculationsSheet();
  populateLowHighValues();
  populateColumnsWithNumbers();
  findAndUpdateHighLowValues();
  determineStubs();
  populateColumnsWithStubData();
  determineCableCorrelations();
  //colorNumericColumns();
  calculateColumnMinMax();
  populateColumnsWithinRanges();
  replaceColumnIWithColumnM();
  populateColumnsKO();
  summarizeValuesLow();
  summarizeValuesHigh();
  combineData();
  removeKOValues();
  removeCellsStartingWithUnderscore();
  removeCellsWithLessThanThreeUnderscores();
  splitColumnCByUnderscore();
  populateColumnB();
  fillDownColumnB();
  populateColumnBWithData();
  removeEmptyRowsAndSortByF();
  insertHeadersAboveData();
  insertColumnBetweenBAndC();
  populateColumnK();
  populateColumnLWithMax();
  populateColumnMFromRawSheet();
  populateColumnNFromRawSheet();
  findJumperBase();

  var duplicatesFound = checkForDuplicates(); // Simulated function to check for duplicates

  if (duplicatesFound) {
    // Follow this path if duplicates are found
    updateJumperBase();
    createNewRowWithJumperBaseData();
    updateJumperNew();
    updateJumperNewOut();
    createTempSheetAndCopyData();
    recalculateColumnNForJumperBase();
    updateJumperNewFromJumperBase();
    recalculateColumnNForJumperTarget();
    rearrangeColumnsInOutputSheet();
    replaceHeaders();
    combineColumnsEFAndSetHeader();
    removeColumnF();
    customSortOutputSheet();
    calculateHighLowDifference();
    calculateHighLowDifferenceJump();
    updateColumnKBasedOnChangeRow();
    updateColumnKFinal();
    deduplicateColumnsJK();
    updateColumnJAndDivideK()
    removeTempSheetAndFormatOutput();
    donePopup();
  } else {
    // Follow this path if no duplicates are found
    createTempSheetAndCopyData();
    rearrangeColumnsInOutputSheet();
    replaceHeaders();
    combineColumnsEFAndSetHeader();
    removeColumnF();
    customSortOutputSheet();
    calculateHighLowDifference();
    updateColumnKFinal();
    deduplicateColumnsJK();
    updateColumnJAndDivideK()
    removeTempSheetAndFormatOutput();
    donePopup();
  }
}

function checkForDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  var lastRow = outputSheet.getLastRow();
  var columnAValues = outputSheet.getRange('A2:A' + lastRow).getValues();

  // Check if there are any non-empty cells in Column A
  for (var i = 0; i < columnAValues.length; i++) {
    if (columnAValues[i][0] !== '') {
      return true; // Duplicates found (non-empty cell)
    }
  }
  return false; // No duplicates found (all cells in Column A are empty)
}


// Function to display a popup indicating the completion of a module
function donePopup() {
  SpreadsheetApp.getUi().alert("Congrats, your splices are done!");
}


function manageSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check for the 'calculations' sheet
  var calculationsSheet = ss.getSheetByName('calculations');
  if (calculationsSheet) {
    ss.deleteSheet(calculationsSheet);
  }

  // Check for the 'TempSheet' sheet
  var TempSheetSheet = ss.getSheetByName('TempSheet');
  if (TempSheetSheet) {
    ss.deleteSheet(TempSheetSheet);
  }

  // Check for the 'Output' sheet
  var outputSheet = ss.getSheetByName('Output');
  if (outputSheet) {
    outputSheet.clear(); // Clear all sheet contents
    outputSheet.clearFormats(); // Remove all formatting
  }
}

function copyDataToOutputSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = ss.getSheetByName('Raw');
  var outputSheet = ss.getSheetByName('Output');
  
  // Get the range containing the headers
  var headersRange = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  
  // Find the indices of the columns with headers "Cable_ID" and "PON_Count"
  var cableIDIndex = headersRange.indexOf("Cable_ID") + 1; // Adding 1 to match 1-indexed columns in Sheets
  var ponCountIndex = headersRange.indexOf("PON_Count") + 1; // Adding 1 to match 1-indexed columns in Sheets
  
  // Get the data from the identified columns
  var lastRow = rawSheet.getLastRow();
  var cableIDData = rawSheet.getRange(2, cableIDIndex, lastRow - 1, 1).getValues(); // Excluding header row
  var ponCountData = rawSheet.getRange(2, ponCountIndex, lastRow - 1, 1).getValues(); // Excluding header row
  
  // Combine the data into a single array for easy pasting
  var outputData = [];
  for (var i = 0; i < cableIDData.length; i++) {
    outputData.push([cableIDData[i][0], ponCountData[i][0]]);
  }
  
  // Paste the data into columns A and B of the Output sheet
  outputSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
}

// Function to sort the Output table based on the lowest value in column B
function sortOutputTable() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  
  var lastRow = outputSheet.getLastRow();
  var outputData = outputSheet.getRange('A1:B' + lastRow).getValues();

  outputData.sort(function(a, b) {
    var numA = extractLowestNumber(a[1]);
    var numB = extractLowestNumber(b[1]);
    return numA - numB;
  });

  outputSheet.getRange('A1:B' + lastRow).setValues(outputData);
}

// Helper function to extract the lowest number from "TEXT:Number1-Number2"
function extractLowestNumber(text) {
  var numbers = text.match(/\d+/g);
  return Math.min.apply(null, numbers);
}

function createCalculationsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  
  var calculationsSheet = ss.getSheetByName('calculations');
  
  // Remove existing calculations sheet if it exists
  if (calculationsSheet !== null) {
    ss.deleteSheet(calculationsSheet);
  }
  
  // Create a new sheet named "calculations"
  calculationsSheet = ss.insertSheet('calculations');
  
  var headers = outputSheet.getRange('A1:A').getValues().filter(String).flat();
  
  // Set the headers in the calculations sheet
  for (var i = 0; i < headers.length; i++) {
    calculationsSheet.getRange(1, i + 1).setValue(headers[i]);
  }
}

function populateCalculationsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var calculationsSheet = ss.getSheetByName('calculations');

  var outputLastRow = outputSheet.getLastRow();
  var outputData = outputSheet.getRange('A1:B' + outputLastRow).getValues();
  
  var calculationsData = [];
  var columnA = [];
  var columnB = [];
  
  // Extract Column A and Column B data, removing text before the colon
  for (var i = 0; i < outputData.length; i++) {
    columnA.push(removeTextBeforeColon(outputData[i][0]));
    columnB.push(removeTextBeforeColon(outputData[i][1]));
  }

  calculationsData.push(columnA); // Push modified Column A data
  calculationsData.push(columnB); // Push modified Column B data
  
  // Set values in the calculations sheet starting from cell A1
  calculationsSheet.getRange(1, 1, calculationsData.length, calculationsData[0].length).setValues(calculationsData);
}

// Function to remove text before the colon (including the colon)
function removeTextBeforeColon(text) {
  if (typeof text === 'string') {
    var colonIndex = text.indexOf(':');
    if (colonIndex !== -1) {
      return text.substring(colonIndex + 1).trim();
    }
    return text;
  }
  return text;
}

function populateLowHighValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');

  var dataRange = calculationsSheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values[1].length; i++) {
    if (values[1][i] !== '' && values[1][i] !== null) {
      var splitValues = values[1][i].toString().split('-');
      if (splitValues.length === 2) {
        calculationsSheet.getRange(2, i + 1).setValue(splitValues[0].trim()); // Set low value in row 2
        calculationsSheet.getRange(3, i + 1).setValue(splitValues[1].trim()); // Set high value in row 3
      }
    }
  }
}

function populateColumnsWithNumbers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');

  var lastColumn = calculationsSheet.getLastColumn();

  for (var i = 1; i <= lastColumn; i++) {
    var lowValue = calculationsSheet.getRange(2, i).getValue(); // Get the low value in row 2 for each column
    var highValue = calculationsSheet.getRange(3, i).getValue(); // Get the high value in row 3 for each column

    var range = highValue - lowValue + 1; // Calculate the range of numbers to fill

    var columnData = [];
    
    // Populate the column with numbers from lowValue to highValue
    for (var j = 0; j < range; j++) {
      columnData.push([lowValue + j]);
    }

    var columnRange = calculationsSheet.getRange(2, i, columnData.length, 1); // Define the range for each column
    columnRange.setValues(columnData); // Set the values in the range
  }
}

function findAndUpdateHighLowValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');

  var lastColumn = calculationsSheet.getLastColumn();

  var lowValues = [];
  var highValues = [];

  // Loop through each column and find low and high values
  for (var i = 1; i <= lastColumn; i++) {
    var columnData = calculationsSheet.getRange(2, i, calculationsSheet.getLastRow() - 1, 1).getValues(); // Get column data excluding header

    var values = columnData.map(function (row) {
      return row[0];
    });

    values = values.filter(Boolean); // Filter out empty values

    if (values.length > 0) {
      lowValues.push(Math.min.apply(null, values));
      highValues.push(Math.max.apply(null, values));
    }
  }

  // Find overall low and high values
  var overallLowValue = Math.min.apply(null, lowValues);
  var overallHighValue = Math.max.apply(null, highValues);

  // Update cell F2 with the low-high value string
  calculationsSheet.getRange('F2').setValue(`(${overallLowValue}-${overallHighValue})`);

  // Calculate and update cell F3 with the difference divided by 12
  var difference = overallHighValue - (overallLowValue - 1); // Subtract (overallLowValue - 1) from overallHighValue
  var divisionResult = difference / 12;
  calculationsSheet.getRange('F3').setValue(divisionResult);
}

function determineStubs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');

  var highRange = calculationsSheet.getRange('F2').getValue(); // Retrieve the range from cell F2
  var regex = /\((\d+)-(\d+)\)/; // Regular expression to match values in parentheses separated by hyphen
  var match = highRange.match(regex); // Match the range format

  var highValue = 0;

  if (match) {
    var value1 = parseInt(match[1]);
    var value2 = parseInt(match[2]);
    highValue = Math.max(value1, value2); // Determine the higher value from the range
  }

  var stub1 = '';
  var stub2 = '';

  if (highValue === 864) {
    stub1 = '433-864';
    stub2 = '1-432';
  } else if (highValue === 576) {
    stub1 = '289-576';
    stub2 = '1-288';
  } else if (highValue === 432) {
    stub1 = '1-432';
    stub2 = 'N/A';
  } else {
    stub1 = 'Undefined';
    stub2 = 'Undefined';
  }

  // Update column G with the heading and stub information
  calculationsSheet.getRange('G1').setValue('Stubs');
  calculationsSheet.getRange('G2').setValue('Stub 1 = ' + stub1);
  calculationsSheet.getRange('G3').setValue('Stub 2 = ' + stub2);
}

function populateColumnsWithStubData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');

  var stub1 = calculationsSheet.getRange('G2').getValue(); // Get Stub 1 data
  var stub2 = calculationsSheet.getRange('G3').getValue(); // Get Stub 2 data

  populateColumnWithStubData(stub1, 'H', 'Stub 1');
  populateColumnWithStubData(stub2, 'L', 'Stub 2');
}

function populateColumnWithStubData(stubData, column, columnHeader) {
  if (stubData) {
    var parts = stubData.split('=');
    if (parts.length === 2) {
      var range = parts[1].trim().split('-').map(Number);

      var lower = Math.min(...range);
      var upper = Math.max(...range);

      var columnValues = [];
      for (var i = upper; i >= lower; i--) {
        columnValues.push([i]);
      }

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var calculationsSheet = ss.getSheetByName('calculations');
      var columnIndex = getColumnIndex(column);

      calculationsSheet.getRange(1, columnIndex).setValue(columnHeader); // Set the header in H1 or L1
      calculationsSheet.getRange(2, columnIndex, columnValues.length, 1).setValues(columnValues); // Set the values starting from row 2
    }
  }
}

function getColumnIndex(column) {
  return column.charCodeAt(0) - 64; // Converts column letter to its numerical index
}

function determineCableCorrelations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  
  var lastRow = calculationsSheet.getLastRow();
  var valuesH = calculationsSheet.getRange('H2:H' + lastRow).getValues();
  var valuesL = calculationsSheet.getRange('L2:L' + lastRow).getValues();

  calculationsSheet.getRange('I2:I' + lastRow).setValues(valuesH);
  calculationsSheet.getRange('M2:M' + lastRow).setValues(valuesL);
}

function replaceValuesWithHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  
  var headers = calculationsSheet.getRange(1, 1, 1, 5).getValues()[0];

  for (var i = 2; i <= calculationsSheet.getLastRow(); i++) {
    var valueI = calculationsSheet.getRange(i, 9).getValue();
    var valueM = calculationsSheet.getRange(i, 13).getValue();

    for (var j = 1; j <= 5; j++) {
      var cellValue = calculationsSheet.getRange(i, j).getValue();
      if (cellValue === valueI) {
        calculationsSheet.getRange(i, 9).setValue(headers[j - 1]);
      }
      if (cellValue === valueM) {
        calculationsSheet.getRange(i, 13).setValue(headers[j - 1]);
      }
    }
  }
}

function colorNumericColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  
  var lastColumn = calculationsSheet.getLastColumn();
  var headers = calculationsSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  
  var colors = ["#FFA07A", "#90EE90", "#87CEEB", "#FA8072", "#ADD8E6", "#F08080"]; // Array of different colors
  colors = shuffleArray(colors); // Shuffle the colors array
  
  var usedColors = {}; // Object to keep track of used colors
  
  for (var col = 1; col <= lastColumn; col++) {
    var header = headers[col - 1];
    var isNumericHeader = /^[0-9.]+$/.test(header);
    
    if (isNumericHeader) {
      var range = calculationsSheet.getRange(1, col, calculationsSheet.getLastRow(), 1);
      
      var color = colors.pop(); // Get the last color from the shuffled array
      
      if (!usedColors[color]) {
        range.setBackground(color);
        usedColors[color] = true; // Mark the color as used
      } else {
        colors.unshift(color); // Put the color back in the array's beginning if it's already used
      }
    }
  }
}

// Function to shuffle the colors array
function shuffleArray(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function calculateColumnMinMax() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  
  var lastColumn = calculationsSheet.getLastColumn();
  var headers = calculationsSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var dataRange = calculationsSheet.getRange(2, 1, calculationsSheet.getLastRow() - 1, lastColumn);
  var dataValues = dataRange.getValues();
  
  var output = [];
  
  for (var col = 0; col < lastColumn; col++) {
    var header = headers[col];
    var isNumericHeader = /^[0-9.]+$/.test(header);
    
    if (isNumericHeader) {
      var columnValues = dataValues.map(function(row) {
        return row[col];
      }).filter(function(value) {
        return value !== ''; // Filter out blank values
      });
      
      if (columnValues.length > 0) {
        var min = Math.min(...columnValues);
        var max = Math.max(...columnValues);
        
        output.push([header, min + "-" + max]);
      } else {
        output.push([header, '']);
      }
    }
  }
  
  calculationsSheet.getRange(5, 6, output.length, 2).setValues(output);
}

function populateColumnsWithinRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  
  var rangesColumnG = calculationsSheet.getRange('G2:G').getValues().filter(String);
  var valuesColumnF = calculationsSheet.getRange('F2:F').getValues().filter(String);
  
  var valuesColumnI = calculationsSheet.getRange('I2:I').getValues().flat();
  var valuesColumnM = calculationsSheet.getRange('M2:M').getValues().flat();
  
  for (var i = 0; i < valuesColumnI.length; i++) {
    for (var j = 0; j < rangesColumnG.length; j++) {
      var range = rangesColumnG[j][0].split("-");
      var min = parseFloat(range[0]);
      var max = parseFloat(range[1]);
      
      if (valuesColumnI[i] >= min && valuesColumnI[i] <= max) {
        calculationsSheet.getRange(i + 2, 10).setValue(valuesColumnF[j][0]);
        break;
      }
    }
  }
  
  for (var i = 0; i < valuesColumnM.length; i++) {
    for (var j = 0; j < rangesColumnG.length; j++) {
      var range = rangesColumnG[j][0].split("-");
      var min = parseFloat(range[0]);
      var max = parseFloat(range[1]);
      
      if (valuesColumnM[i] >= min && valuesColumnM[i] <= max) {
        calculationsSheet.getRange(i + 2, 14).setValue(valuesColumnF[j][0]);
        break;
      }
    }
  }
}

function replaceColumnIWithColumnM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  
  var valuesColumnM = calculationsSheet.getRange('M2:M').getValues().flat();
  
  calculationsSheet.getRange('I2:I').setValues(valuesColumnM.map(value => [value]));
}

function populateColumnsKO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  
  var lastRow = calculationsSheet.getLastRow();
  
  for (var i = 2; i <= lastRow; i++) {
    var valueK = calculationsSheet.getRange(i, 10).getValue() + '/' + 
                 calculationsSheet.getRange(i, 8).getValue() + '/' + 
                 calculationsSheet.getRange(i, 9).getValue();
    
    var valueO = calculationsSheet.getRange(i, 14).getValue() + '/' + 
                 calculationsSheet.getRange(i, 12).getValue() + '/' + 
                 calculationsSheet.getRange(i, 13).getValue();
    
    calculationsSheet.getRange(i, 11).setValue(valueK);
    calculationsSheet.getRange(i, 15).setValue(valueO);
  }
}

function summarizeValuesLow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  var outputSheet = ss.getSheetByName('Output'); // Replace 'Output' with your output sheet name
  
  var lastRow = calculationsSheet.getLastRow();
  
  var valuesK = calculationsSheet.getRange('K2:K' + lastRow).getValues();
  var valuesO = calculationsSheet.getRange('O2:O' + lastRow).getValues();
  var valuesI = calculationsSheet.getRange('I2:I' + lastRow).getValues();
  var valuesM = calculationsSheet.getRange('M2:M' + lastRow).getValues();
  
  var summariesK = {};
  var summariesO = {};
  
  // Process column K
  for (var i = 0; i < valuesK.length; i++) {
    var valueK = valuesK[i][0];
    
    if (!valueK) continue;
    
    var partsK = valueK.split('/');
    var keyK = partsK[0] + '_K';
    var stubK = getStubFromColumn(valueK, calculationsSheet, 'H');
    
    if (!summariesK[keyK]) {
      summariesK[keyK] = { high: -Infinity, low: Infinity, stub: stubK, value3: '' };
    }
    
    var secondValueK = parseInt(partsK[1]);
    
    if (secondValueK > summariesK[keyK].high) {
      summariesK[keyK].high = secondValueK;
    }
    
    if (secondValueK < summariesK[keyK].low) {
      summariesK[keyK].low = secondValueK;
    }
    
    if (valuesI[i][0] !== '') {
      summariesK[keyK].value3 = valuesI[i][0];
    }
  }
  
  // Process column O
  for (var j = 0; j < valuesO.length; j++) {
    var valueO = valuesO[j][0];
    
    if (!valueO) continue;
    
    var partsO = valueO.split('/');
    var keyO = partsO[0] + '_O';
    var stubO = getStubFromColumn(valueO, calculationsSheet, 'L');
    
    if (!summariesO[keyO]) {
      summariesO[keyO] = { high: -Infinity, low: Infinity, stub: stubO, value3: '' };
    }
    
    var secondValueO = parseInt(partsO[1]);
    
    if (secondValueO > summariesO[keyO].high) {
      summariesO[keyO].high = secondValueO;
    }
    
    if (secondValueO < summariesO[keyO].low) {
      summariesO[keyO].low = secondValueO;
    }
    
    if (valuesM[j][0] !== '') {
      summariesO[keyO].value3 = valuesM[j][0];
    }
  }
  
  var outputDataK = [];
  var outputDataO = [];
  
  // Prepare output data for column K
  for (var keyK in summariesK) {
    var highK = summariesK[keyK].high !== -Infinity ? keyK + '_' + summariesK[keyK].high + summariesK[keyK].stub + '_' + summariesK[keyK].value3 : '';
    var lowK = summariesK[keyK].low !== Infinity ? keyK + '_' + summariesK[keyK].low + summariesK[keyK].stub + '_' + summariesK[keyK].value3 : '';
    
    outputDataK.push([highK, lowK]);
  }
  
  // Prepare output data for column O
  for (var keyO in summariesO) {
    var highO = summariesO[keyO].high !== -Infinity ? keyO + '_' + summariesO[keyO].high + summariesO[keyO].stub + '_' + summariesO[keyO].value3 : '';
    var lowO = summariesO[keyO].low !== Infinity ? keyO + '_' + summariesO[keyO].low + summariesO[keyO].stub + '_' + summariesO[keyO].value3 : '';
    
    outputDataO.push([highO, lowO]);
  }
  
  outputSheet.getRange(2, 3, outputDataK.length, 2).setValues(outputDataK);
  outputSheet.getRange(2, 7, outputDataO.length, 2).setValues(outputDataO);
}

function getStubFromColumn(value, sheet, columnLetter) {
  var headerValues = sheet.getRange(columnLetter + '1:1').getValues()[0];
  var valueIndex = headerValues.indexOf(value);
  if (valueIndex !== -1) {
    var stub = headerValues[valueIndex - 1];
    return stub !== '' ? '_' + stub : '';
  }
  return '';
}

function summarizeValuesHigh() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationsSheet = ss.getSheetByName('calculations');
  var outputSheet = ss.getSheetByName('Output'); // Replace 'Output' with your output sheet name
  
  var lastRow = calculationsSheet.getLastRow();
  
  var valuesK = calculationsSheet.getRange('K2:K' + lastRow).getValues();
  var valuesO = calculationsSheet.getRange('O2:O' + lastRow).getValues();
  var valuesI = calculationsSheet.getRange('I2:I' + lastRow).getValues();
  var valuesM = calculationsSheet.getRange('M2:M' + lastRow).getValues();
  
  var summariesK = {};
  var summariesO = {};
  
  // Process column K
  for (var i = 0; i < valuesK.length; i++) {
    var valueK = valuesK[i][0];
    
    if (!valueK) continue;
    
    var partsK = valueK.split('/');
    var keyK = partsK[0] + '_K';
    var stubK = getStubFromColumn(valueK, calculationsSheet, 'H');
    
    if (!summariesK[keyK]) {
      summariesK[keyK] = { high: -Infinity, stub: stubK, value3: '' };
    }
    
    var secondValueK = parseInt(partsK[1]);
    
    if (secondValueK > summariesK[keyK].high) {
      summariesK[keyK].high = secondValueK;
      summariesK[keyK].value3 = valuesI[i][0];
    }
  }
  
  // Process column O
  for (var j = 0; j < valuesO.length; j++) {
    var valueO = valuesO[j][0];
    
    if (!valueO) continue;
    
    var partsO = valueO.split('/');
    var keyO = partsO[0] + '_O';
    var stubO = getStubFromColumn(valueO, calculationsSheet, 'L');
    
    if (!summariesO[keyO]) {
      summariesO[keyO] = { high: -Infinity, stub: stubO, value3: '' };
    }
    
    var secondValueO = parseInt(partsO[1]);
    
    if (secondValueO > summariesO[keyO].high) {
      summariesO[keyO].high = secondValueO;
      summariesO[keyO].value3 = valuesM[j][0];
    }
  }
  
  var outputDataK = [];
  var outputDataO = [];
  
  // Prepare output data for column K
  for (var keyK in summariesK) {
    var highK = summariesK[keyK].high !== -Infinity ? keyK + '_' + summariesK[keyK].high + summariesK[keyK].stub + '_' + summariesK[keyK].value3 : '';
    outputDataK.push([highK]);
  }
  
  // Prepare output data for column O
  for (var keyO in summariesO) {
    var highO = summariesO[keyO].high !== -Infinity ? keyO + '_' + summariesO[keyO].high + summariesO[keyO].stub + '_' + summariesO[keyO].value3 : '';
    outputDataO.push([highO]);
  }
  
  outputSheet.getRange(2, 3, outputDataK.length, 1).setValues(outputDataK);
  outputSheet.getRange(2, 7, outputDataO.length, 1).setValues(outputDataO);
}

function getStubFromColumn(value, sheet, columnLetter) {
  var headerValues = sheet.getRange(columnLetter + '1:1').getValues()[0];
  var valueIndex = headerValues.indexOf(value);
  if (valueIndex !== -1) {
    var stub = headerValues[valueIndex - 1];
    return stub !== '' ? '_' + stub : '';
  }
  return '';
}

function combineData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name

  var lastRow = outputSheet.getLastRow();
  var dataRange = outputSheet.getRange('C2:ZZ' + lastRow); // Assuming the data is in columns C to ZZ

  var values = dataRange.getValues();
  var combinedData = {};

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j]) {
        var parts = values[i][j].toString().split('_');
        var key = parts[0];
        var letter = /[A-Za-z]/.exec(values[i][j].toString());

        if (!combinedData[key]) {
          combinedData[key] = {};
        }

        if (letter) {
          letter = letter[0].toUpperCase();
          if (!combinedData[key][letter]) {
            combinedData[key][letter] = [];
          }
          combinedData[key][letter].push(values[i][j].toString());
        }
      }
    }
  }

  var outputArray = [];
  for (var key in combinedData) {
    for (var letter in combinedData[key]) {
      outputArray.push([key + '_' + combinedData[key][letter].join('_')]);
    }
  }

  if (outputArray.length > 0) {
    outputSheet.getRange('C' + (lastRow + 1)).offset(0, 0, outputArray.length, 1).setValues(outputArray);
  }
}



function removeKOValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name

  var lastRow = outputSheet.getLastRow();
  var dataRange = outputSheet.getRange('C2:ZZ' + lastRow); // Assuming the data is in columns C to ZZ

  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (typeof values[i][j] === 'string') {
        values[i][j] = values[i][j].replace(/K_|O_/g, '');
      }
    }
  }

  dataRange.setValues(values);
}

function removeCellsStartingWithUnderscore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Replace 'Output' with your sheet name
  var lastRow = outputSheet.getLastRow();
  var dataRange = outputSheet.getRange('C2:C' + lastRow); // Assuming the data is in column C

  var values = dataRange.getValues();
  var rowsToDelete = [];

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().startsWith('_')) {
      rowsToDelete.push(i + 2); // Adding 2 to get the actual row number in the sheet
    }
  }

  if (rowsToDelete.length > 0) {
    outputSheet.deleteRows(rowsToDelete);
  }
}

function removeCellsWithLessThanThreeUnderscores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Replace 'Output' with your sheet name
  var lastRow = outputSheet.getLastRow();
  var dataRange = outputSheet.getRange('C2:J' + lastRow); // Target columns C and D

  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] && countUnderscores(values[i][j].toString()) < 3) {
        dataRange.getCell(i + 1, j + 1).setValue(''); // Clear the cell content
      }
    }
  }
}

function countUnderscores(str) {
  return str.split('_').length - 1;
}

function splitColumnCByUnderscore() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Replace 'Output' with your sheet name
  var lastRow = outputSheet.getLastRow();
  var dataRange = outputSheet.getRange('C2:C' + lastRow); // Target column C

  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) {
      var splitData = values[i][0].toString().split('_');
      outputSheet.getRange(i + 2, 3, 1, splitData.length).setValues([splitData]); // Set split data back in column C
    }
  }
}

function populateColumnB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name
  var lastRow = outputSheet.getLastRow();

  for (var i = 1; i <= lastRow; i++) {
    var cellB = outputSheet.getRange('B' + i);
    var cellValue = cellB.getValue();
    
    if (cellValue && cellValue.includes(':')) {
      var extractedValue = cellValue.split(':')[0].trim();
      cellB.setValue(extractedValue);
    }
  }
}

function fillDownColumnB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name
  var lastRow = outputSheet.getLastRow();
  var valueToFill = outputSheet.getRange('B1').getValue();
  
  var rangeToFill = outputSheet.getRange('B1:B' + lastRow);
  var values = [];
  
  for (var i = 0; i < lastRow; i++) {
    values.push([valueToFill]);
  }
  
  rangeToFill.setValues(values);
}

function populateColumnBWithData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name
  var lastRow = outputSheet.getLastRow();

  var range = outputSheet.getRange('B1:B' + lastRow);
  var values = range.getValues();

  for (var i = 0; i < values.length; i++) {
    var currentBValue = values[i][0];
    var jValue = outputSheet.getRange('H' + (i + 1)).getValue();
    var fValue = outputSheet.getRange('E' + (i + 1)).getValue();

    values[i][0] = currentBValue + ':' + jValue + '-' + fValue;
  }

  range.setValues(values);
}

function removeEmptyRowsAndSortByF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name

  var lastRow = outputSheet.getLastRow();
  var range = outputSheet.getRange('C1:C' + lastRow);

  var values = range.getValues();
  var rowsToRemove = [];

  for (var i = values.length - 1; i >= 0; i--) {
    if (!values[i][0]) {
      rowsToRemove.push(i + 1);
    }
  }

  if (rowsToRemove.length > 0) {
    rowsToRemove.forEach(function(row) {
      outputSheet.deleteRow(row);
    });
  }

  outputSheet.getRange('A:Z').sort([{column: 5, ascending: false}]); // Sort by Column F in descending order (assuming Column F is column 6)
}

function insertHeadersAboveData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name

  outputSheet.insertRowsBefore(1, 1); // Insert a row above the existing data

  var headers = [
    ['Active Count', 'Remove', 'Remove', 'PON High', 'Fiber High', 'Cable Out', 'PON Low', 'Fiber Low', 'Cable In', 'Cable In Size', 'Cable Out Fiber']
  ];

  outputSheet.getRange('B1:L1').setValues(headers); // Set the headers in columns B to M
}

function insertColumnBetweenBAndC() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  
  outputSheet.insertColumnBefore(3); // Insert a column before column C
  
  // Set the header for the newly inserted column
  outputSheet.getRange(1, 3).setValue('Formatting Placeholder');
}

function populateColumnK() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name

  var lastRowOutput = outputSheet.getLastRow();
  var fiberHighValues = outputSheet.getRange('F2:F' + lastRowOutput).getValues();
  var fiberLowValues = outputSheet.getRange('G2:G' + lastRowOutput).getValues();
  var stubValues = [];

  for (var i = 0; i < lastRowOutput - 1; i++) {
    var stub = (fiberHighValues[i][0] !== fiberLowValues[i][0]) ? "Stub 1" : "Stub 2";
    stubValues.push([stub]);
  }

  outputSheet.getRange('K2:K' + (lastRowOutput)).setValues(stubValues);
}



function populateColumnLWithMax() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name

  var lastRow = outputSheet.getLastRow();
  var maxNumber = Math.max.apply(null, outputSheet.getRange('G2:G' + lastRow).getValues().flat());

  var columnLValues = Array(lastRow - 1).fill([maxNumber]);
  outputSheet.getRange('L2:L' + lastRow).setValues(columnLValues);
}

function populateColumnMFromRawSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var rawSheet = ss.getSheetByName('Raw');

  var outputLastRow = outputSheet.getLastRow();
  var cableInValues = outputSheet.getRange('H2:H' + outputLastRow).getValues().flat();

  var rawLastRow = rawSheet.getLastRow();
  var rawHeaderRow = rawSheet.getRange('1:1').getValues()[0];
  var cableIdIndex = rawHeaderRow.findIndex(header => header === 'Cable_ID') + 1;
  var totalFiberIndex = rawHeaderRow.findIndex(header => header === 'Total_Fibers') + 1;

  if (cableIdIndex === 0 || totalFiberIndex === 0) {
    Logger.log("Headers 'Cable_ID' or 'Total_Fiber' not found in the Raw sheet.");
    return;
  }

  var cableIdValues = rawSheet.getRange(2, cableIdIndex, rawLastRow - 1).getValues().flat();
  var totalFiberValues = rawSheet.getRange(2, totalFiberIndex, rawLastRow - 1).getValues().flat();

  var fiberData = cableInValues.map(function(cableIn) {
    var index = cableIdValues.indexOf(cableIn);
    if (index !== -1) {
      return [totalFiberValues[index]];
    } else {
      return [''];
    }
  });

  outputSheet.getRange('M2:M' + (outputLastRow)).setValues(fiberData);
}

function populateColumnNFromRawSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var rawSheet = ss.getSheetByName('Raw');

  var outputLastRow = outputSheet.getLastRow();
  var cableInValues = outputSheet.getRange('H2:H' + outputLastRow).getValues().flat();

  var rawLastRow = rawSheet.getLastRow();
  var rawHeaderRow = rawSheet.getRange('1:1').getValues()[0];
  var cableIdIndex = rawHeaderRow.findIndex(header => header === 'Cable_ID') + 1;
  var activeStartsIndex = rawHeaderRow.findIndex(header => header === 'Active_Starts') + 1;
  var activeEndsIndex = rawHeaderRow.findIndex(header => header === 'Active_Ends') + 1;

  if (cableIdIndex === 0 || activeStartsIndex === 0 || activeEndsIndex === 0) {
    Logger.log("Headers 'Cable_ID', 'Active_Starts', or 'Active_Ends' not found in the Raw sheet.");
    return;
  }

  var cableIdValues = rawSheet.getRange(2, cableIdIndex, rawLastRow - 1).getValues().flat();
  var activeStartsValues = rawSheet.getRange(2, activeStartsIndex, rawLastRow - 1).getValues().flat();
  var activeEndsValues = rawSheet.getRange(2, activeEndsIndex, rawLastRow - 1).getValues().flat();

  var fiberData = cableInValues.map(function(cableIn) {
    var index = cableIdValues.indexOf(cableIn);
    if (index !== -1) {
      return [activeStartsValues[index] + ' - ' + activeEndsValues[index]];
    } else {
      return [''];
    }
  });

  outputSheet.getRange('N2:N' + (outputLastRow)).setValues(fiberData);
  
  outputSheet.getRange(1, 14).setValue('Cable Out Fibers'); // Set header after populating data
}

function findJumperBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  var lastRow = outputSheet.getLastRow();
  var columnHValues = outputSheet.getRange('H2:H' + lastRow).getValues().flat();
  
  var duplicates = getDuplicates(columnHValues);
  
  if (duplicates.length === 0) {
    // No duplicates found - Continue the module execution without interruptions
    // You can place any additional logic or code to handle the case when no duplicates are found
  } else {
    // Duplicates found - Process the logic for duplicates
    var minDifference = Number.MAX_VALUE;
    var jumperBaseRow = -1;
    var jumperTargetRow = -1;

    for (var i = 0; i < duplicates.length; i++) {
      var currentHValue = duplicates[i];
      var rowIndices = getAllIndices(columnHValues, currentHValue);

      if (rowIndices.length === 2) {
        var diff1 = Math.abs(outputSheet.getRange('F' + (rowIndices[0] + 1)).getValue() - outputSheet.getRange('I' + (rowIndices[0] + 1)).getValue());
        var diff2 = Math.abs(outputSheet.getRange('F' + (rowIndices[1] + 1)).getValue() - outputSheet.getRange('I' + (rowIndices[1] + 1)).getValue());
        
        var smallerDifference = Math.min(diff1, diff2);
        
        if (smallerDifference < minDifference) {
          minDifference = smallerDifference;
          jumperBaseRow = smallerDifference === diff1 ? rowIndices[0] + 2 : rowIndices[1] + 2; // Selecting the next row
          jumperTargetRow = smallerDifference === diff1 ? rowIndices[1] + 2 : rowIndices[0] + 2; // Storing the other row
        }
      }
    }

    if (jumperBaseRow !== -1) {
      var jumperBaseRange = outputSheet.getRange(jumperBaseRow, 1, 1, outputSheet.getLastColumn());
      jumperBaseRange.setBackground('yellow'); // Highlight the row with yellow color
      outputSheet.getRange('A1').setValue('JumperBase: Row ' + jumperBaseRow); // Store the JumperBase in cell A1
    }

    if (jumperTargetRow !== -1) {
      var jumperTargetRange = outputSheet.getRange(jumperTargetRow, 1, 1, outputSheet.getLastColumn());
      jumperTargetRange.setBackground('orange'); // Highlight the row with orange color
      outputSheet.getRange('A2').setValue('JumperTarget: Row ' + jumperTargetRow); // Store the JumperTarget in cell A2
    }
  }
}


function getDuplicates(arr) {
  var counts = {};
  var duplicates = [];
  for (var i = 0; i < arr.length; i++) {
    var num = arr[i];
    counts[num] = counts[num] ? counts[num] + 1 : 1;
    if (counts[num] === 2) {
      duplicates.push(num);
    }
  }
  return duplicates;
}

function getAllIndices(arr, val) {
  var indices = [];
  for (var i = 0; i < arr.length; i++) {
    if (arr[i] === val) {
      indices.push(i);
    }
  }
  return indices;
}

function updateJumperBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  var jumperBaseRow = outputSheet.getRange('A1').getValue(); // Get the previously identified JumperBase row
  jumperBaseRow = parseInt(jumperBaseRow.split(' ')[2]); // Extract the row number

  if (jumperBaseRow !== -1) {
    var cableSizeOptions = [48, 72, 144];

    var cableFValue = outputSheet.getRange('F' + jumperBaseRow).getValue();
    var cableIValue = outputSheet.getRange('I' + jumperBaseRow).getValue();

    var cableSize = Math.abs(cableFValue - cableIValue);
    var minCableSize = cableSize * 1.15; // 15% bigger than the absolute difference between F and I

    var chosenCableSize = cableSizeOptions.filter(function(size) {
      return size >= minCableSize && size <= cableSize * 2; // Check if the size is within the specified range
    }).reduce(function(a, b) {
      return Math.abs(b - cableSize) < Math.abs(a - cableSize) ? b : a; // Choose the closest size
    });

    outputSheet.getRange('H' + jumperBaseRow).setValue('Jumper'); // Replace value in column H with "Jumper"
    outputSheet.getRange('M' + jumperBaseRow).setValue(chosenCableSize); // Put chosen cable size in column M
  }
}

function createNewRowWithJumperBaseData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  var jumperBaseRow = outputSheet.getRange('A1').getValue(); // Get the previously identified JumperBase row
  jumperBaseRow = parseInt(jumperBaseRow.split(' ')[2]); // Extract the row number

  if (jumperBaseRow !== -1) {
    var newRow = outputSheet.getLastRow() + 1; // Next available row

    var jumperBaseRange = outputSheet.getRange(jumperBaseRow, 1, 1, outputSheet.getLastColumn());
    var jumperBaseValues = jumperBaseRange.getValues();

    var newRowRange = outputSheet.getRange(newRow, 1, 1, outputSheet.getLastColumn());
    newRowRange.setValues(jumperBaseValues);

    // Highlight the new row with a unique color
    newRowRange.setBackground('#ffcc00'); // For example, using a shade of yellow

    outputSheet.getRange('A3').setValue('JumperNew: Row ' + newRow); // Store the JumperNew row in cell A1
  }
}

function updateJumperNew() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  var jumperBaseRow = outputSheet.getRange('A3').getValue(); // Get the previously identified JumperBase row
  jumperBaseRow = parseInt(jumperBaseRow.split(' ')[2]); // Extract the row number

  if (jumperBaseRow !== -1) {
    var cableSizeOptions = [48, 72, 144];

    var cableFValue = outputSheet.getRange(jumperBaseRow, 6).getValue(); // Column F is at index 6 (zero-based index)
    var cableIValue = outputSheet.getRange(jumperBaseRow, 9).getValue(); // Column I is at index 9 (zero-based index)

    var cableSize = Math.abs(cableFValue - cableIValue);
    var minCableSize = cableSize * 1.15; // 15% bigger than the absolute difference between F and I

    var chosenCableSize = cableSizeOptions.filter(function(size) {
      return size >= minCableSize && size <= cableSize * 2; // Check if the size is within the specified range
    }).reduce(function(a, b) {
      return Math.abs(b - cableSize) < Math.abs(a - cableSize) ? b : a; // Choose the closest size
    });

    outputSheet.getRange(jumperBaseRow, 11).setValue('Jumper'); // Column K is at index 11 (zero-based index)
    outputSheet.getRange(jumperBaseRow, 12).setValue(chosenCableSize); // Column L is at index 12 (zero-based index)
  }
}

function updateJumperNewOut() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  var jumperNewRow = outputSheet.getRange('A3').getValue(); // Get the previously identified JumperNew row
  jumperNewRow = parseInt(jumperNewRow.split(' ')[2]); // Extract the row number

  var jumperTargetRow = outputSheet.getRange('A2').getValue(); // Get the previously identified JumperTarget row
  jumperTargetRow = parseInt(jumperTargetRow.split(' ')[2]); // Extract the row number

  if (jumperNewRow !== -1 && jumperTargetRow !== -1) {
    var jumperTargetHValue = outputSheet.getRange(jumperTargetRow, 8).getValue(); // Column H is at index 8 (zero-based index)
    var jumperTargetMValue = outputSheet.getRange(jumperTargetRow, 13).getValue(); // Column M is at index 13 (zero-based index)

    outputSheet.getRange(jumperNewRow, 8).setValue(jumperTargetHValue); // Column H is at index 8 (zero-based index)
    outputSheet.getRange(jumperNewRow, 13).setValue(jumperTargetMValue); // Column M is at index 13 (zero-based index)
  }
}

function createTempSheetAndCopyData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  
  // Create a temporary sheet
  var tempSheet = ss.insertSheet('TempSheet');

  // Copy data from the output sheet to the temporary sheet
  outputSheet.getDataRange().copyTo(tempSheet.getRange('A1'));
}

function recalculateColumnNForJumperBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = ss.getSheetByName('TempSheet'); // Replace with your TempSheet name

  var jumperBaseRow = tempSheet.getRange('A1').getValue(); // Get the JumperBase row number
  jumperBaseRow = parseInt(jumperBaseRow.split(' ')[2]); // Extract the row number

  if (jumperBaseRow !== -1) {
    var columnGValue = tempSheet.getRange('G' + jumperBaseRow).getValue(); // Column G value
    var columnJValue = tempSheet.getRange('J' + jumperBaseRow).getValue(); // Column J value
    var columnMValue = tempSheet.getRange('M' + jumperBaseRow).getValue(); // Column M value

    var calculation = columnMValue - (columnGValue - columnJValue);
    var columnNValue = '[' + calculation + '] - [' + columnMValue + ']';

    tempSheet.getRange('N' + jumperBaseRow).setValue(columnNValue); // Set the new value for column N
  }
}

function updateJumperNewFromJumperBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = ss.getSheetByName('TempSheet'); // Replace with your TempSheet name

  var jumperBaseRowText = tempSheet.getRange('A1').getValue();
  var jumperNewRowText = tempSheet.getRange('A3').getValue();

  var jumperBaseRowIndex = extractNumericValue(jumperBaseRowText);
  var jumperNewRowIndex = extractNumericValue(jumperNewRowText);

  if (isNaN(jumperBaseRowIndex) || isNaN(jumperNewRowIndex)) {
    Logger.log("Invalid or missing row number for JumperBase or JumperNew");
    return;
  }

  var columnNValue = tempSheet.getRange('N' + jumperBaseRowIndex).getValue();
  if (!columnNValue || typeof columnNValue !== 'string') {
    Logger.log("Invalid or missing value in Column N of JumperBase row");
    return;
  }

  var values = columnNValue.match(/\[(\d+)\] - \[(\d+)\]/);
  if (!values || values.length !== 3) {
    Logger.log("Incorrect format or missing values in Column N of JumperBase row");
    return;
  }

  var jumperNewRow = jumperNewRowIndex;

  tempSheet.getRange('G' + jumperNewRow).setValue(values[2]);
  tempSheet.getRange('J' + jumperNewRow).setValue(values[1]);
}

// Function to extract numeric value from a string
function extractNumericValue(text) {
  var numericPart = text.match(/\d+/);
  return numericPart ? parseInt(numericPart[0]) : NaN;
}

function recalculateColumnNForJumperTarget() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tempSheet = ss.getSheetByName('TempSheet'); // Replace with your TempSheet name

  var jumperTargetRow = tempSheet.getRange('A2').getValue(); // Get the JumperTarget row number
  jumperTargetRow = parseInt(jumperTargetRow.split(' ')[2]); // Extract the row number

  var jumperBaseRow = tempSheet.getRange('A1').getValue(); // Get the JumperBase row number
  jumperBaseRow = parseInt(jumperBaseRow.split(' ')[2]); // Extract the row number

  var columnNValue = tempSheet.getRange('N' + jumperBaseRow).getValue();
  if (!columnNValue || typeof columnNValue !== 'string') {
    Logger.log("Invalid or missing value in Column N of JumperBase row");
    return;
  }

  var values = columnNValue.match(/\[(\d+)\] - \[(\d+)\]/);
  if (!values || values.length !== 3) {
    Logger.log("Incorrect format or missing values in Column N of JumperBase row");
    return;
  }

  if (jumperTargetRow !== -1) {
    var columnGValue = tempSheet.getRange('G' + jumperTargetRow).getValue(); // Column G value
    var columnJValue = tempSheet.getRange('J' + jumperTargetRow).getValue(); // Column J value
    var columnMValue = tempSheet.getRange('M' + jumperTargetRow).getValue(); // Column M value

      var value2 = parseInt(values[2]); // Convert value2 to a number

    var calculation = (value2 + 1);
    var columnNValue = '[' + calculation + '] - [' + columnMValue + ']';

    tempSheet.getRange('N' + jumperTargetRow).setValue(columnNValue); // Set the new value for column N
  }
}






function rearrangeColumnsInOutputSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output'); // Change to your output sheet name
  var tempSheet = ss.getSheetByName('TempSheet'); // Change to your temp sheet name

  var ranges = [
    { source: 'B', target: 'A' },
    { source: 'C', target: 'B' },
    { source: 'K', target: 'C' },
    { source: 'L', target: 'D' },
    { source: 'J', target: 'E' },
    { source: 'G', target: 'F' },
    { source: 'H', target: 'G' },
    { source: 'M', target: 'H' },
    { source: 'N', target: 'I' }
  ];

  outputSheet.clear(); // Clear the output sheet

  ranges.forEach(function(range) {
    var sourceRange = tempSheet.getRange('' + range.source + '1:' + range.source);
    var targetRange = outputSheet.getRange('' + range.target + '1');

    sourceRange.copyTo(targetRange);
  });
}

function replaceHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  // Clear contents of column B
  outputSheet.getRange('B:B').clearContent();

  // Set B1 to "Input Splice ID"
  outputSheet.getRange('B1').setValue('Input Splice ID');

}

function combineColumnsEFAndSetHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');

  var lastRow = outputSheet.getLastRow();
  var columnE = outputSheet.getRange('E2:E' + lastRow).getValues();
  var columnF = outputSheet.getRange('F2:F' + lastRow).getValues();
  var combinedValues = [];

  for (var i = 0; i < columnE.length; i++) {
    combinedValues.push([columnE[i][0] + " - " + columnF[i][0]]);
  }

  outputSheet.getRange('E2:E' + (lastRow)).setValues(combinedValues);
  outputSheet.getRange('E1').setValue("Cable In Fibers");
}

function removeColumnF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  
  outputSheet.deleteColumn(6); // Assuming row 6 is intended to be removed (adjust row number as needed)
}

function customSortOutputSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Output');

  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var data = range.getValues();

  // Custom sorting order for Column C
  var customOrder = {
    "Stub 1": 1,
    "Jumper": 2,
    "Stub 2": 3
    // Add more custom values if needed
  };

  data.sort(function (a, b) {
    var columnC_A = customOrder[a[2]]; // Column C value of row A
    var columnC_B = customOrder[b[2]]; // Column C value of row B

    if (columnC_A === undefined) columnC_A = Number.MAX_SAFE_INTEGER;
    if (columnC_B === undefined) columnC_B = Number.MAX_SAFE_INTEGER;

    if (columnC_A !== columnC_B) {
      return columnC_A - columnC_B;
    } else {
      // If Column C values are equal, sort by Column A in descending order
      return b[0].localeCompare(a[0]);
    }
  });

  range.setValues(data);
}

function calculateHighLowDifference() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var lastRow = outputSheet.getLastRow();
  var data = outputSheet.getRange('C2:E' + lastRow).getValues(); // Assuming data starts from Row 2

  var groupedData = {}; // Object to store grouped data based on Column C values

  // Grouping rows by values in Column C
  for (var i = 0; i < data.length; i++) {
    var currentValue = data[i][0]; // Column C value
    var containsJumper = data[i][0].toString().toLowerCase().indexOf('jumper') !== -1; // Check for 'jumper' keyword

    if (!containsJumper) {
      if (!groupedData[currentValue]) {
        groupedData[currentValue] = [];
      }
      
      groupedData[currentValue].push(data[i][2]); // Column E value added to group
    }
  }

  // Calculate high, low, and difference for each group
  for (var key in groupedData) {
    var groupValues = groupedData[key];
    var sortedValues = groupValues.map(function(e) {
      return e.split('-').map(Number);
    }).sort((a, b) => a[1] - b[0]);
    var highLowDifference = 0;

    if (sortedValues.length > 0) {
      var high = sortedValues[sortedValues.length - 1][1];
      var low = sortedValues[0][0];
      highLowDifference = high - low + 1;
    }

    // Place the group value in Column J and difference in Column K
    for (var j = 0; j < data.length; j++) {
      if (data[j][0] === key) {
        outputSheet.getRange(j + 2, 10).setValue(key); // Column J
        outputSheet.getRange(j + 2, 11).setValue(highLowDifference); // Column K
      }
    }
  }
}

function calculateHighLowDifferenceJump() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var lastRow = outputSheet.getLastRow();
  var data = outputSheet.getRange('A2:C' + lastRow).getValues(); // Assuming data starts from Row 2

  var jumperRow = -1;
  var highestValue = 0;
  var lowestValue = Number.MAX_VALUE;

  // Find the row where column C contains "jumper"
  for (var i = 0; i < data.length; i++) {
    if (data[i][2].toString().toLowerCase().indexOf('jumper') !== -1) {
      jumperRow = i + 2; // Adding 2 to i because data starts from Row 2, and arrays are 0-indexed
      var value = data[i][0].toString().split(':')[1].split('-').map(Number); // Extracting values after ':'
      highestValue = value[1];
      lowestValue = value[0];
      break;
    }
  }

  if (jumperRow !== -1) {
    var difference = Math.abs(highestValue - lowestValue); // Calculate the absolute difference

    // Update columns J and K of the "jumper" row
    outputSheet.getRange(jumperRow, 10).setValue(highestValue); // Column J
    outputSheet.getRange(jumperRow, 11).setValue(difference); // Column K
  } else {
    Logger.log("No row contains 'Jumper' in column C.");
  }
}

function updateColumnKBasedOnChangeRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var lastRow = outputSheet.getLastRow();
  var data = outputSheet.getRange('J2:K' + lastRow).getValues(); // Assuming data starts from Row 2

  var changeRow = -1;

  // Find the row where column J contains only numbers
  for (var i = 0; i < data.length; i++) {
    if (!isNaN(data[i][0])) {
      changeRow = i + 2; // Adding 2 to i because data starts from Row 2, and arrays are 0-indexed
      break;
    }
  }

  if (changeRow !== -1) {
    var changeRowKValue = data[changeRow - 2][1]; // Column K value for change row
    var changeRowJValue = data[changeRow - 2][0]; // Column J value for change row

    // Add 1 to changeRow Column K
    outputSheet.getRange(changeRow, 11).setValue(changeRowKValue + 1);

    // Modify other rows' Column K based on the value of the change row's Column J and K
    for (var j = 0; j < data.length; j++) {
      if (j !== (changeRow - 2)) {
        if (changeRowJValue < 433 && data[j][0].toString().toLowerCase().indexOf('stub 1') !== -1) {
          var currentValue = data[j][1];
          outputSheet.getRange(j + 2, 11).setValue(currentValue + changeRowKValue + 1);
        } else if (changeRowJValue >= 433 && data[j][0].toString().toLowerCase().indexOf('stub 2') !== -1) {
          var currentValue = data[j][1];
          outputSheet.getRange(j + 2, 11).setValue(currentValue + changeRowKValue + 1);
        }
      }
    }
  } else {
    Logger.log("No row contains only numbers in column J.");
  }
}

function updateColumnKFinal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var lastRow = outputSheet.getLastRow();

  for (var rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    var currentValue = outputSheet.getRange(rowIndex, 11).getValue(); // Get current value in Column K
    var currentValueColumnJ = outputSheet.getRange(rowIndex, 10).getValue(); // Get current value in Column J

    if (currentValue !== '' && currentValueColumnJ !== '') {
      for (var i = rowIndex + 1; i <= lastRow; i++) {
        var nextValueColumnJ = outputSheet.getRange(i, 10).getValue(); // Get next cell value in Column J

        if (currentValueColumnJ === nextValueColumnJ) {
          outputSheet.getRange(i, 11).setValue(currentValue);
        }
      }
    }
  }
}

function deduplicateColumnsJK() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var lastRow = outputSheet.getLastRow();

  var range = outputSheet.getRange('J2:K' + lastRow); // Get the range of columns J and K excluding headers
  var values = range.getValues();
  
  var uniqueValues = [];
  var uniqueRows = [];

  for (var i = 0; i < values.length; i++) {
    var currentRow = values[i].join('');
    if (!uniqueRows.includes(currentRow)) {
      uniqueRows.push(currentRow);
      uniqueValues.push(values[i]);
    }
  }

  // Clear existing data in columns J and K
  outputSheet.getRange('J2:K' + lastRow).clearContent();

  // Write deduplicated values back to columns J and K
  outputSheet.getRange(2, 10, uniqueValues.length, uniqueValues[0].length).setValues(uniqueValues);

  var data = outputSheet.getRange('J2:K' + lastRow).getValues(); // Get the data in columns J and K

  // Clear cells in columns J and K where column J contains only numbers
  for (var j = 0; j < data.length; j++) {
    var currentRow = data[j][0]; // Value in column J
    var columnKValue = data[j][1]; // Value in column K

    // Check if the cell in column J contains only numbers
    if (!isNaN(currentRow)) {
      // Clear the cell in column J and its corresponding value in column K
      outputSheet.getRange(j + 2, 10).clearContent(); // Clear cell in column J
      outputSheet.getRange(j + 2, 11).clearContent(); // Clear cell in column K
    }
  }
}

function updateColumnJAndDivideK() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName('Output');
  var lastRow = outputSheet.getLastRow();

  // Replace Stub 1 with "Splice 1 Ribbons" and Stub 2 with "Splice 2 Ribbons" in Column J
  for (var i = 2; i <= lastRow; i++) {
    var value = outputSheet.getRange(i, 10).getValue();
    if (value === 'Stub 1') {
      outputSheet.getRange(i, 10).setValue('Splice 1 Ribbons');
    } else if (value === 'Stub 2') {
      outputSheet.getRange(i, 10).setValue('Splice 2 Ribbons');
    }
  }

  // Divide values in Column K by 12
  var rangeK = outputSheet.getRange('K2:K' + lastRow);
  var valuesK = rangeK.getValues();
  for (var j = 0; j < valuesK.length; j++) {
    if (valuesK[j][0] !== '') {
      valuesK[j][0] = valuesK[j][0] / 12;
    }
  }
  rangeK.setValues(valuesK);

  var dataK = outputSheet.getRange('K2:K' + lastRow).getValues(); // Get the values in column K

  var columnLValues = [];
  
  // Loop through values in column K to populate column L only for cells with data in column K
  for (var i = 0; i < dataK.length; i++) {
    if (dataK[i][0] !== "") {
      if (dataK[i][0] < 37) {
        columnLValues.push(["FOSC 450C"]);
      } else {
        columnLValues.push(["FOSC 450D"]);
      }
    } else {
      columnLValues.push([""]);
    }
  }
  
  // Set the values in column L based on the conditions
  outputSheet.getRange(2, 12, columnLValues.length, 1).setValues(columnLValues);
}




function removeTempSheetAndFormatOutput() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Remove TempSheet
  var tempSheet = ss.getSheetByName('TempSheet');
  if (tempSheet) {
    ss.deleteSheet(tempSheet);
  }
  
  // Format Output sheet
  var outputSheet = ss.getSheetByName('Output');
  if (outputSheet) {
    // Clear color fill
    outputSheet.getDataRange().setBackground(null);
    
    // Auto resize columns to fit content
    outputSheet.getDataRange().activate();
    outputSheet.autoResizeColumns(1, outputSheet.getLastColumn());
    
    // Increase each row's width by 5 pixels
    var numCols = outputSheet.getLastColumn();
    for (var i = 1; i <= numCols; i++) {
      var currentWidth = outputSheet.getColumnWidth(i);
      outputSheet.setColumnWidth(i, currentWidth + 30);
    }
    
    // Center all cells in the sheet
    outputSheet.getDataRange().setHorizontalAlignment('center');
    
    // Bold and underline the first row
    var firstRow = outputSheet.getRange('1:1');
    firstRow.setFontWeight('bold');
    firstRow.setFontSize(11)
  }


}
