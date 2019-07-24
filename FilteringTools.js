/* Google Sheets app script - custom filtering tools for cells based on colour, tld extension, etc.
Functions transforms cell tables to filtered lists. 

Installation: in Google Sheets, navigate to Tools > Script Editor. Copy this script to a new file and save. 
Re-load your Sheet. Grant the script permission from your Google account, and a custom "Filtering Tools" menu will appear.

Usage: Execute functions via their buttons on the "Filtering Tools" menu.

Script binds to the Google Sheet it is installed on. */

/* MENU */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Filtering Tools')
    .addItem('List .com Cells', 'menu_getDotComs')
    .addSeparator()
    .addItem('List .co.uk Cells', 'menu_getCoUks')
    .addSeparator()
    .addItem('List Coloured Cells', 'menu_getColouredCells')
    .addSeparator()
    .addItem('List Blue Cells', 'menu_getBlueCells')
    .addSeparator()
    .addItem('Convert to HTML Table', 'menu_tableToHTML')
    .addSeparator()
    .addItem('Convert Table to List', 'menu_tableToList')
    .addSeparator()
    .addItem('Extract Links from HTML', 'menu_extractHrefsAndTitles')
    .addToUi();
}

/*  MENU ITEM FUNCTIONS  */

function menu_getDotComs() {
  var sheet = getActiveSheet();
  var inputRange = getUserInputRange();
  var data = sheet.getRange(inputRange);
  var outputCol = getOutputColumn();
  
  
  var cells = getDomains(data, ".com"); 
  var filterName = ".com";
  SpreadsheetApp.getUi().alert('Selecting .coms');
  listCells(sheet, cells, outputCol, inputRange, filterName);
}

function menu_getCoUks() {
  var sheet = getActiveSheet();
  var inputRange = getUserInputRange();
  var data = sheet.getRange(inputRange);
  var outputCol = getOutputColumn();
  
  SpreadsheetApp.getUi().alert('Selecting .co.uks');
  var cells = getDomains(data, ".co.uk"); 
  var filterName = ".co.uk"  
  listCells(sheet, cells, outputCol, inputRange, filterName);
}

function menu_getColouredCells() {
  var sheet = getActiveSheet();
  var inputRange = getUserInputRange();
  var data = sheet.getRange(inputRange);
  var outputCol = getOutputColumn();
  
  SpreadsheetApp.getUi().alert('Listing Coloured Cells');
  var colouredCells = getColouredCells(data);  
  var filterName = "Coloured"
  listCells(sheet, colouredCells, outputCol, inputRange, filterName);
}

function menu_getBlueCells() {
  var sheet = getActiveSheet();
  var inputRange = getUserInputRange();
  var data = sheet.getRange(inputRange);
  var outputCol = getOutputColumn();
  
  SpreadsheetApp.getUi().alert('Listing Blue Cells');
  var blueCells = getBlueCells(data);   
  var filterName = "Blue" 
  listCells(sheet, blueCells, outputCol, inputRange, filterName);
}

function menu_tableToHTML() {
  var sheet = getActiveSheet();
  var inputRange = getUserInputRange();
  var data = sheet.getRange(inputRange);
  var outputCell = getOutputCell();
  
  SpreadsheetApp.getUi().alert('Creating HTML Table');
  htmlTable = tableToHTML(data) // convert spreadsheet table to HTML table string
  sheet.getRange(outputCell).setValue(htmlTable) // write HTML table string to outputCell
}

function menu_tableToList() {
  var sheet = getActiveSheet();
  var inputRange = getUserInputRange();
  var data = sheet.getRange(inputRange);
  var outputCol = getOutputColumn();
  
  SpreadsheetApp.getUi().alert('Creating List from Table');
  var cells = tableToList(data) // convert spreadsheet table to single column
  listCells(sheet, cells, outputCol, inputRange);
}

function menu_extractHrefsAndTitles() {
  var sheet = getActiveSheet();
  var inputRange = getUserInputRange();
  var data = sheet.getRange(inputRange);
  var outputCol_Hrefs = getOutputColumn();
  var outputCol_Titles = getNextColumnLetter(outputCol_Hrefs);
  
  var linksData = extractHrefsAndTitles(data)
  var hrefs = linksData.hrefs
  var titles = linksData.titles
  
  listCells(sheet, hrefs, outputCol_Hrefs, inputRange);
  listCells(sheet, titles, outputCol_Titles, inputRange);
}
  
/*  DATA OPERATORS */

// Grab all cells containing a tld extension - e.g. ".com"
function getDomains(data, tld){
  var vals = data.getValues(); 
  var dot_coms = [];
  
   for (var i = 0; i < vals.length; i++) {
    for (var j = 0; j < vals[i].length; j++) {
      var cellVal = vals[i][j]
      
      if (typeof cellVal !== "string") continue;  // skip non-string cells
      
      if (cellVal.indexOf(tld) !== -1) {  // select only cells with the tld
        dot_coms.push(vals[i][j])
      }  
    }
  }
  return dot_coms
}
  
function getColouredCells(data) {
  var backgrounds = data.getBackgrounds();
  var vals = data.getValues();
  
  var coloured_cells = []; 
//  Backgrounds and vals share the same indexing - backgrounds[5][3] is the color of vals[5][3], and so on.  
//  Loop through all cells, grab only the coloured ones.
  for (var i = 0; i < backgrounds.length; i++) {
    for (var j = 0; j < backgrounds[i].length; j++) {
      if (backgrounds[i][j] !== "#ffffff") {
        coloured_cells.push(vals[i][j])
      }  
    }
  }
  return coloured_cells;
}

function getBlueCells(data) {
  var backgrounds  = data.getBackgrounds();
  var vals = data.getValues();
  var blue_cells = []; 
  
//  Loop through all cells, grab only the blue ones with color #00ffff
  for (var i = 0; i < backgrounds.length; i++) {
    for (var j = 0; j < backgrounds[i].length; j++) {
      if (backgrounds[i][j] === "#00ffff") {
        blue_cells.push(vals[i][j])
      }  
    }
  }
  return blue_cells;
}

function tableToHTML(data) {
  var vals = data.getValues();
  
  var outputHtml = "<table>"
  
  for (var i = 0; i <vals.length; i++ ){    // loop through table rows 
      var htmlRow = "<tr>"
      for (var j = 0; j < vals[i].length; j++ ){    // loop through table cells in a row
        if (vals[i][j] !== ""){
          htmlRow += "<td>"+vals[i][j]+"</td>"
        }
      }
      htmlRow +="</tr>"
   outputHtml += htmlRow
  }
  outputHtml += "</table>"  
  return outputHtml
}

function tableToList(data) {
  var vals = data.getValues();
  
  var list = []
  for (i = 0; i < vals.length; i++ ){   
    for (j = 0; j < vals[i].length; j++ ){
      if (vals[i][j] !== ""){
        list.push(vals[i][j])
      }
    }
  }
  return list
}
  
// Get all hrefs and titles from a column of HTML link tags.  
function extractHrefsAndTitles(data) {
  var linkList = data.getValues();
  var hrefsCol = []
  var titlesCol = []
  for (i = 0; i < linkList.length; i++ ){  
    var linkData = getHrefAndTitle(linkList[i][0]);  // get href and title from string contents of cell
    hrefsCol.push([linkData.href])  
    titlesCol.push([linkData.title])
  }
    
  return {hrefs: hrefsCol, titles: titlesCol}
}

function getHrefAndTitle(linkTag) {
  var startHrefIdx = linkTag.indexOf("href") + 6   
  var hrefLength = linkTag.slice(startHrefIdx).indexOf('">') 
  var endHrefIdx = startHrefIdx + hrefLength
  
  var startTitleIdx = endHrefIdx + 2                                             
  var endTitleIdx = linkTag.indexOf("</a>")  
  
  var href = linkTag.slice(startHrefIdx, endHrefIdx)
  var title = linkTag.slice(startTitleIdx, endTitleIdx)
  
  return {href: href, title: title}                          
};  
              
// Write a list of cell values to a column
function listCells(sheet, cells, outputCol, inputRange, filterName) {
  if (filterName === undefined) filterName = "Processed"    // set default filterName value

  if (cells.length === 0) {
    SpreadsheetApp.getUi().alert('No cells to list');
    return null;
  }
   
  firstCell = sheet.getRange(outputCol + '1' )
  firstCell.setValue([filterName + ' cells from '+ inputRange]) // set the column title
  outputCol = sheet.getRange( outputCol + "2:" + outputCol + (cells.length + 1) )  // get the range of the output column
  
  var columnData = convertCellsToCol(cells)
  
   outputCol.setValues(columnData) 
}

/* HELPER FUNCTIONS */

function getActiveSheet () {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  return sheet;
}

function getUserInputRange () {
  var inputRange = SpreadsheetApp.getUi().prompt("Please enter the input range - e.g. A1:A10").getResponseText();
  return inputRange;
}

function getOutputColumn () {
  var outputCol = SpreadsheetApp.getUi().prompt("Display results in which column?").getResponseText();
  return outputCol
}

function getOutputCell () {
  var outputCell = SpreadsheetApp.getUi().prompt("Display result in which cell? (e.g. B5)").getResponseText();
  return outputCell
}

function convertCellsToCol(cells) {
  var columnData = []
  for (var i = 0; i < cells.length; i++){   
    columnData.push([cells[i]]);  // append a row to the column 
  }
  return columnData;
}

function getNextColumnLetter(letter) {
  var nextLetter;
  
  if (letter === "Z") nextLetter = "AA";
  
  if (letter.length === 1) {
    var nextLetter =  String.fromCharCode(letter.charCodeAt(0) + 1); 
  } else if (letter.length === 2) {
    var nextLetter = letter[0] +  String.fromCharCode(letter.charCodeAt(1) + 1);
  }
  
  return nextLetter;
}
  
function countColoured(inputRange) {  // Return the number of coloured cells in a range
  coloured = getColouredCells(inputRange);
  num =  [coloured.length];
  return num
}