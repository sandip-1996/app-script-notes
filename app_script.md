# app script notes
## get sheetname using onedit
``` javascript
#  ignore multiple edits 
if(e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
    return;
}
const activesheet = e.range.getSheet();
const sheetname = activesheet.getName();
const activerow = e.range.getRow();
const eachrecord = activesheet.getRange(activerow, 1, 1, CONFIG.FORM_RESPONSE_HEADER.length).getValues()[0];  
``` 
## global variable declaration 
``` javascript
let date; // it means date is used as global variable inside if block
if () {
    date = ()
    return date
}
```
## isNumeric check
```javascript
// javascript doesnot have bydefault isNumeric
function isNumeric(value) {
    return !isNan(value) && value !==;
}
```
## slice string
```javascript
// in javascript slice function can operate on string not on Number
const input  = String(value)
const left = input.slice(0,1) // slice start from 0 end before 1
const right = input.slice(1,2) // slice start from 0 end before 2
```
## date parsing
```javascript
const year = new Date().getFullYear();
const month = new Date().getMonth() + 1;

const date = new Date(year, month - 1, daynumber)
```
## write and read single value in cell
```javascript
const userinput = activesheet.getRange(activeRow, 1).getvalue();
const cell = activesheet.getRange(activeRow, 1)
cell.setValue(date)
cell.setNumberFormat("dd-mm-yyyy")
// setvalues() require 2D data
```
## get error box
```javascript
SpreadsheetApp.getUi().alert()
```
## set validation_list
```javascript
function validation_list(activesheet, activerow) {

  const productType = String(
    activesheet.getRange(activerow, 2).getValue()
  ).trim();

  const validationMap = {
    "VESTS": "vests_list",
    "TRUNK": "trunk_list",
    "KIDS": "kids_list"
  };

  const listName = validationMap[productType];

  if (listName === undefined) {
    return;
  }

  const range = SpreadsheetApp.getActiveSpreadsheet()
    .getRange(listName);

  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(false)
    .build();

  activesheet
    .getRange(activerow, 5)
    .setDataValidation(validation);
}
```
## set Ui
```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Initialize_sheet")
  .addItem("Initialize Structure All", "initializeStructure")
  .addItem("Initialize Processed Sheet", "initializeProcessedSheet")
  .addItem("Initialize Invalid Sheet", "initializeInvalidSheet")
  .addItem("Initialize Duplicate Sheet", "initializeDuplicateSheet")
  .addToUi();

  ui.createMenu("Initialize_Row").addItem("Initialize Row", "initializeRow").addToUi();

  ui.createMenu("Refresh_data").addItem("Refresh Data", "refresh_data").addToUi();
}

function initializeStructure() {
  initializeProcessedSheet(ignore = true);
  initializeInvalidSheet(ignore = true);
  initializeDuplicateSheet(ignore = true);

  // initializeRow()
  SpreadsheetApp.getUi().alert("Initialized all Sheets Successfully.");
};
```
## detect non empty row
```javascript
isEmpty = values.every(value=> String(value).trim() === "");
// returns true where all rows are empty
isEmpty = values.every(value => String(value).trim() !== "");
// returns true where at least one row are empty
```
## check in list
```javascript
  const typeList = ["CAPRI", "FULL", "HALF"]
  if (!typeList.includes(data)) {
    return;
  }
concept
const typeList = ["CAPRI", "FULL", "HALF"];

console.log(1 in typeList); // true: index 1 exists
console.log("FULL" in typeList); // false: not a property key
console.log(typeList.includes("FULL")); // true: value exists
```
## forEach
```javascript
  map = {}
  header.forEach((header, index) => {map[String(header).trim()]=index;})
```
## orderDatabyHeader
```javascript
return headers.map(header => {return data[header] ?? ""});
```
## getSheetbyName
```javascript
const cleanSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CLEAN_DATA");
```
## writeData
```javascript
const lastrow = cleanSpreadsheet.getLastRow()
const nextRow = lastRow + 1
cleanSpreadsheet.getRange(nextrow,1,1,CONFIG.CLEAN_HEADER.length).setValues([data]);
```
```text
map() returns a new array.
forEach() returns undefined.
filter() returns a new filtered array.
reduce() returns an accumulated result.
```

