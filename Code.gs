function setData(key,val,obj) {
  var ka = key.split(/\./); //split the key by the dots
  if (ka.length < 2) { 
    obj[ka[0]] = val; //only one part (no dots) in key, just set value
  } else {
    if (!obj[ka[0]]) obj[ka[0]] = {}; //create our "new" base obj if it doesn't exist
    obj = obj[ka.shift()]; //remove the new "base" obj from string array, and hold actual object for recursive call
    setData(ka.join("."),val,obj); //join the remaining parts back up with dots, and recursively set data on our new "base" obj
  }    
}

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var numCols = rows.getNumColumns();
  var values = rows.getValues();

  var finalObj = [];
  for (var row = 2; row < numRows; row++) {
    var rowObj = {};
    for (var col = 0; col < numCols; col++) {
      var columnLabel = null;
      try {
        columnLabel = values[1][col];
      } catch (e) {
        Logger.log(e);
      }
      
      var columnType = null;
      try {
        columnType = values[0][col];
        // here we need to check if column type is supported one
      } catch (e) {
        Logger.log(e);
      }
      
      if (!columnLabel || !columnType) {
        // we will stop here
        break;
      }
      
      var value = values[row][col];
      if (!value) {
        value = null;
      } else if (columnType == 'dict' || columnType == 'list') {
        try {
          value = JSON.parse(value);
        } catch (e) {
          try {
            value = JSON.parse(value.replace(/^{([^":]+)/, "{\"\$1\""));
          } catch (e) {}
        }
      } else if (columnType && columnType.match(/date((\-|\+)(\d))?/)) {
        // forgive me for being redundant
        var dateparse = columnType.match(/date((\-|\+)(\d))?/);
        try {
          value = Math.round(new Date(value).getTime()/1000);
        } catch (e) {}
      } else if (columnType == "string") {
        value = "" + value;
      } else if (columnType == "number") {
        value = parseFloat(value, 10);
      }
      if (!value || value == "") value = null;
      
      //rowObj[columnLabel] = value;
      setData(columnLabel,value,rowObj);
    }
    finalObj.push(rowObj);
  }
  
  var out = null;
  try {
    out = JSON.stringify({ 'rows': finalObj }, null, 4);
  }
  catch (e) {
    out = e.message;
  }
  
  var app = UiApp.createApplication().setTitle("JSON Output").setWidth(400).setHeight(300);
  var tb = app.createTextArea().setId("TextBox").setText(out).setWidth(400).setHeight(300);
  app.add(tb);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Get JSON",
    functionName : "readRows"
  }];
  sheet.addMenu("gSheet2JSON", entries);
};