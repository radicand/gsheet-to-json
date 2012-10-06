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
      var columnLabel;
      var cellType;
      var columnHeaders;
      
      try {
        columnLabel = values[1][col];
        cellType = values[0][col];
        if (!cellType) {
          cellType = "string";
        }
      } catch (e) {
        columnLabel = "Xx_undefined";
      }
      
      var value = values[row][col];
      if (!value) {
        value = null;
      } else if (cellType == 'dict' || cellType == 'list') {
        try {
          value = JSON.parse(value);
        } catch (e) {
          try {
            value = JSON.parse(value.replace(/^{([^":]+)/, "{\"\$1\""));
          } catch (e) {}
        }
      } else if (cellType && cellType.match(/date((\-|\+)(\d))?/)) {
        // forgive me for being redundant
        var dateparse = cellType.match(/date((\-|\+)(\d))?/);
        try {
          value = Math.round(new Date(value).getTime()/1000);
          if (dateparse[2] && dateparse[3]) {
            value -= (new Date().getTimezoneOffset() * 60)
            if (dateparse[2] == "+") {
              value -= (parseInt(dateparse[3],10) * 60 * 60)
            } else if (dateparse[2] == "-") {
              value += (parseInt(dateparse[3],10) * 60 * 60)
            }
          }
        } catch (e) {}
      } else if (cellType == "string") {
        value = "" + value;
      } else if (cellType == "number") {
        value = parseFloat(value, 10);
      }
      if (!value || value == "") value = null;
      rowObj[columnLabel] = value;
    }
    finalObj.push(rowObj);
  }
  var out = "{\"rows\":[\n";
  for (i in finalObj) {
    out+= JSON.stringify(finalObj[i]) + ",\n";
  } out = out.replace(/,\n$/,'\n') + "]}";
  
  var app = UiApp.createApplication().setTitle("JSON Output").setWidth(400).setHeight(300);
  var tb = app.createTextArea().setId("TextBox").setText(out).setWidth(400).setHeight(300);
  app.add(tb);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

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

