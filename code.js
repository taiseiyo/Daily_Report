let choose_Active_Sheet = (num) =>{
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let name = spreadsheet.getSheets()[num].getName();
  let activesheet = SpreadsheetApp.getActive().getSheetByName(name);
  return activesheet;
};

// Sheet1 script

let colorCell = () => {
  let spreadsheet = choose_Active_Sheet(0);
  let range = spreadsheet.getRange(2, 2, 1, spreadsheet.getLastColumn() - 1);
  let time_values = range.getValues()[0];

  for (let i in time_values) {
    let tar_cell = spreadsheet.getRange(2, Number(i) + 2);

    if (Number(time_values[i]) >= 0.5) {
      tar_cell.setBackground("#7FFF00");
    } else {
      tar_cell.setBackground("#FFC0CB");
    }
  }
};

let deleteSheet1Cell = () => {
  let spreadsheet = choose_Active_Sheet(0);
  let range = spreadsheet.getRange(2, 2, 1, spreadsheet.getLastColumn() - 1);
  let tmp = SpreadsheetApp.getActiveSheet().getRange("A70:B75");
  tmp.clearContent();
  range.clear();
};

let graph = () => {
  let spreadsheet = choose_Active_Sheet(0);;
  let range = spreadsheet.getDataRange().getValues();
  range = range.splice(0, 2);
  let values = [];

  for (let i in range[0]) {
    values.push([range[0][i], range[1][i]]);
  }

  let tmp = spreadsheet.getRange("A70:B75");
  tmp.setValues(values);

  if (spreadsheet.getCharts()[0] == undefined) {
    makeGraph(spreadsheet, tmp);
  } else {
    changeGraph(spreadsheet, tmp);
  }
};

let makeGraph = (spreadsheet, tmp) => {
  let graph = spreadsheet
    .newChart()
    .addRange(tmp)
    .setChartType(Charts.ChartType.LINE)
    .setPosition(5, 5, 0, 0)
    .setOption("title", "作業時間")
    .build();

  spreadsheet.insertChart(graph);
};

let changeGraph = (spreadsheet) => {
  let chart = spreadsheet.getCharts()[0];
  let newchart = chart.modify().setChartType(Charts.ChartType.LINE).build();

  spreadsheet.updateChart(newchart);
};

// Sheet2 script
