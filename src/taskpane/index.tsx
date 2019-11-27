import React from "react";
import ReactDOM from "react-dom";
import { App } from "./components";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { StateProvider, IAppTools, IExcelContext } from '../state'
import "office-ui-fabric-react/dist/css/fabric.min.css";

initializeIcons();

let state: IAppTools;
let evData: OfficeExtension.EventHandlerResult<Excel.WorksheetChangedEventArgs>;

ReactDOM.render(<AppContainer><StateProvider tools={o => state = o }><App /></StateProvider></AppContainer>, document.getElementById("container"));

const inferCellType = (range: Excel.Range, row: number, column?: number) => {

  if (!column) column = 0; //if range is for a single column
  const columnType = range.valueTypes[row][column];
  let dataType: string = null;

  switch (columnType) {
    case Excel.RangeValueType.boolean:
      dataType = 'boolean';
      break;
    case Excel.RangeValueType.double:
      const columnFormat = range.numberFormat[row][column];
      if (/(?:yy|mm|dd)/i.test(columnFormat)) dataType = 'date';
      else dataType = 'number';
      break;
    case Excel.RangeValueType.integer:
      dataType = 'number';
      break;
    case Excel.RangeValueType.string:
      dataType = 'string';
      break;
  }

  return dataType;

};

const readFile = (sheetId?: string): void => {
  Excel.run<IExcelContext>((context) => {

    let result: IExcelContext = { };
    let sheet: Excel.Worksheet;
    let workbook = context.workbook;

    if (sheetId) sheet = workbook.worksheets.getItem(sheetId); // TODO: need this ??????
    else sheet = workbook.worksheets.getActiveWorksheet();

    let range = sheet.getUsedRange();
    let headerRow = range.getRow(0);
    let tables = sheet.tables;
    let charts = sheet.charts;

    sheet.load(['id', 'name']);
    range.load(['columnIndex', 'rowIndex', 'rowCount', 'columnCount']);
    headerRow.load(['text']);
    tables.load(['id', 'name']); 
    charts.load(['id', 'name', 'title']);
    
    let tableData: { key: string, name: string, headers: Excel.Range, range: Excel.Range }[];
    let columnData: { index: number, used?: Excel.Range, range?: Excel.Range }[];

    // evData = sheet.onChanged.add(dataChanged);
    // workbook.worksheets.onActivated.add(sheetChanged);

    return context.sync()
      .then(() => {
        result.currentSheet = { key: sheet.id, name: sheet.name };
        result.currentSheet.columns = headerRow.text[0].map((x, i) => ({ key: x, index: i + range.columnIndex })).filter(x => x.key && x.key.length > 0);
        result.currentSheet.charts = charts.items.map((x) => ({ key: x.id, name: x.name, title: x.title.text }));
      })
      .then(() => {

        tableData = tables.items.map(table => {
          const tableItem = { key: table.id, name: table.name, headers: table.getHeaderRowRange(), range: table.getRange() };
          tableItem.headers.load('text');
          tableItem.range.load(['columnIndex', 'rowIndex', 'rowCount', 'columnCount', 'valueTypes', 'numberFormat']);
          return tableItem;
        });

        return context.sync();

      })
      .then(() => {

        result.currentSheet.tables = tableData.map(table => {
          return {
            key: table.key,
            name: table.name,
            columns: table.headers.text[0].map((key, index) => ( { key, index }))
          };
        });

        columnData = result.currentSheet.columns.map(column => {
          const tablesOnThisColumn = tableData.filter(table => column.index >= table.range.columnIndex && column.index <= (table.range.columnIndex + table.range.columnCount)).map(x => x.range.rowIndex);
          tablesOnThisColumn.push(range.rowCount);
          const maxRowToSearch = Math.min(...tablesOnThisColumn);
          const columnItem = { index: column.index, used: sheet.getRangeByIndexes(0, column.index, maxRowToSearch, 1).getUsedRangeOrNullObject(true) };
          columnItem.used.load(['rowCount']);
          return columnItem;
        });

        return context.sync();

      })
      .then(() => {

        columnData.forEach(x => {
          x.range = sheet.getRangeByIndexes(0, x.index, x.used.rowCount, 1);
          x.range.load(['valueTypes', 'numberFormat']);
        });

        return context.sync();

      })
      .then(() => {

        columnData.forEach(column => {
          const columnRange = column.range;
          let columnTypes: string[] = [];
          for (let j = 1; j < columnRange.valueTypes.length; j++) { 
            const dataType = inferCellType(columnRange, j);
            if (dataType && columnTypes.indexOf(dataType) === -1) columnTypes.push(dataType);
            if (columnTypes.length > 1) break;
          }
          if (columnTypes.length === 1) result.currentSheet.columns.filter(x => x.index === column.index)[0].type = columnTypes[0];
        });

        result.currentSheet.tables.forEach(table => {
          const data = tableData.filter(x => x.key == table.key)[0];
          const tableRange = data.range;
          for (let i = 0; i < table.columns.length; i++) {
            let columnTypes: string[] = [];
            for (let j = 1; j < tableRange.valueTypes[i].length; j++) {
              const dataType = inferCellType(tableRange, j, i);
              if (dataType && columnTypes.indexOf(dataType) === -1) columnTypes.push(dataType);
              if (columnTypes.length > 1) break;
            }
            if (columnTypes.length === 1) table.columns[i].type = columnTypes[0];
          }
        });

      })
      .then(() => {
        state.load(result);
        return result;
      });

  });
};

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) readFile();
  else state.mock();
});

const sheetChanged = (args: Excel.WorksheetActivatedEventArgs) => {

  let sheetId = args.worksheetId;
  
  const removeHandler = () => { //https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events#remove-an-event-handler
    if (evData) {
      return Excel.run(evData.context, (context) => {
        evData.remove();
        return context.sync().then(() => { evData = null; });
      });
    }
    return Promise.resolve();
  };

  return removeHandler().then(() => readFile(sheetId));

};

const dataChanged = (args: Excel.WorksheetChangedEventArgs) => {
  
  console.log(args.address);

  return new Promise(resolve => {
    resolve();
  });

};
