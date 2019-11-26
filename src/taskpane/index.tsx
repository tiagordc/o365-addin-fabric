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

// let debug = "";
// const render = () => {
//   ReactDOM.render(
//     <AppContainer>
//       <StateProvider tools={o => state = o}>
//         <div>{debug}</div>
//       </StateProvider>
//     </AppContainer>, document.getElementById("container"));
// }
// render();

const readFile = (sheetId?: string): Promise<IExcelContext> => {
  return Excel.run<IExcelContext>((context) => {

    // https://docs.microsoft.com/en-us/javascript/api/excel/....

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
    range.load(['columnIndex']);
    headerRow.load(['text']);
    tables.load(['id', 'name']); // excel.interfaces.tableloadoptions?view=excel-js-preview
    charts.load(['id', 'name', 'title']); // excel.interfaces.chartloadoptions?view=excel-js-preview
    
    // evData = sheet.onChanged.add(dataChanged);
    // workbook.worksheets.onActivated.add(sheetChanged);

    return context.sync()
      .then(() => {
        result.currentSheet = { key: sheet.id, name: sheet.name };
        result.currentSheet.columns = headerRow.text[0].map((x, i) => { return { key: x, index: i + range.columnIndex }; }).filter(x => x.key && x.key.length > 0);
        result.currentSheet.charts = charts.items.map((x) => { return { key: x.id, name: x.name, title: x.title.text }; });
      })
      .then(() => {
        let tablesHeaders = tables.items.map(x => x.getHeaderRowRange());
        tablesHeaders.forEach(x => x.load('text'));
        return context.sync(tablesHeaders);
      })
      .then((tablesHeaders) => {

        result.currentSheet.tables = tables.items.map((table, tableIndex) => { 
          return { 
            key: table.id, 
            name: table.name,
            columns: tablesHeaders[tableIndex].text[0].map((key, index) => { return { key, index }; })
          }; 
        });
        
        state.load(result);
        // debug = JSON.stringify(result);
        // render();
        
      })
      .then(() => {
        return Promise.resolve(result);
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
