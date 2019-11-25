import React from "react";
import ReactDOM from "react-dom";
import { App } from "./components";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { StateProvider, IAppTools } from '../state'
import "office-ui-fabric-react/dist/css/fabric.min.css";

initializeIcons();

let state: IAppTools;
let evData: OfficeExtension.EventHandlerResult<Excel.WorksheetChangedEventArgs>;

ReactDOM.render(<AppContainer><StateProvider tools={o => state = o }><App /></StateProvider></AppContainer>, document.getElementById("container"));

const loadSheet = (sheet: Excel.Worksheet) => {
  state.load({
    currentSheet: {
      key: sheet.id,
      name: sheet.name
    }
  });
};

Office.onReady(info => { //Get info of the current spreadsheet
  if (info.host === Office.HostType.Excel) {
    Excel.run((context) => {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(['id', 'name']);  
      evData = sheet.onChanged.add(dataChanged);
      context.workbook.worksheets.onActivated.add(sheetChanged);
      return context.sync().then(() => loadSheet(sheet) );
    });
  }
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

  return removeHandler().then(() => {
    return Excel.run((context) => {
      let sheet = context.workbook.worksheets.getItem(sheetId);
      sheet.load(['id', 'name']);
      evData = sheet.onChanged.add(dataChanged);
      return context.sync().then(() => loadSheet(sheet) );
    });
  });

};

const dataChanged = (args: Excel.WorksheetChangedEventArgs) => {
  
  console.log(args.address);

  return new Promise(resolve => {
    resolve();
  });

};
