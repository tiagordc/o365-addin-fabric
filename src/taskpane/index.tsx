import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();

let current = { initialized: false, id: null, sheet: null, changed: null };

Office.initialize = () => {
  current.initialized = true;
  render(App);
};

Office.onReady(info => { // Get info of the current spreadsheet

  if (info.host === Office.HostType.Excel) {
    Excel.run((context) => {

      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(['id']);
      
      let changedEvent = sheet.onChanged.add(dataChanged);

      context.workbook.worksheets.onActivated.add(sheetChanged);
      
      return context.sync().then(() => {
        current.id = sheet.id;
        current.sheet = sheet;
        current.changed = changedEvent;
        render(App);
      });

    });

  }
});

const sheetChanged = (eArgs: Excel.WorksheetActivatedEventArgs) => {

  let sheetId = eArgs.worksheetId;
  if (sheetId === current.id) return Promise.resolve();

  const removeHandler = () => { //https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events#remove-an-event-handler
    if (current.changed) {
      return Excel.run(current.changed.context, function (context) {
        current.changed.remove();
        return context.sync().then(() => { delete current.changed; });
      });
    }
    return Promise.resolve();
  };

  return removeHandler().then(() => {

    return Excel.run((context) => {
  
      let sheet = context.workbook.worksheets.getItem(sheetId);
      sheet.load(['id']);
  
      let changedEvent = sheet.onChanged.add(dataChanged);
  
      return context.sync().then(() => {
        current.id = sheet.id;
        current.sheet = sheet;
        current.changed = changedEvent;
        render(App);
      });
  
    });

  });

};

const dataChanged = (eArgs: Excel.WorksheetChangedEventArgs) => {
  
  console.log(eArgs.address);

  return new Promise(resolve => {
    resolve();
  });

};

const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component initialized={current.initialized} worksheet={current.id} />
    </AppContainer>,
    document.getElementById("container")
  );
};

render(App); //Initial render showing a progress bar

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
