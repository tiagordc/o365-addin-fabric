import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();

let state = { loaded: false, id: null, sheet: null, changed: null };

Office.onReady(info => { // Get info of the current spreadsheet

  state.loaded = false;
  render(App);

  if (info.host === Office.HostType.Excel) {
    Excel.run((context) => {

      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(['id']);
      
      let changedEvent = sheet.onChanged.add(dataChanged);

      context.workbook.worksheets.onActivated.add(sheetChanged);
      
      return context.sync().then(() => {
        state.loaded = true;
        state.id = sheet.id;
        state.sheet = sheet;
        state.changed = changedEvent;
        render(App);
      });

    });

  }
});

const sheetChanged = (eArgs: Excel.WorksheetActivatedEventArgs) => {

  state.loaded = false;
  render(App);

  let sheetId = eArgs.worksheetId;
  if (sheetId === state.id) return Promise.resolve();

  const removeHandler = () => { //https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events#remove-an-event-handler
    if (state.changed) {
      return Excel.run(state.changed.context, (context) => {
        state.changed.remove();
        return context.sync().then(() => { delete state.changed; });
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
        state.loaded = true;
        state.id = sheet.id;
        state.sheet = sheet;
        state.changed = changedEvent;
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
      <Component loaded={state.loaded} worksheet={state.id} />
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
