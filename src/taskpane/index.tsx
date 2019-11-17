import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();

let isOfficeInitialized = false;
let worksheetId = null;

const title = "Contoso Task Pane Add-in";

const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} worksheet={worksheetId} />
    </AppContainer>,
    document.getElementById("container")
  );
};

Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    Excel.run((context) => {

      var sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(['id']);

      context.workbook.worksheets.onActivated.add((eArgs) => {
        return new Promise(resolve => {
          worksheetId = eArgs.worksheetId;
          render(App);
          resolve();
        });
      });

      return context.sync().then(() => {
        worksheetId = sheet.id;
        render(App);
      });

    });
  }
});


render(App); //Initial render showing a progress bar

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
