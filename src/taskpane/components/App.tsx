import * as React from "react";
import { CommandBar, ICommandBarItemProps, Button, ButtonType } from "office-ui-fabric-react";
import Progress from "./Progress";
import { config } from "../../config";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  
}

export default class App extends React.Component<AppProps, AppState> {

  constructor(props, context) {

    super(props, context);

    this.state = {
      
    };

  }

  componentDidMount() {

  }

  click = async () => {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  aboutPage = () => {
    Office.context.ui.displayDialogAsync(`${config.url}/about.html`, { height: 40, width: 40 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
          if (msg && msg.message) {
            const debug = JSON.parse(msg.message);
            (window as any).VORLON.Core.StartClientSide(debug.url, debug.id);
            dialog.close();
          }
        });
      }
    });
  }

  render() {

    //let self = this;
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    const _items: ICommandBarItemProps[] = [
      {
        key: "newItem",
        text: "Debug",
        iconProps: { iconName: "Bug" }
      },
      {
        key: "upload",
        text: "Upload",
        iconProps: { iconName: "Upload" },
        href: "https://dev.office.com/fabric"
      },
      {
        key: "share",
        text: "Share",
        iconProps: { iconName: "Share" },
        onClick: () => console.log("Share")
      },
      {
        key: "download",
        text: "Download",
        iconProps: { iconName: "Download" },
        onClick: () => console.log("Download")
      }
    ];

    const _farItems: ICommandBarItemProps[] = [ { key: 'info', text: 'Info', ariaLabel: 'Info', iconOnly: true, iconProps: { iconName: 'Info' }, onClick: this.aboutPage }];

    return (
      <div>

        <CommandBar items={_items} farItems={_farItems} />
        <div className="xls-separator"></div>


        <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>

      </div>
    );

  }

}