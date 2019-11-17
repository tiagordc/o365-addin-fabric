import * as React from "react";
import { CommandBar, ICommandBarItemProps, Button, ButtonType } from "office-ui-fabric-react";
import Progress from "./Progress";
import { config } from "../../config";

export interface IAppProps {
  title: string;
  isOfficeInitialized: boolean;
  worksheet: string;
}

export interface IAppState {
  loading: boolean;
  debug: boolean;
  sheet: string;
}

export default class App extends React.Component<IAppProps, IAppState> {

  public static getDerivedStateFromProps(nextProps: IAppProps, prevState: IAppState) {
    let result: any = {};
    if (nextProps.isOfficeInitialized && prevState.loading) result.loading = false;
    if (nextProps.worksheet !== prevState.sheet) result.sheet = nextProps.worksheet;
    if (Object.keys(result).length === 0) return null;
    return result;
  }

  constructor(props: IAppProps, context) {

    super(props, context);

    this.state = {
      debug: false,
      loading: true,
      sheet: props.worksheet
    };

  }

  debug = (msg: string, ...args: any[]) => {
    if (this.state.debug) {
      console.log(msg, args);
    }
  }

  click = async () => {
    try {
      await Excel.run(async context => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        range.format.fill.color = "yellow";
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  aboutPage = () => {

    const self = this;

    Office.context.ui.displayDialogAsync(`${config.url}/about.html`, { height: 40, width: 40 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
          if (msg && msg.message) {
            self.setState({ debug: true });
            const debug = JSON.parse(msg.message);
            (window as any).VORLON.Core.StartClientSide(debug.url, debug.id);
            dialog.close();
          }
        });
      }
    });

  }

  render() {

    let self = this;

    if (self.state.loading) {
      return (
        <Progress title={self.props.title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    const _items: ICommandBarItemProps[] = [
      { key: "upload", text: "Upload", iconProps: { iconName: "Upload" }, onClick: () => console.log("Upload") },
      { key: "share", text: "Share", iconProps: { iconName: "Share" }, onClick: () => console.log("Share") },
      { key: "download", text: "Download", iconProps: { iconName: "Download" }, onClick: () => console.log("Download") }
    ];

    const _farItems: ICommandBarItemProps[] = [ { key: 'info', text: 'Info', ariaLabel: 'Info', iconOnly: true, iconProps: { iconName: 'Info' }, onClick: this.aboutPage }];

    return (
      <div>
        <CommandBar items={_items} farItems={_farItems} />
        <div className="xls-separator"></div>
        <p>{this.state.sheet}</p>

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