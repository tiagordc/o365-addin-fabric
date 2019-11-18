import * as React from "react";
import { CommandBar, ICommandBarItemProps, Button, ButtonType, Spinner, SpinnerType } from "office-ui-fabric-react";
import { config } from "../../config";
import { TabList, ITabListItem } from './TabList';
import { Separator } from './Separator';

export interface IAppProps {
  loaded: boolean;
  worksheet: string;
}

export interface IAppState {
  loaded: boolean;
  debug: boolean;
  sheet: string;
}

export default class App extends React.Component<IAppProps, IAppState> {

  public static getDerivedStateFromProps(nextProps: IAppProps, prevState: IAppState) {
    let result: any = {};
    if (nextProps.loaded !== prevState.loaded) result.loaded = nextProps.loaded;
    if (nextProps.worksheet !== prevState.sheet) result.sheet = nextProps.worksheet; 
    if (Object.keys(result).length === 0) return null;
    return result;
  }

  constructor(props: IAppProps, context) {

    super(props, context);

    this.state = {
      debug: false,
      loaded: false,
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

    if (!self.state.loaded) {
      return (
        <Spinner type={SpinnerType.large} label="Loading..." style={{marginTop: '45%'}} />
      );
    }

    const _items: ICommandBarItemProps[] = [
      { key: "add", text: "Add", iconProps: { iconName: "Add" }, onClick: () => console.log("Add") }
    ];

    const _farItems: ICommandBarItemProps[] = [ { key: 'info', text: 'Info', ariaLabel: 'Info', iconOnly: true, iconProps: { iconName: 'Info' }, onClick: this.aboutPage }];

    const tabs: ITabListItem[] = [

    ];

    return (
      <div>
        <CommandBar items={_items} farItems={_farItems} />
        <Separator />
        <TabList items={tabs} />
        <Separator />
        
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