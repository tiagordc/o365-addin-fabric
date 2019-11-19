import * as React from "react";
import { CommandBar, ICommandBarItemProps, Spinner, SpinnerType } from "office-ui-fabric-react";
import { config } from "../../config";
import { TabList, TabForm, ITabItem } from './Tab';

export interface IAppProps {
  loaded: boolean;
  development: boolean;
  worksheet: string;
}

export interface IAppState {

  loaded: boolean;
  debug: boolean;
  sheet: string;

  menuItems?: ICommandBarItemProps[];
  tabItems?: ITabItem[];
  tabChecked?: number;

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

  addTab = () => {
    let items = this.state.tabItems || [];
    const newItem: ITabItem = { key: 'list', title: 'New Item', description: '', icon: 'List' };
    const index = items.length;
    items.push(newItem);
    this.setState({ tabItems: items, tabChecked: index });
  }

  aboutPage = () => {

    const self = this;
    const url = `${config.url}/about.html`;
    const win = window as any;

    if (Office.context.ui) {
      Office.context.ui.displayDialogAsync(url, { height: 40, width: 40 }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
            if (msg && msg.message) {
              self.setState({ debug: true });
              const debug = JSON.parse(msg.message);
              if (win.VORLON) win.VORLON.Core.StartClientSide(debug.url, debug.id);
              dialog.close();
            }
          });
        }
      });
    }
    else {
      const dialog = window.open(url);
      dialog.addEventListener('storage', (ev) => {
        if (ev.key != 'message') return; 
        self.setState({ debug: true });
        const debug = JSON.parse(ev.oldValue ? ev.oldValue : ev.newValue);
        if (win.VORLON) win.VORLON.Core.StartClientSide(debug.url, debug.id);
        dialog.close();  
      });

    }

  }

  tabChange = (index) => {
    this.setState({tabChecked: index});
  }

  tabDeleted = (index) => {
    let tabs = this.state.tabItems;
    tabs.splice(index, 1);
    this.setState({ tabItems: tabs, tabChecked: null });
  }

  tabItemChange = (field: string, value: any) => {
    const items = this.state.tabItems;
    items[this.state.tabChecked][field] = value;
    this.setState({ tabItems: items });
  }

  render() {

    let self = this;

    if (!self.state.loaded) {
      return (
        <Spinner type={SpinnerType.large} label="Loading..." style={{marginTop: '45%'}} />
      );
    }

    const info: ICommandBarItemProps[] = [ { key: 'info', text: 'Info', ariaLabel: 'Info', iconOnly: true, iconProps: { iconName: 'Info' }, onClick: this.aboutPage }];

    let innerContent: JSX.Element = null;

    if (this.state.tabChecked >= 0) {
      innerContent = <TabForm item={this.state.tabItems[this.state.tabChecked]} onChange={this.tabItemChange} />;
    }

    return (
      <div>
        <CommandBar items={this.state.menuItems} farItems={info} />
        <TabList items={this.state.tabItems} checked={this.state.tabChecked} separator={true} checkedChanged={this.tabChange} deleteTab={this.tabDeleted} />
        {innerContent}
      </div>
    );

  }

}