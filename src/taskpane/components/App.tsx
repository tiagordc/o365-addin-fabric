import * as React from "react";
import { CommandBar, ICommandBarItemProps, Button, ButtonType } from "office-ui-fabric-react";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  debugging: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      debugging: false
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
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

  render() {

    let self = this;
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
        iconProps: { iconName: "Bug" },
        disabled: self.state.debugging,
        onClick: () => {
          self.setState({ debugging: true }, () => {
            const win = window as any;
            win.VORLON.Core.StartClientSide("https://localhost:1337", "default");
          });
        }
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

    return (
      <div>

        <CommandBar items={_items} />
        <div className="xls-separator"></div>


        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );

  }

}