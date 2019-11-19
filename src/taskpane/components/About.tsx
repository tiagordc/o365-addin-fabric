import "office-ui-fabric-react/dist/css/fabric.min.css";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { AppContainer } from "react-hot-loader";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { config } from "../../config";

initializeIcons();

export interface AboutState {
    clickCout: number;
}

export class About extends React.Component<{}, AboutState> {

    constructor(props, context) {

        super(props, context);

        this.state = {
            clickCout: 0
        };

    }

    componentWillMount = () => {
        this.setState({ clickCout: 0 });
    }

    debug = () => {

        const newCount = this.state.clickCout + 1;

        if (newCount > 2) {
            this.setState({clickCout: 0}, () => {

                const msg = JSON.stringify({ id: "default", url: config.vorlon });

                if (Office.context.ui) {
                    Office.context.ui.messageParent(msg);
                }
                else {
                    window.opener.localStorage.setItem('message', msg);
                    window.opener.localStorage.removeItem('message');
                }

            });
        }
        else {
            this.setState({clickCout: newCount });
        }

    }

    render() {
        return <div>
            <p>Add in version: <span onClick={this.debug}>{config.version}</span></p>
            <p>MIT License</p>
        </div>;
    }

}

ReactDOM.render(<AppContainer><About /></AppContainer>, document.getElementById("container"));
