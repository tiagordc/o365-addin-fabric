import * as React from "react";
import { FontIcon, Stack, IStackStyles, IStackItemStyles, IStackTokens, mergeStyles } from 'office-ui-fabric-react';
import Truncate from 'react-truncate';

export interface ITabListItem {
    icon: string;
    title: string;
    description: string;
}

export interface ITabListProps {

    items: ITabListItem[];

    iconWidth?: number;
    iconHeight?: number;
    iconFont?: number;
    iconBackground?: string;
    iconColor?: string;

}

export class TabList extends React.Component<ITabListProps> {

    updateDimensions() {
        this.setState({}); //force render
    }

    componentDidMount() {
        window.addEventListener("resize", this.updateDimensions.bind(this));
    }

    componentWillUnmount() {
        window.removeEventListener("resize", this.updateDimensions.bind(this));
    }

    render() {

        if (!this.props.items || this.props.items.length === 0) return null;

        const rootStyles: IStackStyles = { root: { overflow: 'hidden', width: '100%' } };

        const iconWidth = this.props.iconWidth || 70;
        const iconHeight = this.props.iconHeight || 50;
        const iconFont = this.props.iconFont || 30;
        const iconBackground = this.props.iconBackground || "#0078d4";
        const iconColor = this.props.iconColor || "#fff";
        const iconStack: IStackItemStyles = { root: { display: 'flex', alignItems: 'center', background: iconBackground, color: iconColor, justifyContent: 'center', overflow: 'hidden', height: iconHeight, width: iconWidth, cursor: 'pointer' }};
        const iconClass = mergeStyles({ fontSize: iconFont });

        const textStack: IStackItemStyles = { root: { display: 'flex', background: '#fff', margin: 5, borderRight: '1px solid #bbb', paddingLeft: 10, cursor: 'pointer' } };
        const textWidth = window.innerWidth - iconWidth - 60;

        const toolsStack: IStackItemStyles = { root: { display: 'flex', alignItems: 'center', background: '#fff', justifyContent: 'center', overflow: 'hidden', width: 30 }};
        const toolsClass = mergeStyles({ fontSize: 14, color: iconBackground });
        const toolsTokens: IStackTokens = { childrenGap: 6 };

        const separator = mergeStyles({ height: '1px', width: '100%', borderTop: '1px solid #bfbfbf', borderBottom: '1px solid #bfbfbf', backgroundColor: Office.context.officeTheme.bodyBackgroundColor });

        const listItems = this.props.items.map((item) => (
            <Stack horizontal styles={rootStyles}>
                <Stack.Item disableShrink styles={iconStack}>
                    <FontIcon iconName={item.icon} className={iconClass} />
                </Stack.Item>
                <Stack.Item grow styles={textStack}>
                    <Stack>
                        <Truncate width={textWidth}>{item.title}</Truncate>
                        <Truncate width={textWidth}>{item.description}</Truncate>
                    </Stack>
                </Stack.Item>
                <Stack.Item disableShrink styles={toolsStack}>
                    <Stack tokens={toolsTokens}>
                        <FontIcon iconName="Delete" className={toolsClass} />
                        <FontIcon iconName="Move" className={toolsClass} />
                    </Stack>
                </Stack.Item>
            </Stack>
        ));

        for (let i = 1; i < listItems.length; i += 2) {
            listItems.splice(i, 0, <div className={separator}></div>);
        }

        return <Stack>{listItems}</Stack>;

    }

}
