import * as React from "react";
import { Icon, Stack, IStackStyles, IStackItemStyles, IStackTokens, mergeStyles } from 'office-ui-fabric-react';
import Truncate from 'react-truncate';
import { ITabItem } from './TabItem';

export interface ITabListProps {

    items: ITabItem[];

    checked?: number;
    checkedChanged?: (index: number) => void;
    deleteTab?: (index: number) => void;

    separator?: boolean;
    allowSort?: boolean;

    iconWidth?: number;
    iconHeight?: number;
    iconFont?: number;
    iconBackground?: string;
    iconColor?: string;

}

export class TabList extends React.Component<ITabListProps> {

    updateDimensions() {
        this.forceUpdate();
    }

    uncheck = () => {
        if (typeof this.props.checkedChanged !== 'function') return;
        this.props.checkedChanged(null);
    }

    componentDidMount() {
        window.addEventListener("resize", this.updateDimensions.bind(this));
    }

    componentWillUnmount() {
        window.removeEventListener("resize", this.updateDimensions.bind(this));
    }

    click(index: number) {
        if (this.props.checkedChanged) {
            this.props.checkedChanged(index);
        }
    }

    delete(index: number) {
        if (this.props.deleteTab) {
            this.props.deleteTab(index);
        }
    }

    render() {

        const iconWidth = this.props.iconWidth || 70;
        const iconHeight = this.props.iconHeight || 50;
        const iconFont = this.props.iconFont || 30;
        const iconBackground = this.props.iconBackground || "#0078d4";
        const iconColor = this.props.iconColor || "#fff";
        const borderColor = '#bfbfbf';
        const officeColor = Office.context && Office.context.officeTheme ? Office.context.officeTheme.bodyBackgroundColor : '#e6e6e6';
        
        if (!this.props.items || this.props.items.length === 0) {
            const separator = mergeStyles({  height: '12px', width: '100%', borderTop: `1px solid ${borderColor}`, borderBottom: `1px solid ${borderColor}`, backgroundColor: officeColor });
            return <div className={separator}></div>;
        }

        const itemStyles: IStackStyles = { root: { overflow: 'hidden', width: '100%', borderTop: `1px solid ${borderColor}`, borderBottom: `1px solid ${borderColor}`, borderLeft: `1px solid ${iconBackground}`, borderRight: '1px solid white' } };
        const checkedStyles: IStackStyles = { root: { overflow: 'hidden', width: '100%', border: `2px solid ${iconBackground}` } };

        const iconStack: IStackItemStyles = { root: { background: iconBackground, color: iconColor, overflow: 'hidden', height: iconHeight, width: iconWidth }};
        const iconStyle = mergeStyles({ display: 'flex', width: '100%', height: '100%', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' });
        const iconClass = mergeStyles({ fontSize: iconFont });

        const textStack: IStackItemStyles = { root: { display: 'flex', background: '#fff', margin: 5, borderRight: '1px solid #bbb', paddingLeft: 10 } };
        const textStyle = mergeStyles({ display: 'flex', width: '100%', height: '100%', cursor: 'pointer' });
        const textWidth = window.innerWidth - iconWidth - 60;

        const toolsStack: IStackItemStyles = { root: { display: 'flex', alignItems: 'center', background: '#fff', justifyContent: 'center', overflow: 'hidden', width: 30 }};
        const toolsClass = mergeStyles({ fontSize: 14, color: iconBackground, cursor: 'pointer' });
        const toolsTokens: IStackTokens = { childrenGap: 6 };

        const separator = mergeStyles({ height: '2px', width: '100%', backgroundColor: officeColor });

        const listItems = this.props.items.map((item, index) => (
            <Stack key={"item_" + index} horizontal styles={index === this.props.checked ? checkedStyles : itemStyles}>
                <Stack.Item disableShrink styles={iconStack}>
                    <div className={iconStyle} onClick={this.click.bind(this, index)}>
                        <Icon iconName={item.icon} className={iconClass} />
                    </div>
                </Stack.Item>
                <Stack.Item grow styles={textStack}>
                    <div className={textStyle} onClick={this.click.bind(this, index)}>
                        <Stack>
                            <Truncate width={textWidth}>{item.title}</Truncate>
                            <Truncate width={textWidth}>{item.description}</Truncate>
                        </Stack>
                    </div>
                </Stack.Item>
                <Stack.Item disableShrink styles={toolsStack}>
                    <Stack tokens={toolsTokens}>
                        <Icon iconName="Delete" className={toolsClass} onClick={this.delete.bind(this, index)} />
                        <Icon iconName="Move" className={toolsClass} />
                    </Stack>
                </Stack.Item>
            </Stack>
        ));

        for (let i = 1; i < listItems.length; i += 2) {
            listItems.splice(i, 0, <div key={"separator_" + i} className={separator}></div>);
        }
        
        return (
            <div>
                {this.props.separator && <div onClick={this.uncheck.bind(this)} className={mergeStyles({ height: '12px', width: '100%', borderTop: `1px solid ${borderColor}`, backgroundColor: officeColor })}></div>}
                <Stack>
                    {listItems}
                </Stack>
                {this.props.separator && <div onClick={this.uncheck.bind(this)} className={ mergeStyles({ height: '12px', width: '100%', borderBottom: `1px solid ${borderColor}`, backgroundColor: officeColor })}></div>}
            </div>
        );

    }

}
