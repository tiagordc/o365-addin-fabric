import React, { useState, useEffect } from "react";
import { Icon, Stack, IStackStyles, IStackItemStyles, mergeStyles } from 'office-ui-fabric-react';
import Truncate from 'react-truncate';
import { IAppView } from '../../../state';

export interface IViewListProps {

    items: IAppView[];

    checked?: string;
    checkedChanged?: (id) => void;
    deleteTab?: (id) => void;

    separator?: boolean;
    allowSort?: boolean;

    iconWidth?: number;
    iconHeight?: number;
    iconFont?: number;
    iconBackground?: string;
    iconColor?: string;

}

export const ViewList: React.FunctionComponent<IViewListProps> = props => {

    const [textWidth, setTextWidth] = useState(0);

    const iconWidth = props.iconWidth || 70;
    const iconHeight = props.iconHeight || 50;
    const iconFont = props.iconFont || 30;
    const iconBackground = props.iconBackground || "#0078d4";
    const iconColor = props.iconColor || "#fff";
    const borderColor = '#bfbfbf';
    const officeColor = Office.context && Office.context.officeTheme ? Office.context.officeTheme.bodyBackgroundColor : '#e6e6e6';

    const updateDimensions = () => {
        let width = window.innerWidth - iconWidth - 67;
        if (width <= 10) width = 10;
        setTextWidth(width);
    };

    useEffect(() => {
        updateDimensions();
        window.addEventListener('resize', updateDimensions);
        return () => {
            window.removeEventListener('resize', updateDimensions);
        };
    }, []);

    const uncheck = () => {
        if (typeof props.checkedChanged !== 'function') return;
        props.checkedChanged(null);
    };

    const click = (id: string) => {
        if (typeof props.checkedChanged === 'function') {
            props.checkedChanged(id);
        }
    };

    const remove = (id: string) => {
        if (typeof props.deleteTab === 'function') {
            props.deleteTab(id);
        }
    };

    if (!props.items || props.items.length === 0) {
        const separator = mergeStyles({ height: '12px', width: '100%', borderTop: `1px solid ${borderColor}`, borderBottom: `1px solid ${borderColor}`, backgroundColor: officeColor });
        return <div className={separator}></div>;
    }

    const normalTab: IStackStyles = { root: { overflow: 'hidden', width: '100%', borderTop: `1px solid ${borderColor}`, borderBottom: `1px solid ${borderColor}`, borderLeft: `1px solid ${iconBackground}`, borderRight: '1px solid white' } };
    const checkedTab: IStackStyles = { root: { overflow: 'hidden', width: '100%', border: `2px solid ${iconBackground}` } };

    const iconContainer: IStackItemStyles = { root: { background: iconBackground, color: iconColor, overflow: 'hidden', height: iconHeight, width: iconWidth } };
    const iconStyle = mergeStyles({ display: 'flex', width: '100%', height: '100%', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' });

    const textContainer: IStackItemStyles = { root: { display: 'flex', background: '#fff', margin: 5, borderRight: '1px solid #bbb', paddingLeft: 10 } };
    const textStyle = mergeStyles({ display: 'flex', width: '100%', height: '100%', cursor: 'pointer' });

    const actionContainer: IStackItemStyles = { root: { display: 'flex', alignItems: 'center', background: '#fff', justifyContent: 'center', overflow: 'hidden', width: 30 } };
    const separator = mergeStyles({ height: '2px', width: '100%', backgroundColor: officeColor });

    const listItems = props.items.map((item, index) => (
        <Stack key={"item_" + index} horizontal styles={item.id === props.checked ? checkedTab : normalTab}>
            <Stack.Item disableShrink styles={iconContainer}>
                <div className={iconStyle} onClick={click.bind(this, item.id)}>
                    <Icon iconName={item.icon} className={mergeStyles({ fontSize: iconFont })} />
                </div>
            </Stack.Item>
            <Stack.Item grow styles={textContainer}>
                <div className={textStyle} onClick={click.bind(this, item.id)}>
                    <Stack>
                        <Truncate width={textWidth} lines={1}>{item.title}</Truncate>
                        <Truncate width={textWidth} lines={1}>{item.description}</Truncate>
                    </Stack>
                </div>
            </Stack.Item>
            <Stack.Item disableShrink styles={actionContainer}>
                <Stack tokens={{ childrenGap: 6 }}>
                    <Icon iconName="Delete" className={mergeStyles({ fontSize: 14, color: iconBackground, cursor: 'pointer' })} onClick={remove.bind(this, item.id)} />
                    <Icon iconName="Move" className={mergeStyles({ fontSize: 14, color: iconBackground, cursor: 'move' })} />
                </Stack>
            </Stack.Item>
        </Stack>
    ));

    for (let i = 1; i < listItems.length; i += 2) {
        listItems.splice(i, 0, <div key={"separator_" + i} className={separator}></div>);
    }

    return (
        <div>
            {props.separator && <div onClick={uncheck.bind(this)} className={mergeStyles({ height: '12px', width: '100%', borderTop: `1px solid ${borderColor}`, backgroundColor: officeColor })}></div>}
            <Stack>
                {listItems}
            </Stack>
            {props.separator && <div onClick={uncheck.bind(this)} className={mergeStyles({ height: '12px', width: '100%', borderBottom: `1px solid ${borderColor}`, backgroundColor: officeColor })}></div>}
        </div>
    );

}
