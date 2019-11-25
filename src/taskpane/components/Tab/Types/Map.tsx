import React from 'react'
import { Stack, mergeStyles } from 'office-ui-fabric-react';
import { IAppView } from '../../../../state';

export const MapType: React.FunctionComponent<{item: IAppView, onChange: (event, config) => void}> = props => {

    if (props == null || props.item == null || typeof props.onChange !== 'function') return null;

    let config = props.item.config || {};

    if (config._map == null) {
        config._map = { };
        props.onChange(event, config);
    }

    const change = function(_field: string, event: any){
        
        let propagate = false;

        if (propagate) {
            props.onChange(event, config);
        }

    };

    const officeColor = Office.context && Office.context.officeTheme ? Office.context.officeTheme.bodyBackgroundColor : '#e6e6e6';
    const separator = mergeStyles({ height: '12px', width: '100%', borderTop: '1px solid #bfbfbf', borderBottom: '1px solid #bfbfbf', backgroundColor: officeColor });

    return (
        <div>
            <div className={separator}></div>
            <Stack tokens={{ padding: 10 }}>
                <h3 className="panel-header">Map Properties</h3>
            </Stack>
            <div className={separator}></div>
        </div>
    );

}