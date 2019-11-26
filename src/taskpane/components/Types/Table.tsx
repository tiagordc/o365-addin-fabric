import React from 'react'
import { Stack, mergeStyles } from 'office-ui-fabric-react';
import { useStateValue, ActionType } from '../../../state';

export const TableType: React.FunctionComponent<{id: string}> = props => {

    const [{ views }, dispatch] = useStateValue();
    const item = views.filter(x => x.id === props.id)[0];

    let config = item.config || {};

    if (config._table == null) {
        config._table = {  };
        dispatch({ type: ActionType.VIEW_UPDATE, payload: { id: props.id, field: 'config', value: config } });
    }

    const change = function(_field: string, _event: any){
        
    };

    const officeColor = Office.context && Office.context.officeTheme ? Office.context.officeTheme.bodyBackgroundColor : '#e6e6e6';
    const separator = mergeStyles({ height: '12px', width: '100%', borderTop: '1px solid #bfbfbf', borderBottom: '1px solid #bfbfbf', backgroundColor: officeColor });

    return (
        <div>
            <div className={separator}></div>
            <Stack tokens={{ padding: 10 }}>
                <h3 className="panel-header">Table Properties</h3>
            </Stack>
            <div className={separator}></div>
        </div>
    );

}