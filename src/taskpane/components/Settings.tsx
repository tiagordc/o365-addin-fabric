import React from 'react'
import { Stack, mergeStyles } from 'office-ui-fabric-react';

export interface ISettingsProps {

}

export const Settings: React.FunctionComponent<ISettingsProps> = _props => {

    const officeColor = Office.context && Office.context.officeTheme ? Office.context.officeTheme.bodyBackgroundColor : '#e6e6e6';
    const separator = mergeStyles({ height: '12px', width: '100%', borderTop: '1px solid #bfbfbf', borderBottom: '1px solid #bfbfbf', backgroundColor: officeColor });

    //licemse, theme, navigation, authentication, analytics, offline

    return (
        <div>
            <div className={separator}></div>
            <Stack tokens={{ padding: 10 }}>
                <h3 className="panel-header">App Settings</h3>
            </Stack>
            <div className={separator}></div>
        </div>
    );

}