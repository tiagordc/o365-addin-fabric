import * as React from "react";
import { mergeStyles } from 'office-ui-fabric-react';

export const Separator: React.FunctionComponent = () => {

    const merged = mergeStyles({ 
        height: '12px',
        width: '100%',
        borderTop: '1px solid #bfbfbf',
        borderBottom: '1px solid #bfbfbf',
        backgroundColor: Office.context.officeTheme.bodyBackgroundColor
    });

    return (
        <div className={merged}></div>
    );

}