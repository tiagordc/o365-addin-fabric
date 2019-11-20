import * as React from 'react'
import { Stack, Toggle, mergeStyles } from 'office-ui-fabric-react';
import { ITabItem } from '../TabItem';
import { ColumnPicker, QueryPicker } from '../../Pickers';

export interface IListTypeProps {
    item: ITabItem;
}

export const ListType: React.FunctionComponent<IListTypeProps> = _props => {

    const officeColor = Office.context && Office.context.officeTheme ? Office.context.officeTheme.bodyBackgroundColor : '#e6e6e6';
    const separatorClass = mergeStyles({ height: '12px', width: '100%', borderTop: '1px solid #bfbfbf', borderBottom: '1px solid #bfbfbf', backgroundColor: officeColor });

    let stackProps = { tokens: { padding: 10 } };

    return (
        <div>
            <div className={separatorClass}></div>
        </div>
    );

}
