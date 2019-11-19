import * as React from 'react'
import { ITabItem } from '../TabItem';
import { TextField, mergeStyles } from 'office-ui-fabric-react';

export interface IListTypeProps {
    item: ITabItem;
}

export const ListType: React.FunctionComponent<IListTypeProps> = _props => {
    
    return (
        <div>
            <TextField label="Test" />
        </div>
    );

}
