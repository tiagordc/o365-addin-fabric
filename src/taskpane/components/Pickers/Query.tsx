import React, { useContext } from 'react'
import { QueryBuilderComponent, ColumnsModel } from '@syncfusion/ej2-react-querybuilder';
import { mergeStyles } from 'office-ui-fabric-react';
import { useStateValue } from '../../../state';

export interface IQueryPickerProps {
    label: string;
    placeholder?: string;
    value: any;
    onChange: (value) => void;
}

export const QueryPicker: React.FunctionComponent<IQueryPickerProps> = props => {

    const [{ file },] = useStateValue();

    let columnData: ColumnsModel[] = file.currentSheet.columns.map(item => {
        return { field: item.key, label: item.key, type: item.type };
    });

    const labelClass = mergeStyles({ padding: '5px 0', fontSize: 14, fontWeight: 600, display: 'block' });

    return (
        <div>
            {props.label && <label className={labelClass}>{props.label}</label>}
            <QueryBuilderComponent width='100%' columns={columnData} rule={props.value} />
        </div>
    );

}