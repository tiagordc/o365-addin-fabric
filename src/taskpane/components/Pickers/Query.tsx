import * as React from 'react'
import { QueryBuilderComponent } from '@syncfusion/ej2-react-querybuilder';
import { mergeStyles } from 'office-ui-fabric-react';

export interface IQueryPickerProps {
    label: string;
    placeholder?: string;
    value: any;
    onChange: (value) => void;
}

export const QueryPicker: React.FunctionComponent<IQueryPickerProps> = props => {

    let columnData = [
        { field: 'EmployeeID', label: 'EmployeeID', type: 'number' },
        { field: 'FirstName', label: 'FirstName', type: 'string' },
        { field: 'TitleOfCourtesy', label: 'Title Of Courtesy', type: 'boolean', values: ['Mr.', 'Mrs.'] },
        { field: 'Title', label: 'Title', type: 'string' },
        { field: 'HireDate', label: 'HireDate', type: 'date', format: 'dd/MM/yyyy' },
        { field: 'Country', label: 'Country', type: 'string' },
        { field: 'City', label: 'City', type: 'string' }
    ];

    const labelClass = mergeStyles({ padding: '5px 0', fontSize: 14, fontWeight: 600, display: 'block' });

    return (
        <div>
            {props.label && <label className={labelClass}>{props.label}</label>}
            <QueryBuilderComponent width='100%' columns={columnData} rule={props.value} />
        </div>
    );

}