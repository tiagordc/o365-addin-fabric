import * as React from 'react'
import { Dropdown, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react';

export interface IColumnPickerProps {
    label: string;
    placeholder?: string;
    value: string;
    onChange: (value) => void;
    color?: string;
    required?: boolean;
}

export const ColumnPicker: React.FunctionComponent<IColumnPickerProps> = props => {

    const options: IDropdownOption[] = [
        { key: 'id', text: 'ID' },
        { key: 'company', text: 'Company' },
        { key: 'address', text: 'Address' },
        { key: 'owner', text: 'Owner' }
    ];

    return <Dropdown label={props.label} placeholder={props.placeholder} options={options} responsiveMode={ResponsiveMode.small} required={props.required} />;

}