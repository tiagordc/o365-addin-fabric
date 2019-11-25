import React, { useContext } from 'react'
import { Dropdown, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react';
import { useStateValue } from '../../../state';

export interface IColumnPickerProps {
    label: string;
    placeholder?: string;
    value: string;
    onChange: (event, option?, index?) => void;
    color?: string;
    required?: boolean;
}

export const ColumnPicker: React.FunctionComponent<IColumnPickerProps> = props => {

    const [{ file },] = useStateValue();

    const options: IDropdownOption[] = file.currentSheet.columns.map((item) => {
        return { key: item.key, text: item.key };
    });

    if (!props.required) {
        options.splice(0, 0, { key: '', text: '' });
    }

    return <Dropdown label={props.label} placeholder={props.placeholder} options={options} defaultSelectedKey={props.value} responsiveMode={ResponsiveMode.small} required={props.required} onChange={props.onChange} />;

}