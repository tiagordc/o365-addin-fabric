import React, { useEffect } from 'react'
import { Dropdown, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react';
import { useStateValue } from '../../../state';

export interface ISourcePickerProps {
    label: string;
    placeholder?: string;
    value: string;
    onChange: (column: string | { name: string, table?: string }) => void;
}

export const SourcePicker: React.FunctionComponent<ISourcePickerProps> = props => {

    const [{ file },] = useStateValue();
    const sheet = file.currentSheet;

    let options: IDropdownOption[] = [{ key: '', text: sheet.name }];

    if (sheet.tables) {
        sheet.tables.forEach(table => {
            options.push({ key: table.key, text: table.name });
        });
    }
    
    const change = (_e, option?) => {
        if (typeof props.onChange === 'function') {
            if (option && option.key) props.onChange(option.key);
            else props.onChange(null);
        }
    };

    let stringValue = props.value;
    if (stringValue == null) stringValue = '';

    return <Dropdown label={props.label} placeholder={props.placeholder} options={options} defaultSelectedKey={stringValue} responsiveMode={ResponsiveMode.small} required={true} onChange={change} />;

}