import React, { useEffect } from 'react'
import { Dropdown, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react';
import { useStateValue } from '../../../state';

export interface IColumnPickerProps {
    label: string;
    required?: boolean;
    placeholder?: string;
    value: string;
    source: string;
    onChange: (column: string) => void;
}

export const ColumnPicker: React.FunctionComponent<IColumnPickerProps> = props => {

    const [{ file },] = useStateValue();
    const sheet = file.currentSheet;

    useEffect(() => {
        if (props.required && !props.value && options.length > 0) {
            change(null, options[0]);
        }
    }, []);

    let options: IDropdownOption[];

    if (props.source == null || props.source === '') {
        options = sheet.columns.map((item) => ({ key: item.key, text: item.key }));
    }
    else {
        const tables = sheet.tables.filter(x => x.key === props.source);
        if (tables.length === 1) {
            options = tables[0].columns.map((item) => ({ key: item.key, text: item.key }));
        }
    }

    if (!props.required) {
        options.splice(0, 0, { key: '', text: '' });
    }

    const change = (_e, option?) => {
        if (typeof props.onChange === 'function') {
            if (option && option.key) props.onChange(option.key);
            else props.onChange(null);
        }
    };

    return <Dropdown label={props.label} placeholder={props.placeholder} options={options} defaultSelectedKey={props.value} responsiveMode={ResponsiveMode.small} required={props.required} onChange={change} />;

}