import React, { useEffect } from 'react'
import { Dropdown, IDropdownOption, ResponsiveMode } from 'office-ui-fabric-react';
import { useStateValue } from '../../../state';

export interface IColumnPickerProps {
    label: string;
    required?: boolean;
    placeholder?: string;
    value: string | { name: string, table?: string };
    onChange: (column: string | { name: string, table?: string }) => void;
}

export const ColumnPicker: React.FunctionComponent<IColumnPickerProps> = props => {

    const [{ file },] = useStateValue();
    const sheet = file.currentSheet;

    useEffect(() => {
        if (props.required && !stringValue && options.length > 0) {
            change(null, options[0]);
        }
    }, []);

    let options: IDropdownOption[] = sheet.columns.map((item) => {
        return { key: btoa(`\|/${item.key}`), text: item.key };
    });

    if (sheet.tables && sheet.tables.length > 0) {
        for (let i = 0; i < sheet.tables.length; i++) {
            const table = sheet.tables[i];
            options.push(...table.columns.map((item) => {
                return { key: btoa(`${table.name}\|/${item.key}`), text: item.key, data: { table } };
            }));
        }
    }
    
    if (!props.required) {
        options.splice(0, 0, { key: '', text: '' });
    }

    const renderOption = (option: IDropdownOption) => {
        return (
            <div>
                {option.key && (!option.data || !option.data.table) && <span>{sheet.name}: </span>}
                {option.data && option.data.table && <span>{option.data.table.name}: </span>}
                <span>{option.text}</span>
            </div>
        );
    };

    const renderTitle = (options: IDropdownOption[]) => {
        return renderOption(options[0]);
    }

    let stringValue: string;

    if (props.value) {
        if (typeof props.value === 'string') stringValue = btoa(`\|/${props.value}`);
        else if (typeof props.value === 'object' && props.value.name) stringValue = btoa(`${props.value.table ? props.value.table : ''}\|/${props.value.name}`);
    }

    const change = (_e, option?) => {
        if (typeof props.onChange === 'function') {
            if (option && option.key) {
                const parts = atob(option.key).split('\|/');
                if (parts[0].length === 0) props.onChange(parts[1]);
                else props.onChange( { name: parts[1], table: parts[0 ]});
            }
            else props.onChange(null);
        }
    };

    return <Dropdown label={props.label} placeholder={props.placeholder} options={options} defaultSelectedKey={stringValue} responsiveMode={ResponsiveMode.small} required={props.required} onChange={change} onRenderOption={renderOption} onRenderTitle={renderTitle} />;

}