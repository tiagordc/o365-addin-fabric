import React from 'react'
import { Dropdown, IDropdownOption, ResponsiveMode, Icon, mergeStyles } from 'office-ui-fabric-react';

export interface IDisplayPickerProps {
    label?: string;
    placeholder?: string;
    value: string;
    onChange: (event, option?, index?) => void;
}

export const DisplayPicker: React.FunctionComponent<IDisplayPickerProps> = props => {

    const color =  "#0078d4";
    const iconClass = mergeStyles({ marginRight: '8px', color: color });

    const options: IDropdownOption[] = [
        { key: '', text: 'Hidden', data: { icon: "Hide3" } },
        { key: 'label', text: 'Read only text', data: { icon: "TextOverflow" } },
        { key: 'text', text: 'Basic text', data: { icon: "TextField" } },
        { key: 'textarea', text: 'Multi-line text', data: { icon: "TextBox" } },
        { key: 'select', text: 'Drop-down list', data: { icon: "Dropdown" } },
        { key: 'date', text: 'Calendar', data: { icon: "Calendar" } },
        { key: 'check', text: 'Yes / No', data: { icon: "ToggleRight" } },
        { key: 'picture', text: 'Picture', data: { icon: "Camera" } },
        { key: 'upload', text: 'File Upload', data: { icon: "Upload" } },
        { key: 'map', text: 'Map Location', data: { icon: "MapPin" } }
    ];

    const renderOption = (option: IDropdownOption) => {
        return (
            <div>
                {option.data && option.data.icon && (
                    <Icon className={iconClass} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
                )}
                <span>{option.text}</span>
            </div>
        );
    };

    const renderTitle = (options: IDropdownOption[]) => {
        return renderOption(options[0]);
    }

    return <Dropdown label={props.label} placeholder={props.placeholder} options={options} responsiveMode={ResponsiveMode.small} defaultSelectedKey={props.value} onChange={props.onChange} onRenderOption={renderOption} onRenderTitle={renderTitle} />;

}