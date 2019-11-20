import * as React from 'react'
import { Icon, mergeStyles } from 'office-ui-fabric-react';

export interface ITypePickerProps {
    value: string;
    onChange: (value) => void;
    color?: string;
    label?: string;
}

export const TypePicker: React.FunctionComponent<ITypePickerProps> = props => {

    const color = props.color || "#0078d4";
    const label = props.label || "Type";

    const types = [
    ];

    const labelClass = mergeStyles({ padding: '5px 0', fontSize: 14, fontWeight: 600, display: 'block' });

    const checkedContainer = mergeStyles({ display: 'inline-flex', width: 48, height: 48, background: color, border: `1px solid #fff`, cursor: 'pointer', alignItems: 'center', justifyContent: 'center', margin: '0 5px 5px 0' });
    const checkedIcon = mergeStyles({ fontSize: 32, color: '#fff' });

    const uncheckedContainer = mergeStyles({ display: 'inline-flex', width: 48, height: 48, border: `1px solid ${color}`, cursor: 'pointer', alignItems: 'center', justifyContent: 'center', margin: '0 5px 5px 0' });
    const uncheckedIcon = mergeStyles({ fontSize: 32, color: color });

    const changed = (event) => {
        if (typeof props.onChange === 'function') props.onChange(event);
    };

    const items = types.map((item) => (
        <div key={item.value} data-key={item.value} className={item.value === props.value ? checkedContainer : uncheckedContainer} title={item.name} onClick={changed}>
            <Icon iconName={item.icon} className={item.value === props.value ? checkedIcon : uncheckedIcon} />
        </div>
    ));

    return (
    <div>
        <label className={labelClass}>{label}</label>
        {items}
    </div>
    );

}