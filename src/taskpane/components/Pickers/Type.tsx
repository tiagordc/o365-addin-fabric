import * as React from 'react'
import { Icon, mergeStyles } from 'office-ui-fabric-react';

export interface ITypePickerProps {
    value: string;
    onChange: (value: string) => void;
    color?: string;
    label?: string;
}

export const TypePicker: React.FunctionComponent<ITypePickerProps> = props => {

    const color = props.color || "#0078d4";
    const label = props.label || "Type";

    const types = [

    ];

    const itemClass = mergeStyles({ display: 'inline-flex', width: 48, height: 48, border: `1px solid ${color}`, cursor: 'pointer', alignItems: 'center', justifyContent: 'center', margin: '0 5px 5px 0' });
    const iconClass = mergeStyles({ fontSize: 32, color: color });
    const labelClass = mergeStyles({ padding: '5px 0', fontSize: 14, fontWeight: 600, display: 'block' });

    const items = types.map((item) => (
        <div key={item.value} className={itemClass} title={item.name}>
            <Icon iconName={item.icon} className={iconClass} />
        </div>
    ));

    return (
    <div>
        <label className={labelClass}>{label}</label>
        {items}
    </div>
    );

}