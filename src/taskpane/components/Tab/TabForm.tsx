import React from "react";
import * as Types from './Types';
import { Stack, TextField, mergeStyles } from 'office-ui-fabric-react';
import { IAppView } from '../../../state';
import { TypePicker, IconPicker } from '../Pickers';


export interface ITabFormProps {
    item: IAppView;
    onChange?: (field: string, value: any) => void;
}

export const TabForm: React.FunctionComponent<ITabFormProps> = props => {

    const change = function(field: string, event: React.FormEvent) {

        if (typeof props.onChange !== 'function') return;
        let value = null;

        switch (field) {
            case 'title':
                value = (event.target as HTMLInputElement).value;
                break;
            case 'description':
                value = (event.target as HTMLTextAreaElement).value;
                break;
            case 'icon':
                value = arguments[2].key;
                break;
            case 'type':
                value = (event.target as HTMLElement).closest('div').attributes['data-key'].value;
                break;
            case 'config':
                value = arguments[2];
                break;
            default:
                return;
        }

        props.onChange(field, value);

    };

    const readOnly = typeof props.onChange !== 'function';
    if (props.item == null) return null;

    return (
        <div>
            <Stack tokens={{ padding: 10 }}>
                <h3 className="panel-header">View Properties</h3>
                <TextField label="Title" readOnly={readOnly} value={props.item.title} onChange={change.bind(this, "title")} required={true} />
                <TextField label="Description" multiline rows={3} value={props.item.description} onChange={change.bind(this, "description")} />
                <IconPicker label="Icon" value={props.item.icon} onChange={change.bind(this, "icon")} required={true} />
                <TypePicker label="Type" value={props.item.type} onChange={change.bind(this, "type")} required={true} />
            </Stack>
            {props.item.type === 'calendar' && <Types.CalendarType item={props.item} onChange={change.bind(this, "config")} />}
            {props.item.type === 'cards' && <Types.CardsType item={props.item} onChange={change.bind(this, "config")} />}
            {props.item.type === 'chart' && <Types.ChartType item={props.item} onChange={change.bind(this, "config")} />}
            {props.item.type === 'gallery' && <Types.GalleryType item={props.item} onChange={change.bind(this, "config")} />}
            {props.item.type === 'list' && <Types.ListType item={props.item} onChange={change.bind(this, "config")} />}
            {props.item.type === 'map' && <Types.MapType item={props.item} onChange={change.bind(this, "config")} />}
            {props.item.type === 'table' && <Types.TableType item={props.item} onChange={change.bind(this, "config")} />}
        </div>
    );

}