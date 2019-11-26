import React from "react";
import * as Types from '../Types';
import { Stack, TextField } from 'office-ui-fabric-react';
import { useStateValue, ActionType, IAppView } from '../../../state';
import { TypePicker, IconPicker } from '../Pickers';

export const ViewForm: React.FunctionComponent<{id: string}> = props => {

    const [{views}, dispatch] = useStateValue();
    const item = views.filter(x => x.id === props.id)[0];

    const change = function(field: string, event: React.FormEvent) {

        let value = null;

        switch (field) {
            case 'title':
                value = (event.target as HTMLInputElement).value;
                break;
            case 'description':
                value = (event.target as HTMLTextAreaElement).value;
                break;
            case 'icon':
                value = arguments[2].key; //TODO: fix this
                break;
            case 'type':
                value = (event.target as HTMLElement).closest('div').attributes['data-key'].value;
                break;
            default:
                return;
        }

        dispatch({ type: ActionType.VIEW_UPDATE, payload: { id: props.id, field, value } });

    };

    return (
        <div>
            <Stack tokens={{ padding: 10 }}>
                <h3 className="panel-header">View Properties</h3>
                <TextField label="Title" value={item.title} onChange={change.bind(this, "title")} required={true} />
                <TextField label="Description" multiline rows={3} value={item.description} onChange={change.bind(this, "description")} />
                <IconPicker label="Icon" value={item.icon} onChange={change.bind(this, "icon")} required={true} />
                <TypePicker label="Type" value={item.type} onChange={change.bind(this, "type")} required={true} />
            </Stack>
            {/* {item.type === 'calendar' && <Types.CalendarType item={item} onChange={change.bind(this, "config")} />}
            {item.type === 'cards' && <Types.CardsType item={item} onChange={change.bind(this, "config")} />}
            {item.type === 'chart' && <Types.ChartType item={item} onChange={change.bind(this, "config")} />}
            {item.type === 'gallery' && <Types.GalleryType item={item} onChange={change.bind(this, "config")} />} */}
            {item.type === 'list' && <Types.ListType id={props.id} />}
            {/* {item.type === 'map' && <Types.MapType item={item} onChange={change.bind(this, "config")} />}
            {item.type === 'table' && <Types.TableType item={item} onChange={change.bind(this, "config")} />} */}
        </div>
    );

}