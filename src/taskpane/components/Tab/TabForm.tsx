import * as React from "react";
import { Stack, TextField } from 'office-ui-fabric-react';
import { ITabItem } from './TabItem';
import { TypePicker, IconPicker } from '../Pickers';
import { ListType } from './Types';

export interface ITabFormProps {
    item: ITabItem;
    onChange?: (field: string, value: any) => void;
}

export class TabForm extends React.Component<ITabFormProps> {

    change(field: string, event: React.FormEvent) {
        
        if (typeof this.props.onChange !== 'function') return;
        let value = null;

        switch (field) {
            case 'title':
                value = (event.target as HTMLInputElement).value;
                break;
            case 'description':
                value = (event.target as HTMLTextAreaElement).value;
                break;
            case 'icon':
                value = arguments[2].key; //value
                break;
            case 'type':
                field = 'key';
                value = (event.target as HTMLElement).closest('div').attributes['data-key'].value;
                break;
            default:
                return;
        }

        this.props.onChange(field, value);

    }

    render() {

        const readOnly = typeof this.props.onChange !== 'function';

        if (this.props.item == null) return null;

        let stackProps = {
            tokens: {
                padding: 10
            }
        };

        let typeEditor: JSX.Element = null;

        switch (this.props.item.key) {
            case 'list':
                typeEditor = <ListType item={this.props.item} />
                break;
        }

        return (
            <Stack {...stackProps}>
                <TextField label="Title" readOnly={readOnly} value={this.props.item.title} onChange={this.change.bind(this, "title")} />
                <TextField label="Description" multiline rows={3} value={this.props.item.description} onChange={this.change.bind(this, "description")} />
                <IconPicker label="Icon" value={this.props.item.icon} onChange={this.change.bind(this, "icon")} />
                <TypePicker label="Type" value={this.props.item.key} onChange={this.change.bind(this, "type")}  />
                {typeEditor}
            </Stack>
        );

    }

}