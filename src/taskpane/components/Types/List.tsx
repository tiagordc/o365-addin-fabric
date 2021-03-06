import React from 'react'
import { Stack, Toggle, Icon, IStackStyles, IStackItemStyles, TextField, mergeStyles } from 'office-ui-fabric-react';
import { useStateValue, ActionType, IExcelColumn } from '../../../state';
import { ColumnPicker, QueryPicker, DisplayPicker, SourcePicker } from '../Pickers';

export const ListType: React.FunctionComponent<{id: string}> = props => {

    const [{ file, views }, dispatch] = useStateValue();
    const item = views.filter(x => x.id === props.id)[0];

    let config = item.config || {};

    if (config._list == null) {
        config._list = { title: null, description: null, image: null, search: true, details: false, group: false, filter: false, groupBy: null, filterBy: null, items: null, addItem: true };
        dispatch({ type: ActionType.VIEW_UPDATE, payload: { id: props.id, field: 'config', value: config } });
    }

    const change = function (field: string) {
        let argumentIndex = 1;
        if (['search', 'details', 'filter', 'addItem', 'group'].indexOf(field) >= 0) argumentIndex = 2;
        config._list[field] = arguments[argumentIndex];
        dispatch({ type: ActionType.VIEW_UPDATE, payload: { id: props.id, field: 'config', value: config } });
    };

    const officeColor = Office.context && Office.context.officeTheme ? Office.context.officeTheme.bodyBackgroundColor : '#e6e6e6';
    const separator = mergeStyles({ height: '12px', width: '100%', borderTop: '1px solid #bfbfbf', borderBottom: '1px solid #bfbfbf', backgroundColor: officeColor });
    const columnStyle: IStackStyles = { root: { overflow: 'hidden', width: '100%', borderBottom: '1px solid #bfbfbf' } };
    const columnFields: IStackItemStyles = { root: { background: '#fff', margin: 5, borderRight: '1px solid #bbb', padding: '0 10px 0 0' } };
    const columnActions: IStackItemStyles = { root: { display: 'flex', background: '#fff', justifyContent: 'center', overflow: 'hidden', width: 30, paddingTop: 5 }};
    const showSource = file.currentSheet.tables && file.currentSheet.tables.length > 0;

    let columnsSource: IExcelColumn[] = [];

    if (config._list.source){
        if (file.currentSheet.tables) {
            const sourceTable = file.currentSheet.tables.filter(x => x.key === config._list.source);
            if (sourceTable.length === 1) {
                columnsSource = sourceTable[0].columns;
            }
        }
    }
    else {
        columnsSource = file.currentSheet.columns;
    }

    const columns = columnsSource.map((item, index) => (
        <Stack key={"item_" + index} horizontal styles={columnStyle}>
            <Stack.Item grow styles={columnFields}>
                <Stack tokens={{ childrenGap: 6 }}>
                    <TextField placeholder={item.key} value={item.key} />
                    <DisplayPicker value={null} onChange={null} />
                </Stack>
            </Stack.Item>
            <Stack.Item disableShrink styles={columnActions}>
                <Stack tokens={{ childrenGap: 6 }}>
                    <Icon iconName="Move" title="Reorder" className={mergeStyles({ fontSize: 14, color: "#0078d4", cursor: 'move' })} />
                </Stack>
            </Stack.Item>
        </Stack>
    ));

    return (
        <div>
            <div className={separator}></div>
            <Stack tokens={{ padding: 10 }}>
                <h3 className="panel-header">List Properties</h3>
                {showSource && <SourcePicker label="Source" value={config._list.source} onChange={change.bind(this, 'source')} />}
                <ColumnPicker label="Title" value={config._list.title} source={config._list.source} onChange={change.bind(this, 'title')} required={true} />
                <ColumnPicker label="Description" value={config._list.description} source={config._list.source} onChange={change.bind(this, 'description')} />
                <ColumnPicker label="Image" value={config._list.image} source={config._list.source} onChange={change.bind(this, 'image')} />
                <Stack horizontal tokens={{childrenGap: 30}} style={{margin: '5px 0 0 0'}}>
                    <Toggle label="Search" checked={config._list.search} onChange={change.bind(this, 'search')}/>
                    <Toggle label="Details" checked={config._list.details} onChange={change.bind(this, 'details')} />
                    <Toggle label="Group" checked={config._list.group} onChange={change.bind(this, 'group')} />
                    <Toggle label="Filter" checked={config._list.filter} onChange={change.bind(this, 'filter')} />
                </Stack>
                {config._list.group && <ColumnPicker label="Group by" value={config._list.groupBy} source={config._list.source} onChange={change.bind(this, 'groupBy')} required={true} />}
                {config._list.filter && <div style={{ marginTop: 5}}><QueryPicker label="Filter Criteria" source={config._list.source} value={config._list.filterBy} onChange={change.bind(this, 'filterBy')} /></div>}
            </Stack>
            {config._list.details && (
            <div>
                <div className={separator}></div>
                <Stack tokens={{ padding: 10 }}>
                    <h3 className="panel-header" style={{ paddingBottom: 10, borderBottom: '1px solid #bfbfbf'}}>Item Details</h3>
                    {columns}
                    <Stack reversed horizontal style={{ marginRight: 45, marginTop: 10}}>
                        <Stack.Item align="end">
                            <Toggle label="Add"  inlineLabel checked={config._list.addItem} onChange={change.bind(this, 'addItem')} />
                        </Stack.Item>
                    </Stack>
                </Stack>
            </div>)}
            <div className={separator}></div>
        </div>
    );

}