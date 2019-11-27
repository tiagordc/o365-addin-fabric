import React, { useRef, useEffect } from 'react'
import { QueryBuilderComponent, ColumnsModel } from '@syncfusion/ej2-react-querybuilder';
import { mergeStyles } from 'office-ui-fabric-react';
import { useStateValue } from '../../../state';

export interface IQueryPickerProps {
    label: string;
    placeholder?: string;
    value: string;
    source: string;
    onChange: (value) => void;
}

export const QueryPicker: React.FunctionComponent<IQueryPickerProps> = props => {

    const [{ file },] = useStateValue();
    const queryElement = useRef(null);

    useEffect(() => {
        try {
            if (props.value) {
                queryElement.current.setRulesFromSql(props.value);
            }
        }
        catch (ex) { }
    }, []);

    let columnData: ColumnsModel[];
    
    //https://ej2.syncfusion.com/react/documentation/query-builder/
    //https://ej2.syncfusion.com/react/documentation/query-builder/columns/

    if (props.source) {
        if (file.currentSheet.tables) {
            const sourceTable = file.currentSheet.tables.filter(x => x.key === props.source);
            if (sourceTable.length === 1) {
                columnData = sourceTable[0].columns.map(item => ({ field: item.key, label: item.key, type: item.type ? item.type : 'string' }));
            }
        }
    }
    else {
        columnData = file.currentSheet.columns.map(item => {
            return { field: item.key, label: item.key, type: item.type ? item.type : 'string' };
        });
    }

    const labelClass = mergeStyles({ padding: '5px 0', fontSize: 14, fontWeight: 600, display: 'block' });
    
    const change = () => {
        const rules = queryElement.current.getRules();
        let sql = queryElement.current.getSqlFromRules(rules);
        if (sql === '') sql = null;
        if (typeof props.onChange === 'function' && sql != props.value) {
            if (sql) props.onChange(sql);
            else props.onChange(null);
        }
    };

    return (
        <div>
            {props.label && <label className={labelClass}>{props.label}</label>}
            <QueryBuilderComponent ref={queryElement} width='100%' columns={columnData} change={change} />
        </div>
    );

}