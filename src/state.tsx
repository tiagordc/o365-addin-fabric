import React, { createContext, useContext, useReducer } from 'react';

const defaultProvider: [IAppState, React.Dispatch<IAction>] = null;

export interface IAppState {
  title?: string;
  debug?: string; //debug session id
  navigation?: any[];
  views?: IAppView[];
  file?: IExcelContext;
}

export enum ActionType {
  FILE_LOAD, DEBUG_START, DEBUG_MOCK,
  VIEW_ADD, VIEW_DELETE, VIEW_UPDATE
};

const reducer = (state: IAppState, action: IAction): IAppState => {
  return state;
};

export const StateContext = createContext(defaultProvider);

export const StateProvider: React.FunctionComponent<{ tools: (obj: IAppTools) => void }> = props => {

  //https://medium.com/simply/state-management-with-react-hooks-and-context-api-at-10-lines-of-code-baf6be8302c

  const initialState: IAppState = { title: 'My App', navigation: [], views: [] };
  const value = useReducer(reducer, initialState);

  props.tools({
    load: (file) => value[1]({ type: ActionType.FILE_LOAD, payload: file }),
    mock: () => value[1]({ type: ActionType.DEBUG_MOCK, payload: null })
  });

  return <StateContext.Provider value={value}>{props.children}</StateContext.Provider>;

};

export const useStateValue = () => useContext(StateContext);

export interface IAction {
  type: ActionType;
  payload: any;
}

export interface IAppTools {
  load: (file: IExcelContext) => void;
  mock: () => void;
}

export interface IAppView {
  id: string;
  sheet: string;
  order: number;
  type: string;
  title: string;
  description?: string;
  icon: string;
  config?: any;
}

export interface IExcelContext {
  currentSheet?: IExcelSheet;
}

export interface IExcelSheet {
  key?: string;
  name?: string;
  tables?: IExcelTable[];
  charts?: IExcelChart[];
  columns?: IExcelColumn[];
}

export interface IExcelTable {
  key: string;
  name?: string;
  columns?: IExcelColumn[];
}

export interface IExcelChart {
  key: string;
  name?: string;
  title?: string;
}

export interface IExcelColumn {
  key: string;
  type?: string;
  index?: number;
}
