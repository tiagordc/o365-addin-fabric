import React, { useState } from "react";
import { CommandBar, ICommandBarItemProps, Spinner, SpinnerType } from "office-ui-fabric-react";
import { config } from "../../config";
import { TabList, TabForm } from './Tab';
import { Settings } from './Settings';
import { useStateValue, ActionType, IAppView } from '../../state'
      
export const App: React.FunctionComponent = () => {

  const [{ views, file }, dispatch] = useStateValue();
  const [activeMenu, setActiveMenu] = useState('views');
  const [activeView, setActiveView] = useState('');
  
  const addView = () => {
    const newItem: IAppView = { id: config.id(), sheet: file.currentSheet.key, order: 1, type: 'list', title: 'New Item', description: '', icon: 'List' };
    dispatch({ type: ActionType.VIEW_ADD, payload: newItem });
    setActiveView(newItem.id);
  }

  const viewChanged = (field: string, value: any) => {
    dispatch({ type: ActionType.VIEW_UPDATE, payload: { id: activeView, field, value }})
  }

  const viewDeleted = (id: string) => {
    dispatch({ type: ActionType.VIEW_DELETE, payload: id });
    setActiveView(null);
  }

  const openMenu = (page: string) => setActiveMenu(page);

  const aboutPage = () => {

    const url = `${config.url}/about.html`;
    const win = window as any;
    let debugInfo: any = null;

    if (Office.context.ui) {
      Office.context.ui.displayDialogAsync(url, { height: 40, width: 40 }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (msg) => {
            if (msg && msg.message) {
              dispatch({ type: ActionType.DEBUG_START, payload: debugInfo.id });
              debugInfo = JSON.parse(msg.message);
              if (win.VORLON) win.VORLON.Core.StartClientSide(debugInfo.url, debugInfo.id);
              dialog.close();
            }
          });
        }
      });
    }
    else {
      const dialog = win.open(url);
      dialog.addEventListener('storage', (ev) => {
        if (ev.key != 'message') return;
        debugInfo = JSON.parse(ev.oldValue ? ev.oldValue : ev.newValue);
        dispatch({ type: ActionType.DEBUG_START, payload: debugInfo.id });
        if (win.VORLON) win.VORLON.Core.StartClientSide(debugInfo.url, debugInfo.id);
        dialog.close();
      });

    }

  };

  if (!file || !file.currentSheet) return <Spinner type={SpinnerType.large} label="Loading..." style={{ marginTop: '45%' }} />;
  
  const info: ICommandBarItemProps[] = [{ key: 'info', text: 'Info', ariaLabel: 'Info', iconOnly: true, iconProps: { iconName: 'Info' }, onClick: aboutPage }];
  const sheetTabs = views.filter(item => item.sheet === file.currentSheet.key);
  let tabItem = activeView ? sheetTabs.filter(x => x.id === activeView)[0] : null;

  let menus:  ICommandBarItemProps[] = [ { key: "back", text: "Back", iconProps: { iconName: "ChevronLeft" }, onClick: () => openMenu('views') } ];
  if (activeMenu === 'views') {
    menus = [
      { key: "add", text: "Add View", iconProps: { iconName: "Add" }, onClick: () => addView() },
      { key: "settings", text: "Settings", iconProps: { iconName: "CellPhone" }, onClick: () => openMenu('settings') },
      { key: "preview", text: "Preview", iconProps: { iconName: "RedEye" }, onClick: () => console.log("Preview") },
      { key: "save", text: "Deploy", iconProps: { iconName: "WebPublish" }, onClick: () => console.log("Save") }
    ];
  }

  return (
    <div>
      <CommandBar items={menus} farItems={info} />
      {activeMenu === 'views' && <TabList items={views} checked={activeView} separator={true} checkedChanged={(id) => setActiveView(id)} deleteTab={viewDeleted} />}
      {activeMenu === 'views' && activeView && (<TabForm item={tabItem} onChange={viewChanged} />)}
      {activeMenu === 'settings' && <Settings />}
    </div>
  );

}