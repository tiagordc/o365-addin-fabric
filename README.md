# o365-addin-fabric

## TODO

- Infer types for columns
- Reorder elements

## References

Build a message compose Outlook add-in:\
https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial

Excel JavaScript API reference:\
https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-core-concepts
https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-tables
https://docs.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-preview

UI Fabric:\
https://developer.microsoft.com/en-us/fabric#/

TypeScript:\
https://levelup.gitconnected.com/usetypescript-a-complete-guide-to-react-hooks-and-typescript-db1858d1fb9c
https://www.taniarascia.com/using-context-api-in-react/
https://medium.com/simply/state-management-with-react-hooks-and-context-api-at-10-lines-of-code-baf6be8302c

## Vorlon on Windows 

Open config.json file (%AppData%\npm\node_modules\vorlon) and set both useSSL property and OFFICE plugin enabled to true

Go to /Vorlon/server/cert/ folder and:
- open the server.crt certificate and launch the install certificate wizard
- select the Local Machine certificate store
- select the Trusted Root Cetification Autorithies and ok

## Webpack on Windows

Beginner’s guide to Webpack:\
https://medium.com/javascript-training/beginner-s-guide-to-webpack-b1f1a3638460

Resources served by local dev server:\
https://localhost:3000/webpack-dev-server

To kill webpack dev server on **Windows**:

    netstat -aon | findstr ":3000"
    taskkill /f /pid ____
