# o365-addin-fabric

## References

Build a message compose Outlook add-in:\
https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial

Excel JavaScript API reference:\
https://docs.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-preview

UI Fabric:\
https://developer.microsoft.com/en-us/fabric#/

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

    netstat -aon | findstr 3000
    taskkill /pid ____
