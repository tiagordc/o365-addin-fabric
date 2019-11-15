# o365-addin-fabric

## References

https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial

https://docs.microsoft.com/en-us/javascript/api/excel/excel.workbook?view=excel-js-preview

https://developer.microsoft.com/en-us/fabric#/

## Vorlon setup

First of all, open the /Vorlon/config.json file and set both useSSL property and OFFICE plugin enabled to true

On **Windows** go to /Vorlon/server/cert/ folder and:
- open the server.crt certificate and launch the install certificate wizard
- select the Local Machine certificate store
- select the Trusted Root Cetification Autorithies and ok
