# o365-addin-fabric

## TODO

- Reorder elements
- Item details

## References

Build a message compose Outlook add-in:\
https://docs.microsoft.com/en-us/outlook/add-ins/addin-tutorial

## Vorlon

Open config.json file (%AppData%\npm\node_modules\vorlon) and set both useSSL property and OFFICE plugin enabled to true

Go to /Vorlon/server/cert/ folder and:
- open the server.crt certificate and launch the install certificate wizard
- select the Local Machine certificate store
- select the Trusted Root Cetification Autorithies and ok

## Webpack

Beginner’s guide to Webpack:\
https://medium.com/javascript-training/beginner-s-guide-to-webpack-b1f1a3638460

Resources served by local dev server:\
https://localhost:3000/webpack-dev-server

To kill webpack dev server:
    netstat -aon | findstr ":3000"
    taskkill /f /pid ____
