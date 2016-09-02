# SiteInfo

This is a sample SharePoint Framework web part using React, Office-UI-Fabric-React and PnP-JS-Core.

![SiteInfo Client-side Web Part](SiteInfo.png)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* commonjs components - this allows this package to be reused from other packages.
* dist/* - a single bundle containing the components used for uploading to a cdn pointing a registered Sharepoint webpart library to.
* example/* a test page that hosts all components in this package.