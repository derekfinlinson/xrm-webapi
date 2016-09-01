# xrm-webapi

A Dynamics CRM Web API TypeScript module for use in Web Resources.

All methods return a generic Promise. The module depends on es6-promise to add Promise support for Internet Explorer but any promise polyfill may be used when deploying to CRM.

*Requires Dynamics CRM 2016 Online/On-Prem or later*

### Installation

##### Node

```
npm install --save-dev xrm-webapi
```

### Usage

Import the module into your TypeScript files

```typescript
import {WebApi} from "../node_modules/xrm-webapi/dist/xrm-webapi";
```

##### Supports the following Web API operations

* Create
* Retrieve
* Retrieve Multiple
* Update
* Update Property
* Delete
* Delete Property
* Action

### Useful Links

[Web API Reference](https://msdn.microsoft.com/en-us/library/mt593051.aspx)
