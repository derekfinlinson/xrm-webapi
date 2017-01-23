# xrm-webapi
[![Build Status](https://travis-ci.org/derekfinlinson/xrm-webapi.png?branch=master)](https://travis-ci.org/derekfinlinson/xrm-webapi)
[![Join the chat at https://gitter.im/xrm-webapi/Lobby](https://badges.gitter.im/xrm-webapi/Lobby.svg)](https://gitter.im/xrm-webapi/Lobby?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

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
import {WebApi} from "../node_modules/xrm-webapi/src/xrm-webapi";

const api = new WebApi("8.1");
```

### Samples
See [xrm-webapi-test.ts](test/xrm-webapi-test.ts) for samples

### Whats New
* 0.6.4
    * Add support for deleting navigation properties
* 0.6.2 - BREAKING CHANGE
    * WebApi class object must now be instantiated
    * Allows to enter in WebApi version to use
    * Exposes getClientUrl method for use with batch operations
* 0.6.1
    * Support for batch operations

### Useful Links

[Web API Reference](https://msdn.microsoft.com/en-us/library/mt593051.aspx)
