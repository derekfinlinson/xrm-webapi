# xrm-webapi
[![Travis branch](https://img.shields.io/travis/rust-lang/rust/master.svg?style=flat-square)](https://travis-ci.org/derekfinlinson/xrm-webapi)
[![npm](https://img.shields.io/npm/v/xrm-webapi.svg?style=flat-square)](https://www.npmjs.com/package/xrm-webapi)
[![semantic-release](https://img.shields.io/badge/%20%20%F0%9F%93%A6%F0%9F%9A%80-semantic--release-e10079.svg?style=flat-square)](https://github.com/semantic-release/semantic-release)
[![Gitter](https://img.shields.io/gitter/room/nwjs/nw.js.svg?style=flat-square)](https://gitter.im/xrm-webapi/Lobby)

A Dynamics CRM Web API TypeScript module for use in Web Resources.

All methods return a generic Promise. The module depends on [es6-promise](https://github.com/stefanpenner/es6-promise) to add Promise support for Internet Explorer but any promise polyfill may be used when deploying to CRM.

*Requires Dynamics CRM 2016 Online/On-Prem or later*

### Installation

##### Node

```
npm install --save-dev xrm-webapi
```
### Usage

Import the module into your TypeScript files

```typescript
import {WebApi} from "xrm-webapi";

const api = new WebApi("8.1");
```

### Samples
See [xrm-webapi-test.ts](test/xrm-webapi-test.ts) for samples

### Useful Links

[Web API Reference](https://msdn.microsoft.com/en-us/library/mt593051.aspx)
