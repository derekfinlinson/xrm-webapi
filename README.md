# xrm-webapi
|Build|NPM|Semantic-Release|
|-----|---|----------------|
|[![Build Status](https://derekfinlinson.visualstudio.com/GitHub/_apis/build/status/derekfinlinson.xrm-webapi)](https://derekfinlinson.visualstudio.com/GitHub/_build/latest?definitionId=2)|[![npm](https://img.shields.io/npm/v/xrm-webapi.svg?style=flat-square)](https://www.npmjs.com/package/xrm-webapi)|[![semantic-release](https://img.shields.io/badge/%20%20%F0%9F%93%A6%F0%9F%9A%80-semantic--release-e10079.svg?style=flat-square)](https://github.com/semantic-release/semantic-release)|

A Dynamics 365 Web Api TypeScript module for use in web resources or external web apps in the browser or node.

All requests return Promises. To support Internet Explorer, be sure to include a promise polyfill when deploying to CRM.

*Requires Dynamics CRM 2016 Online/On-Prem or later*

### Installation

##### Node

```
npm install --save-dev xrm-webapi
```
### Usage

#### Browser
```typescript
import { Guid, retrieve, WebApiConfig } from "xrm-webapi";

const config = new WebApiConfig("8.2");

const account = await retrieve(config, "accounts", new Guid(""), "$select=name");

console.log(account.name);
```

#### Node
```typescript
import { Guid, retrieve, WebApiConfig } from "xrm-webapi/dist/xrm-webapi-node";

const config = new WebApiConfig("8.2");

const account = await retrieve(config, "accounts", new Guid(""), "$select=name");

console.log(account.name);
```

#### Angular

For use in Angular applications, I'd first recommend using their built in [HttpClient](https://angular.io/guide/http). Besides batch operations, most D365 Web Api requests are
pretty simple to construct. If you do want to use this library, the usage is the same as the browser usage:

```typescript
import { Guid, retrieveNode, WebApiConfig } from "xrm-webapi";

const config = new WebApiConfig("8.2");

const account = await retrieve(config, "accounts", new Guid(""), "$select=name");

console.log(account.name);
```

#### Supported methods
* Retrieve
* Retrieve multiple (multiple pages)
* Retrieve multiple with Fetch XML
* Create
* Create with returned data
* Update
* Update with returned data
* Update single property
* Delete
* Delete single property
* Associate
* Disassociate
* Web API Functions
* Web API Actions
* Batch operations
* Impersonation

#### Samples
See [xrm-webapi-samples.ts](samples/xrm-webapi-samples.ts) for examples

### Useful Links

[Web API Reference](https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/perform-operations-web-api)
