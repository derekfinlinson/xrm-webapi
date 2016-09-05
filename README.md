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

##### Create
```typescript
const account = { name: "Test Account" };

WebApi.create("accounts", account)
    .then(
        (accountId) => {
            console.log(accountId);
        },
        (error) => {
            console.log(error);
        }
    );
```
##### Retrieve
```typescript
WebApi.retrieve("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "$select=name")
    .then(
        (account) => {
            console.log(account["name"]);
        },
        (error) => {
            console.log(error);
        }
    );
```
##### Retrieve Multiple
```typescript
const options = "$filter=name eq 'Test Account'&$select=name,accountid";

WebApi.retrieveMultiple("accounts", options)
    .then(
        (results) => {
            var accounts = [];
            for (let i = 0; i < results["value"].length; i++) {
                accounts.push(results["value"][i]);
            }

            console.log(accounts.length);
        },
        (error) => {
            console.log(error);
        }
    );
```
##### Update
```typescript
const account = { name: "Updated Account" };

// Update returns no content
WebApi.update("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", account)
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );
```
##### Update Property
```typescript
// Update property returns no content
WebApi.updateProperty("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "name", "Updated Account")
    .then(() => {},
        (error) => {
            console.log(error);
        }
    );
```
##### Delete
```typescript
// Delete returns no content
WebApi.delete("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6")
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );
```
##### Delete Property
```typescript
// Delete property returns no content
WebApi.deleteProperty("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "address1_line1")
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );
```
##### Execute Custom Action
```typescript
// Custom action - Add note to account
const inputs = new Object();
inputs["title"] = "Note Title";
inputs["body"] = "Note body";

WebApi.executeAction("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "", JSON.stringify(inputs))
    .then(
        (result) => {
            console.log(result["annotationid"]);
        },
        (error) => {
            console.log(error);
        }
    );
```
### Useful Links

[Web API Reference](https://msdn.microsoft.com/en-us/library/mt593051.aspx)
