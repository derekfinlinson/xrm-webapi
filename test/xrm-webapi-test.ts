import {WebApi} from "../src/xrm-webapi";
import {FunctionInput} from "../src/xrm-webapi";

/// Demonstrate create
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

/// Demonstrate retrieve
WebApi.retrieve("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "$select=name")
    .then(
        (account) => {
            console.log(account["name"]);
        },
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate retrieve multiple
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

/// Demonstrate update. Update returns no content
WebApi.update("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", account)
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate update property. Update property returns no content
WebApi.updateProperty("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "name", "Updated Account")
    .then(() => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate delete. Delete returns no content
WebApi.delete("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6")
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate delete property. Delete property returns no content
WebApi.deleteProperty("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "address1_line1")
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
);

/// Demonstrate bound action
const inputs = new Object();
inputs["string"] = "Text";
inputs["number"] = 100;

WebApi.boundAction("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "sample_BoundAction", inputs)
    .then(
        (result) => {
            console.log(result["annotationid"]);
        },
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate unbound action
const inputs2 = new Object();
inputs2["string"] = "Text";
inputs2["number"] = 100;

WebApi.unboundAction("sample_UnboundAction", inputs2)
    .then(
        (result) => {
            console.log(result["annotationid"]);
        },
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate bound function
const inputs3 = new Array<FunctionInput>();
inputs3.push({name: "Argument", value: "Value"});

WebApi.boundAction("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "sample_BoundFunction", inputs3)
    .then(
        (result) => {
            console.log(result["annotationid"]);
        },
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate create. Custom action - Add note to account
const inputs4 = new Array<FunctionInput>();
inputs4.push({name: "Target", value: "{'@odata.id':'accounts(87989176-0887-45D1-93DA-4D5F228C10E6)'}", alias: "tid"});

WebApi.unboundAction("sample_UnboundAction", inputs4)
    .then(
        (result) => {
            console.log(result["annotationid"]);
        },
        (error) => {
            console.log(error);
        }
    );