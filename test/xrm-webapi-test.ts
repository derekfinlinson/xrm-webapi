import {WebApi} from "../src/xrm-webapi";
import {Guid, Entity, Attribute, FunctionInput} from "../src/xrm-types";

/// Demonstrate create
const account: Entity = { attributes: new Array<Attribute>() };
account.attributes.push({name: "name", value: "Test Account"});

WebApi.create("accounts", account)
    .then(
        (accountId) => {
            account.id = new Guid(accountId);
        },
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate retrieve
WebApi.retrieve("accounts", account.id, "$select=name")
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
WebApi.update("accounts", account.id, account)
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate update property. Update property returns no content
WebApi.updateProperty("accounts", account.id, {name: "name", value: "Updated Account"})
    .then(() => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate delete. Delete returns no content
WebApi.delete("accounts", account.id)
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate delete property. Delete property returns no content
WebApi.deleteProperty("accounts", account.id, {name: "address1_line1"})
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

WebApi.boundAction("accounts", account.id, "sample_BoundAction", inputs)
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

WebApi.boundAction("accounts", account.id, "sample_BoundFunction", inputs3)
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