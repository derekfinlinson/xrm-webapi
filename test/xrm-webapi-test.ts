import {WebApi} from "../src/xrm-webapi";

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

/// Demonstrate create. Update property returns no content
WebApi.updateProperty("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "name", "Updated Account")
    .then(() => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate create. Delete returns no content
WebApi.delete("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6")
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
    );

/// Demonstrate create. Delete property returns no content
WebApi.deleteProperty("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "address1_line1")
    .then(
        () => {},
        (error) => {
            console.log(error);
        }
);

/// Demonstrate create. Custom action - Add note to account
const inputs = new Object();
inputs["title"] = "Note Title";
inputs["body"] = "Note body";

WebApi.executeAction("accounts", "87989176-0887-45D1-93DA-4D5F228C10E6", "", inputs)
    .then(
        (result) => {
            console.log(result["annotationid"]);
        },
        (error) => {
            console.log(error);
        }
    );
