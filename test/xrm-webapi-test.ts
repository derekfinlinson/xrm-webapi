import {
    WebApi,
    Guid,
    Entity,
    Attribute,
    FunctionInput,
    ChangeSet
} from "../src/xrm-webapi";

const api = new WebApi("8.1");

/// Demonstrate create
const account: Entity = {
    attributes: new Array < Attribute > ()
};

account.attributes.push({
    name: "name",
    value: "Test Account"
});

api.create("accounts", account)
    .then(createdAccount => {
        account.id = createdAccount.id;
    }, error => {
        console.log(error);
    });

/// Demonstrate create with returned odata
api.createWithReturnData("accounts", account, "$select=name,accountid")
    .then(account => {
        console.log(account.name);
    });

/// Demonstrate retrieve
api.retrieve("accounts", account.id, "$select=name")
    .then(account => {
        console.log(account.name);
    }, error => {
        console.log(error);
    });

/// Demonstrate retrieve multiple
const options = "$filter=name eq 'Test Account'&$select=name,accountid";

api.retrieveMultiple("accounts", options)
    .then(results => {
        var accounts = [];
        for (let i = 0; i < results.value.length; i++) {
            accounts.push(results.value[i]);
        }

        console.log(accounts.length);
    }, error => {
        console.log(error);
    });

/// Demonstrate update. Update returns no content
api.update("accounts", account.id, account)
    .then(() => {}, error => {
        console.log(error);
    });

/// Demonstrate update property. Update property returns no content
api.updateProperty("accounts", account.id, {
        name: "name",
        value: "Updated Account"
    })
    .then(() => {}, error => {
        console.log(error);
    });

/// Demonstrate delete. Delete returns no content
api.delete("accounts", account.id)
    .then(() => {}, error => {
        console.log(error);
    });

/// Demonstrate delete property. Delete property returns no content
api.deleteProperty("accounts", account.id, {
        name: "address1_line1"
    }, false)
    .then(() => {}, error => {
        console.log(error);
    });

/// Demonstrate delete navigation property. Delete property returns no content
api.deleteProperty("accounts", account.id, {
        name: "primarycontactid"
    }, true)
    .then(() => {}, error => {
        console.log(error);
    });

/// Demonstrate bound action
const inputs = new Object();
inputs["string"] = "Text";
inputs["number"] = 100;

api.boundAction("accounts", account.id, "sample_BoundAction", inputs)
    .then(result => {
        console.log(result.annotationid);
    }, error => {
        console.log(error);
    });

/// Demonstrate unbound action
const inputs2 = new Object();
inputs2["string"] = "Text";
inputs2["number"] = 100;

api.unboundAction("sample_UnboundAction", inputs2)
    .then(result => {
        console.log(result.annotationid);
    }, error => {
        console.log(error);
    });

/// Demonstrate bound function
const inputs3 = new Array < FunctionInput > ();
inputs3.push({
    name: "Argument",
    value: "Value"
});

api.boundAction("accounts", account.id, "sample_BoundFunction", inputs3)
    .then(result => {
        console.log(result.annotationid);
    }, error => {
        console.log(error);
    });

/// Demonstrate create. Custom action - Add note to account
const inputs4 = new Array < FunctionInput > ();
inputs4.push({
    name: "Target",
    value: "{'@odata.id':'accounts(87989176-0887-45D1-93DA-4D5F228C10E6)'}",
    alias: "tid"
});

api.unboundAction("sample_UnboundAction", inputs4)
    .then(result => {
        console.log(result.annotationid);
    }, error => {
        console.log(error);
    });

/// Demonstrate batch operation
const changeSets = [{
        queryString: "accounts",
        object: {
            attributes: [{
                name: "name",
                value: "Test 1"
            }]
        }
    },
    {
        queryString: "accounts",
        object: {
            attributes: [{
                name: "name",
                value: "Test 2"
            }]
        }
    },
];

const gets = [
    "accounts?$select=name"
];

api.batchOperation("BATCH123", "CHANGESET123", changeSets, gets)
    .then(result => {
        console.log(result);
    }, error => {
        console.log(error);
    });