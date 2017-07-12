// tslint:disable:no-console
// tslint:disable:no-empty

import {
    ChangeSet,
    FunctionInput,
    Guid,
    WebApi,
} from "../src/xrm-webapi";

const api = new WebApi("8.1");

/// Demonstrate create
const account = {
    name: "Test Account"
};

api.create("accounts", account)
    .then((createdAccount) => {
        console.log(createdAccount.id);
    }, (error) => {
        console.log(error);
    });

/// Demonstrate create with returned odata
api.createWithReturnData("accounts", account, "$select=name,accountid")
    .then((created: any) => {
        console.log(created.name);
    });

/// Demonstrate retrieve
api.retrieve("accounts", new Guid(""), "$select=name")
    .then((retrieved) => {
        console.log(retrieved.name);
    }, (error) => {
        console.log(error);
    });

/// Demonstrate retrieve multiple
const options = "$filter=name eq 'Test Account'&$select=name,accountid";

api.retrieveMultiple("accounts", options)
    .then((results) => {
        const accounts = [];
        for (let record of results.value) {
            accounts.push(record);
        }

        console.log(accounts.length);
    }, (error) => {
        console.log(error);
    });

/// Demonstrate update. Update returns no content
api.update("accounts", new Guid(""), account)
    .then(() => {}, (error) => {
        console.log(error);
    });

/// Demonstrate update property. Update property returns no content
api.updateProperty("accounts", new Guid(""), "name", "Updated Account")
    .then(() => {}, (error) => {
        console.log(error);
    });

/// Demonstrate delete. Delete returns no content
api.delete("accounts", new Guid(""))
    .then(() => {}, (error) => {
        console.log(error);
    });

/// Demonstrate delete property. Delete property returns no content
api.deleteProperty("accounts", new Guid(""), "address1_line1", false)
    .then(() => {}, (error) => {
        console.log(error);
    });

/// Demonstrate delete navigation property. Delete property returns no content
api.deleteProperty("accounts", new Guid(""), "primarycontactid", true)
    .then(() => {}, (error) => {
        console.log(error);
    });

/// Demonstrate associate. Associate returns no content
api.associate("accounts", new Guid(""), "contact_customer_accounts", "contacts", new Guid(""))
    .then(() => {}, (error) => {
        console.log(error);
    });

/// Demonstrate disassociate. Disassociate returns no content
api.disassociate("accounts", new Guid(""), "contact_customer_accounts", "contacts")
    .then(() => {}, (error) => {
        console.log(error);
    });

/// Demonstrate bound action
const inputs = {
    NumberInput: 100,
    StringInput: "Text",
};

api.boundAction("accounts", new Guid(""), "sample_BoundAction", inputs)
    .then((result) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

api.unboundAction("sample_UnboundAction", inputs)
    .then((result) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

/// Demonstrate bound function
const inputs3: FunctionInput[] = [];

inputs3.push({
    name: "Argument",
    value: "Value",
});

api.boundAction("accounts", new Guid(""), "sample_BoundFunction", inputs3)
    .then((result) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

/// Demonstrate create. Custom action - Add note to account
const inputs4: FunctionInput[] = [];

inputs4.push({
    alias: "tid",
    name: "Target",
    value: "{'@odata.id':'accounts(87989176-0887-45D1-93DA-4D5F228C10E6)'}",
});

api.unboundAction("sample_UnboundAction", inputs4)
    .then((result) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

/// Demonstrate batch operation
const changeSets = [
    {
        entity: {
            name: "Test 1"
        },
        queryString: "accounts",
    },
    {
        entity: {
            name: "Test 2"
        },
        queryString: "accounts",
    },
];

const gets = [
    "accounts?$select=name",
];

api.batchOperation("BATCH123", "CHANGESET123", changeSets, gets)
    .then((result) => {
        console.log(result);
    }, (error) => {
        console.log(error);
    });
