// tslint:disable:no-console
// tslint:disable:no-empty

import {
    ChangeSet,
    FunctionInput,
    Guid,
    WebApi,
    WebApiConfig
} from "../src/xrm-webapi";

const apiConfig: WebApiConfig = {
    version: "8.1"
};

const api: WebApi = new WebApi(apiConfig);

// demonstrate create
const account: any = {
    name: "Test Account"
};

api.create("accounts", account)
    .then(() => {
        console.log();
    }, (error) => {
        console.log(error);
    });

// demonstrate create with returned odata
api.createWithReturnData("accounts", account, "$select=name,accountid")
    .then((created: any) => {
        console.log(created.name);
    });

// demonstrate retrieve
api.retrieve("accounts", new Guid(""), "$select=name")
    .then((retrieved) => {
        console.log(retrieved.data.name);
    }, (error) => {
        console.log(error);
    });

// demonstrate retrieve multiple
const options: string = "$filter=name eq 'Test Account'&$select=name,accountid";

api.retrieveMultiple("accounts", options)
    .then(
        (results) => {
            const accounts: any[] = [];
            for (let record of results.value) {
                accounts.push(record);
            }

            // demonstrate getting next page from retreiveMultiple
            api.getNextPage(results["@odata.nextlink"]).then(
                (moreResults) => {
                    console.log(moreResults.value.length);
                }
            );

            console.log(accounts.length);
        },
        (error) => {
            console.log(error);
        }
    );

// demonstrate update. Update returns no content
api.update("accounts", new Guid(""), account)
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate update property. Update property returns no content
api.updateProperty("accounts", new Guid(""), "name", "Updated Account")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate delete. Delete returns no content
api.delete("accounts", new Guid(""))
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate delete property. Delete property returns no content
api.deleteProperty("accounts", new Guid(""), "address1_line1")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate delete navigation property. Delete property returns no content
api.deleteProperty("accounts", new Guid(""), "primarycontactid")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate associate. Associate returns no content
api.associate("accounts", new Guid(""), "contact_customer_accounts", "contacts", new Guid(""))
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate disassociate. Disassociate returns no content
api.disassociate("accounts", new Guid(""), "contact_customer_accounts")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate bound action
const inputs: object = {
    NumberInput: 100,
    StringInput: "Text",
};

api.boundAction("accounts", new Guid(""), "sample_BoundAction", inputs)
    .then((result: any) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

api.unboundAction("sample_UnboundAction", inputs)
    .then((result: any) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

// demonstrate bound function
const inputs3: FunctionInput[] = [];

inputs3.push({
    name: "Argument",
    value: "Value",
});

api.boundAction("accounts", new Guid(""), "sample_BoundFunction", inputs3)
    .then((result: any) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

// demonstrate create. Custom action - Add note to account
const inputs4: FunctionInput[] = [];

inputs4.push({
    alias: "tid",
    name: "Target",
    value: "{'@odata.id':'accounts(87989176-0887-45D1-93DA-4D5F228C10E6)'}",
});

api.unboundAction("sample_UnboundAction", inputs4)
    .then((result: any) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

// demonstrate batch operation
const changeSets: ChangeSet[] = [
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

const gets: string[] = [
    "accounts?$select=name",
];

api.batchOperation("BATCH123", "CHANGESET123", changeSets, gets)
    .then((result) => {
        console.log(result);
    }, (error) => {
        console.log(error);
    });
