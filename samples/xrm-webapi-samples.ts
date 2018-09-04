// tslint:disable:no-console
// tslint:disable:no-empty

import {
    ChangeSet,
    FunctionInput,
    Guid,
    create,
    createWithReturnData,
    retrieve,
    retrieveMultiple,
    update,
    updateProperty,
    deleteProperty,
    deleteRecord,
    disassociate,
    associate,
    boundAction,
    unboundAction,
    batchOperation,
    WebApiConfig,
    retrieveMultipleNextPage
} from "../src/xrm-webapi";

const config: WebApiConfig = new WebApiConfig("8.1");

// demonstrate create
const account: any = {
    name: "Test Account"
};

create(config, "accounts", account)
    .then(() => {
        console.log();
    }, (error) => {
        console.log(error);
    });

// demonstrate create with returned odata
createWithReturnData(config, "accounts", account, "$select=name,accountid")
    .then((created: any) => {
        console.log(created.name);
    });

// demonstrate retrieve
retrieve(config, "accounts", new Guid(""), "$select=name")
    .then((retrieved) => {
        console.log(retrieved.data.name);
    }, (error) => {
        console.log(error);
    });

// demonstrate retrieve multiple
const options: string = "$filter=name eq 'Test Account'&$select=name,accountid";

retrieveMultiple(config, "accounts", options)
    .then(
        (results) => {
            const accounts: any[] = [];
            for (let record of results.value) {
                accounts.push(record);
            }

            // demonstrate getting next page from retreiveMultiple
            retrieveMultipleNextPage(config, results["@odata.nextlink"]).then(
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
update(config, "accounts", new Guid(""), account)
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate update property. Update property returns no content
updateProperty(config, "accounts", new Guid(""), "name", "Updated Account")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate delete. Delete returns no content
deleteRecord(config, "accounts", new Guid(""))
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate delete property. Delete property returns no content
deleteProperty(config, "accounts", new Guid(""), "address1_line1")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate delete navigation property. Delete property returns no content
deleteProperty(config, "accounts", new Guid(""), "primarycontactid")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate associate. Associate returns no content
associate(config, "accounts", new Guid(""), "contact_customer_accounts", "contacts", new Guid(""))
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate disassociate. Disassociate returns no content
disassociate(config, "accounts", new Guid(""), "contact_customer_accounts")
    .then(() => {}, (error) => {
        console.log(error);
    });

// demonstrate bound action
const inputs: object = {
    NumberInput: 100,
    StringInput: "Text",
};

boundAction(config, "accounts", new Guid(""), "sample_BoundAction", inputs)
    .then((result: any) => {
        console.log(result.annotationid);
    }, (error) => {
        console.log(error);
    });

unboundAction(config, "sample_UnboundAction", inputs)
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

boundAction(config, "accounts", new Guid(""), "sample_BoundFunction", inputs3)
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

unboundAction(config, "sample_UnboundAction", inputs4)
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

batchOperation(config, "BATCH123", "CHANGESET123", changeSets, gets)
    .then((result) => {
        console.log(result);
    }, (error) => {
        console.log(error);
    });
