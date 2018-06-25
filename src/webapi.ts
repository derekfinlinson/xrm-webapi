import { Guid, QueryOptions, Entity, RetrieveMultipleResponse, FunctionInput, ChangeSet } from "./models";
import { WebApiRequest, WebApiRequestConfig, WebApiRequestResult } from "./request";

export interface WebApiConfig {
    version: string;
    accessToken?: string;
    url?: string;
}

export class WebApi {
    config: WebApiConfig;

    /**
     * Constructor
     * @param config WebApiConfig
     */
    constructor (config: WebApiConfig) {
        // If URL not provided, get it from Xrm.Context
        if (config.url == null) {
            const context: Xrm.Context = typeof GetGlobalContext !== "undefined" ? GetGlobalContext() : Xrm.Page.context;
            const url: string = `${context.getClientUrl()}/api/data/v${config.version}`;

            config.url = url;
        } else {
            config.url = `${config.url}/api/data/v${config.version}`;
        }

        this.config = config;
    }

    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    public retrieve(entitySet: string, id: Guid, queryString?: string, queryOptions?: QueryOptions): Promise<Entity> {
        if (queryString != null && ! /^[?]/.test(queryString)) {
            queryString = `?${queryString}`;
        }

        let query: string = (queryString != null) ? `${entitySet}(${id.value})${queryString}` : `${entitySet}(${id.value})`;

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "GET",
            contentType: "application/json; charset=utf-8",
            queryString: query
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve(JSON.parse(result.response));
                    }
                }
            );
        });
    }

    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    public retrieveMultiple(entitySet: string, queryString?: string, queryOptions?: QueryOptions): Promise<RetrieveMultipleResponse> {
        if (queryString != null && ! /^[?]/.test(queryString)) {
            queryString = `?${queryString}`;
        }

        let query: string = (queryString != null) ? entitySet + queryString : entitySet;

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "GET",
            contentType: "application/json; charset=utf-8",
            queryString: query
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve(JSON.parse(result.response));
                    }
                }
            );
        });
    }

    /**
     * Retrieve next page from a retrieveMultiple request
     * @param url Query from the @odata.nextlink property of a retrieveMultiple
     * @param queryOptions Various query options for the query
     */
    public getNextPage(url: string, queryOptions?: QueryOptions): Promise<RetrieveMultipleResponse> {
        this.config.url = url;

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "GET",
            contentType: "application/json; charset=utf-8",
            queryString: ""
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve(JSON.parse(result.response));
                    }
                }
            );
        });
    }

    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param queryOptions Various query options for the query
     */
    public create(entitySet: string, entity: Entity, queryOptions?: QueryOptions): Promise<null> {
        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "POST",
            contentType: "application/json; charset=utf-8",
            queryString: entitySet,
            body: JSON.stringify(entity)
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param select Select odata query parameter
     * @param queryOptions Various query options for the query
     */
    public createWithReturnData(entitySet: string, entity: Entity, select: string, queryOptions?: QueryOptions): Promise<Entity> {
        if (select != null && ! /^[?]/.test(select)) {
            select = `?${select}`;
        }

        // set reprensetation
        if (queryOptions == null) {
            queryOptions = {};
        }

        queryOptions.representation = true;

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "POST",
            contentType: "application/json; charset=utf-8",
            queryString: entitySet + select,
            body: JSON.stringify(entity)
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve(JSON.parse(result.response));
                    }
                }
            );
        });
    }

    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param queryOptions Various query options for the query
     */
    public update(entitySet: string, id: Guid, entity: Entity, queryOptions?: QueryOptions): Promise<null> {
        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "PATCH",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})`,
            body: JSON.stringify(entity)
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param select Select odata query parameter
     * @param queryOptions Various query options for the query
     */
    public updateWithReturnData(entitySet: string, id: Guid, entity: Entity, select: string, queryOptions?: QueryOptions): Promise<Entity> {
        if (select != null && ! /^[?]/.test(select)) {
            select = `?${select}`;
        }

        // set representation
        if (queryOptions == null) {
            queryOptions = {};
        }

        queryOptions.representation = true;

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "PATCH",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})${select}`,
            body: JSON.stringify(entity)
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve(JSON.parse(result.response));
                    }
                }
            );
        });
    }

    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     * @param queryOptions Various query options for the query
     */
    public updateProperty(entitySet: string, id: Guid, attribute: string, value: any, queryOptions?: QueryOptions): Promise<null> {
        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "PUT",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})/${attribute}`,
            body: JSON.stringify({ value: value })
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    public delete(entitySet: string, id: Guid): Promise<null> {
        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "DELETE",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})`
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, null,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    /**
     * Delete a property from a record in CRM. Non navigation properties only
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    public deleteProperty(entitySet: string, id: Guid, attribute: string): Promise<null> {
        let queryString: string = `/${attribute}`;

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "DELETE",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})${queryString}`
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, null,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    /**
     * Associate two records
     * @param entitySet Type of entity for primary record
     * @param id Id of primary record
     * @param relationship Schema name of relationship
     * @param relatedEntitySet Type of entity for secondary record
     * @param relatedEntityId Id of secondary record
     * @param queryOptions Various query options for the query
     */
    public associate(entitySet: string, id: Guid, relationship: string, relatedEntitySet: string,
        relatedEntityId: Guid, queryOptions?: QueryOptions): Promise<null> {
        const related: object = {
            "@odata.id": `${this.config.url}/${relatedEntitySet}(${relatedEntityId.value})`
        };

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "POST",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})/${relationship}/$ref`,
            body: JSON.stringify(related)
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    /**
     * Disassociate two records
     * @param entitySet Type of entity for primary record
     * @param id  Id of primary record
     * @param property Schema name of property or relationship
     * @param relatedEntityId Id of secondary record. Only needed for collection-valued navigation properties
     */
    public disassociate(entitySet: string, id: Guid, property: string, relatedEntityId?: Guid): Promise<null> {
        let queryString: string = property;

        if (relatedEntityId != null) {
            queryString += `(${relatedEntityId.value})`;
        }

        queryString += "/$ref";

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "DELETE",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})/${queryString}`
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, null,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        resolve();
                    }
                }
            );
        });
    }

    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public boundAction(entitySet: string, id: Guid, actionName: string, inputs?: Object, queryOptions?: QueryOptions): Promise<any> {
        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "POST",
            contentType: "application/json; charset=utf-8",
            queryString: `${entitySet}(${id.value})/Microsoft.Dynamics.CRM.${actionName}`
        };

        if (inputs != null) {
            config.body = JSON.stringify(inputs);
        }

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        if (result.response) {
                            resolve(JSON.parse(result.response));
                        } else {
                            resolve();
                        }
                    }
                }
            );
        });
    }

    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public unboundAction(actionName: string, inputs?: Object, queryOptions?: QueryOptions): Promise<any> {
        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "POST",
            contentType: "application/json; charset=utf-8",
            queryString: actionName
        };

        if (inputs != null) {
            config.body = JSON.stringify(inputs);
        }

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        if (result.response) {
                            resolve(JSON.parse(result.response));
                        } else {
                            resolve();
                        }
                    }
                }
            );
        });
    }

    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public boundFunction(entitySet: string, id: Guid, functionName: string,
        inputs?: FunctionInput[], queryOptions?: QueryOptions): Promise<any> {
        let queryString: string = `${entitySet}(${id.value})/Microsoft.Dynamics.CRM.${functionName}(`;
        queryString = this.getFunctionInputs(queryString, inputs);

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "GET",
            contentType: "application/json; charset=utf-8",
            queryString: queryString
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        if (result.response) {
                            resolve(JSON.parse(result.response));
                        } else {
                            resolve();
                        }
                    }
                }
            );
        });
    }

    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public unboundFunction(functionName: string, inputs?: FunctionInput[], queryOptions?: QueryOptions): Promise<any> {
        let queryString: string = `${functionName}(`;
        queryString = this.getFunctionInputs(queryString, inputs);

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "GET",
            contentType: "application/json; charset=utf-8",
            queryString: queryString
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        if (result.response) {
                            resolve(JSON.parse(result.response));
                        } else {
                            resolve();
                        }
                    }
                }
            );
        });
    }

    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     * @param queryOptions Various query options for the query
     */
    public batchOperation(batchId: string, changeSetId: string, changeSets: ChangeSet[],
        batchGets: string[], queryOptions?: QueryOptions): Promise<any> {
        // build post body
        const body: string[] = [];

        if (changeSets.length > 0) {
            body.push(`--batch_${batchId}`);
            body.push(`Content-Type: multipart/mixed;boundary=changeset_${changeSetId}`);
            body.push("");
        }

        // push change sets to body
        for (let i: number = 0; i < changeSets.length; i++) {
            body.push(`--changeset_${changeSetId}`);
            body.push("Content-Type: application/http");
            body.push("Content-Transfer-Encoding:binary");
            body.push(`Content-ID: ${i + 1}`);
            body.push("");
            body.push(`POST ${this.config.url}/${changeSets[i].queryString} HTTP/1.1`);
            body.push("Content-Type: application/json;type=entry");
            body.push("");

            body.push(JSON.stringify(changeSets[i].entity));
        }

        if (changeSets.length > 0) {
            body.push(`--changeset_${changeSetId}--`);
            body.push("");
        }

        // push get requests to body
        for (let get of batchGets) {
            body.push(`--batch_${batchId}`);
            body.push("Content-Type: application/http");
            body.push("Content-Transfer-Encoding:binary");
            body.push("");
            body.push(`GET ${this.config.url}/${get} HTTP/1.1`);
            body.push("Accept: application/json");
        }

        if (batchGets.length > 0) {
            body.push("");
        }

        body.push(`--batch_${batchId}--`);

        const request: WebApiRequest = new WebApiRequest(this.config);

        const config: WebApiRequestConfig = {
            method: "POST",
            contentType: `multipart/mixed;boundary=batch_${batchId}`,
            queryString: "$batch",
            body: body.join("\r\n")
        };

        return new Promise((resolve, reject) => {
            request.submitRequest(config, queryOptions,
                (result: WebApiRequestResult) => {
                    if (result.error) {
                        reject(this.handleError(result.response));
                    } else {
                        if (result.response) {
                            resolve(JSON.parse(result.response));
                        } else {
                            resolve();
                        }
                    }
                }
            );
        });
    }

    private getFunctionInputs(queryString: string, inputs: FunctionInput[]): string {
        if (inputs == null) {
            return queryString + ")";
        }

        let aliases: string[] = [];

        for (let i: number = 0; i < inputs.length; i++) {
            queryString += inputs[i].name;

            if (inputs[i].alias) {
                queryString += `=@${inputs[i].alias},`;
                aliases.push(`@${inputs[i].alias}=${inputs[i].value}`);
            } else {
                queryString += `=${inputs[i].value},`;
            }
        }

        queryString = queryString.substr(0, queryString.length - 1) + ")";

        if (aliases.length > 0) {
            queryString += `?${aliases.join("&")}`;
        }

        return queryString;
    }

    private handleError(result: any): any {
        try {
            return JSON.parse(result).error;
        } catch (e) {
            return new Error("Unexpected Error");
        }
    }
}
