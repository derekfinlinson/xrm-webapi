export interface FunctionInput {
    name: string;
    value: string;
    alias?: string;
}

export class Guid {
    public value: string;

    constructor(value: string) {
        value = value.replace(/[{}]/g, "");

        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value)) {
            this.value = value.toUpperCase();
        } else {
            throw Error(`Id ${value} is not a valid GUID`);
        }
    }
}

export interface Attribute {
    name: string;
    value?: any;
}

export interface Entity {
    id?: Guid;
    attributes: Attribute[];
}

export interface CreatedEntity {
    id: Guid;
    uri: string;
}

export interface ChangeSet {
    queryString: string;
    object: Entity;
}

export interface QueryOptions {
    includeFormattedValues?: boolean;
    includeLookupLogicalNames?: boolean;
    includeAssociatedNavigationProperties?: boolean;
    maxPageSize?: number;
}

export class WebApi {
    private version;
    private accessToken;

    /**
     * Constructor
     * @param version Version must be 8.0, 8.1 or 8.2
     * @param accessToken Optional access token if using from outside Dynamics 365
     */
    constructor (version: string, accessToken?: string) {
        this.version = version;
        this.accessToken = accessToken;
    }

    /**
     * Get the OData URL
     * @param queryString Query string to append to URL. Defaults to a blank string
     */
    public getClientUrl(queryString: string = ""): string {
        const context: Xrm.Context = typeof GetGlobalContext !== "undefined" ? GetGlobalContext() : Xrm.Page.context;
        const url: string = context.getClientUrl() + `/api/data/v${this.version}/` + queryString;

        return url;
    }

    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    public retrieve(entitySet: string, id: Guid, queryString?: string, queryOptions?: QueryOptions): Promise<any> {
        if (queryString != null && ! /^[?]/.test(queryString)) {
            queryString = `?${queryString}`;
        }

        let query: string = (queryString != null) ? `${entitySet}(${id.value})${queryString}` : `${entitySet}(${id.value})`;
        const req: XMLHttpRequest = this.getRequest("GET", query);

        if (queryOptions != null && typeof(queryOptions) !== "undefined") {
          req.setRequestHeader("Prefer", this.getPreferHeader(queryOptions));
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            req.send();
        });
    }

    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    public retrieveMultiple(entitySet: string, queryString?: string, queryOptions?: QueryOptions): Promise<any> {
        if (queryString != null && ! /^[?]/.test(queryString)) {
            queryString = `?${queryString}`;
        }

        let query: string = (queryString != null) ? entitySet + queryString : entitySet;
        const req = this.getRequest("GET", query);

        if (queryOptions != null && typeof(queryOptions) !== "undefined") {
          req.setRequestHeader("Prefer", this.getPreferHeader(queryOptions));
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            req.send();
        });
    }

    /**
     * Retrieve next page from a retrieveMultiple request
     * @param query Query from the @odata.nextlink property of a retrieveMultiple
     * @param queryOptions Various query options for the query
     */
    public getNextPage(query: string, queryOptions?: QueryOptions): Promise<any> {
        const req = this.getRequest("GET", query, undefined, false);

        if (queryOptions != null) {
          req.setRequestHeader("Prefer", this.getPreferHeader(queryOptions));
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            req.send();
        });
    }

    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param impersonateUser Impersonate another user
     */
    public create(entitySet: string, entity: Entity, impersonateUser?: Guid): Promise<CreatedEntity> {
        const req = this.getRequest("POST", entitySet);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        const uri = req.getResponseHeader("OData-EntityId");
                        const start = uri.indexOf("(") + 1;
                        const end = uri.indexOf(")", start);
                        const id = uri.substring(start, end);

                        const createdEntity: CreatedEntity = {
                            id: new Guid(id),
                            uri,
                        };

                        resolve(createdEntity);
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            let attributes = {};

            entity.attributes.forEach((attribute) => {
                attributes[attribute.name] = attribute.value;
            });

            req.send(JSON.stringify(attributes));
        });
    }

    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param select Select odata query parameter
     * @param impersonateUser Impersonate another user
     */
    public createWithReturnData(entitySet: string, entity: Entity, select: string, impersonateUser?: Guid): Promise<any> {
        if (select != null && ! /^[?]/.test(select)) {
            select = `?${select}`;
        }

        const req = this.getRequest("POST", entitySet + select);

        req.setRequestHeader("Prefer", "return=representation");

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 201) {
                        resolve(JSON.parse(req.response));
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            let attributes = {};

            entity.attributes.forEach((attribute) => {
                attributes[attribute.name] = attribute.value;
            });

            req.send(JSON.stringify(attributes));
        });
    }

    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param impersonateUser Impersonate another user
     */
    public update(entitySet: string, id: Guid, entity: Entity, impersonateUser?: Guid): Promise<any> {
        const req = this.getRequest("PATCH", `${entitySet}(${id.value})`);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            let attributes = {};

            entity.attributes.forEach((attribute) => {
                attributes[attribute.name] = attribute.value;
            });

            req.send(JSON.stringify(attributes));
        });
    }

    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     * @param impersonateUser Impersonate another user
     */
    public updateProperty(entitySet: string, id: Guid, attribute: Attribute, impersonateUser?: Guid): Promise<any> {
        const req = this.getRequest("PUT", `${entitySet}(${id.value})/${attribute.name}`);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            req.send(JSON.stringify({ value: attribute.value }));
        });
    }

    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    public delete(entitySet: string, id: Guid): Promise<any> {
        const req = this.getRequest("DELETE", `${entitySet}(${id.value})`);

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            req.send();
        });
    }

    /**
     * Delete a property from a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    public deleteProperty(entitySet: string, id: Guid, attribute: Attribute, isNavigationProperty: boolean): Promise<any> {
        let queryString = `/${attribute.name}`;

        if (isNavigationProperty) {
            queryString += "/$ref";
        }

        const req = this.getRequest("DELETE", `${entitySet}(${id.value})${queryString}`);

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            req.send();
        });
    }

    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    public boundAction(entitySet: string, id: Guid, actionName: string, inputs?: Object, impersonateUser?: Guid): Promise<any> {
        const req = this.getRequest("POST", `${entitySet}(${id.value})/Microsoft.Dynamics.CRM.${actionName}`);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    } else if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    }

    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    public unboundAction(actionName: string, inputs?: Object, impersonateUser?: Guid): Promise<any> {
        const req = this.getRequest("POST", actionName);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    } else if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    }

    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    public boundFunction(entitySet: string, id: Guid, functionName: string, inputs?: FunctionInput[], impersonateUser?: Guid): Promise<any> {
        let queryString = `${entitySet}(${id.value})/Microsoft.Dynamics.CRM.${functionName}(`;
        queryString = this.getFunctionInputs(queryString, inputs);

        const req = this.getRequest("GET", queryString);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    } else if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    }

    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    public unboundFunction(functionName: string, inputs?: FunctionInput[], impersonateUser?: Guid): Promise<any> {
        let queryString = `${functionName}(`;
        queryString = this.getFunctionInputs(queryString, inputs);

        const req = this.getRequest("GET", queryString);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    } else if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    }

    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     * @param impersonateUser Impersonate another user
     */
    public batchOperation(batchId: string, changeSetId: string, changeSets: ChangeSet[], batchGets: string[], impersonateUser?: Guid): Promise<any> {
        const req = this.getRequest("POST", "$batch", `multipart/mixed;boundary=batch_${batchId}`);

        if (impersonateUser != null) {
            req.setRequestHeader("MSCRMCallerID", impersonateUser.value);
        }

        // build post body
        const body = [
            `--batch_${batchId}`,
            `Content-Type: multipart/mixed;boundary=changeset_${changeSetId}`,
            "",
        ];

        // push change sets to body
        for (let i = 0; i < changeSets.length; i++) {
            body.push(`--changeset_${changeSetId}`);
            body.push("Content-Type: application/http");
            body.push("Content-Transfer-Encoding:binary");
            body.push(`Content-ID: ${i + 1}`);
            body.push("");
            body.push(`POST ${this.getClientUrl(changeSets[i].queryString)} HTTP/1.1`);
            body.push("Content-Type: application/json;type=entry");
            body.push("");

            let attributes = {};

            changeSets[i].object.attributes.forEach((attribute) => {
                attributes[attribute.name] = attribute.value;
            });

            body.push(JSON.stringify(attributes));
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
            body.push(`GET ${this.getClientUrl(get)} HTTP/1.1`);
            body.push("Accept: application/json");
        }

        body.push(`--batch_${batchId}--`);

        return new Promise((resolve, reject) => {
            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(req.response);
                    } else if (req.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };

            req.send(body.join("\r\n"));
        });
    }

    private getRequest(method: string, queryString: string, contentType: string = "application/json; charset=utf-8", needsUrl: boolean = true): XMLHttpRequest {
        let url: string;

        if (needsUrl) {
            url = this.getClientUrl(queryString);
        } else {
            url = queryString;
        }

        // Build XMLHttpRequest
        const request: XMLHttpRequest = new XMLHttpRequest();
        request.open(method, url, true);
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", contentType);
        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");
        request.setRequestHeader("Cache-Control", "no-cache");

        if (this.accessToken != null) {
            request.setRequestHeader("Authorization", `Bearer ${this.accessToken}`);
        }
        
        return request;
    }

    private getFunctionInputs(queryString: string, inputs: FunctionInput[]): string {
        let aliases: string = "?";

        for (let i: number = 0; i < inputs.length; i++) {
            queryString += inputs[i].name;

            if (inputs[i].alias) {
                queryString += `=@${inputs[i].alias},`;
                aliases += `@${inputs[i].alias}=${inputs[i].value}`;
            } else {
                queryString += `=${inputs[i].value},`;
            }
        }

        queryString += queryString.substr(0, queryString.length - 1) + ")";

        if (aliases !== "?") {
            queryString += aliases;
        }

        return queryString;
    }

    private getPreferHeader(queryOptions: QueryOptions): string {
        let prefer: string[] = [];

        // Add max page size to prefer request header
        if (queryOptions.maxPageSize) {
            prefer.push(`odata.maxpagesize=${queryOptions.maxPageSize}`);
        }

        // Add formatted values to prefer request header
        if (queryOptions.includeFormattedValues && queryOptions.includeLookupLogicalNames && queryOptions.includeAssociatedNavigationProperties) {
            prefer.push("odata.include-annotations=\"*\"");
        } else {
            const preferExtra: string = [
                queryOptions.includeFormattedValues ? "OData.Community.Display.V1.FormattedValue" : "",
                queryOptions.includeLookupLogicalNames ? "Microsoft.Dynamics.CRM.lookuplogicalname" : "",
                queryOptions.includeAssociatedNavigationProperties ? "Microsoft.Dynamics.CRM.associatednavigationproperty" : "",
            ].filter((v, i) => {
                return v !== "";
            }).join(",");

            prefer.push("odata.include-annotations=\"" + preferExtra + "\"");
        }

        return prefer.join(",");
    }
}
