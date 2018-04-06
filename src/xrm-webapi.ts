import { AxiosRequestConfig, AxiosPromise } from 'axios';
import axios from 'axios';

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

    areEqual(compare: Guid): boolean {
        if (this === null || compare === null || this === undefined || compare === undefined) {
            return false;
        }

        return this.value.toLowerCase() === compare.value.toLowerCase();
    }
}

export interface CreatedEntity {
    id: Guid;
    uri: string;
}

export interface ChangeSet {
    queryString: string;
    entity: object;
}

export interface QueryOptions {
    includeFormattedValues?: boolean;
    includeLookupLogicalNames?: boolean;
    includeAssociatedNavigationProperties?: boolean;
    maxPageSize?: number;
    impersonateUser?: Guid;
    representation?: boolean;
}

export interface Entity {
    [propName: string]: any;
}

export interface RetrieveMultipleResponse {
    value: Entity[];
    "@odata.nextlink": string;
}

export class WebApi {
    private version: string;
    private accessToken: string;
    private url: string;

    /**
     * Constructor
     * @param version Version must be 8.0, 8.1, 8.2 or 9.0
     * @param accessToken Optional access token if using from outside Dynamics 365
     * @param url Optional url if using from outside Dynamics 365
     */
    constructor (version: string, accessToken?: string, url?: string) {
        this.version = version;
        this.accessToken = accessToken;
        this.url = url;

        axios.defaults.headers.common["Accept"] = "application/json";
        axios.defaults.headers.common["OData-MaxVersion"] = "4.0";
        axios.defaults.headers.common["OData-Version"] = "4.0";

        if (this.accessToken != null) {
            axios.defaults.headers.common["Authorization"] = `Bearer ${this.accessToken}`;
        }
    }

    /**
     * Get the OData URL
     * @param queryString Query string to append to URL. Defaults to a blank string
     */
    public getClientUrl(queryString: string = ""): string {
        if (this.url != null) {
            return `${this.url}/api/data/v${this.version}/${queryString}`;
        }

        const context: Xrm.Context = typeof GetGlobalContext !== "undefined" ? GetGlobalContext() : Xrm.Page.context;
        const url: string = `${context.getClientUrl()}/api/data/v${this.version}/${queryString}`;

        return url;
    }

    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    public retrieve(entitySet: string, id: Guid, queryString?: string, queryOptions?: QueryOptions): AxiosPromise<Entity> {
        if (queryString != null && ! /^[?]/.test(queryString)) {
            queryString = `?${queryString}`;
        }

        let query: string = (queryString != null) ? `${entitySet}(${id.value})${queryString}` : `${entitySet}(${id.value})`;
        const config: AxiosRequestConfig = this.getRequestConfig("GET", query, queryOptions);

        return axios(config);
    }

    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    public retrieveMultiple(entitySet: string, queryString?: string, queryOptions?: QueryOptions): AxiosPromise<RetrieveMultipleResponse> {
        if (queryString != null && ! /^[?]/.test(queryString)) {
            queryString = `?${queryString}`;
        }

        let query: string = (queryString != null) ? entitySet + queryString : entitySet;
        const config: AxiosRequestConfig = this.getRequestConfig("GET", query, queryOptions);

        return axios(config);
    }

    /**
     * Retrieve next page from a retrieveMultiple request
     * @param query Query from the @odata.nextlink property of a retrieveMultiple
     * @param queryOptions Various query options for the query
     */
    public getNextPage(query: string, queryOptions?: QueryOptions): AxiosPromise<RetrieveMultipleResponse> {
        const config: AxiosRequestConfig = this.getRequestConfig("GET", query, queryOptions, null, false);

        return axios(config);
    }

    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param queryOptions Various query options for the query
     */
    public create(entitySet: string, entity: Entity, queryOptions?: QueryOptions): AxiosPromise<CreatedEntity> {
        const config: AxiosRequestConfig = this.getRequestConfig("POST", entitySet, queryOptions);

        config.transformResponse = (data, headers) => {
            const uri: string = headers["OData-EntityId"];
            const start: number = uri.indexOf("(") + 1;
            const end: number = uri.indexOf(")", start);
            const id: string = uri.substring(start, end);

            data = {
                id: new Guid(id),
                uri
            };
        };
        
        config.data = JSON.stringify(entity);

        return axios(config);        
    }

    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param select Select odata query parameter
     * @param queryOptions Various query options for the query
     */
    public createWithReturnData(entitySet: string, entity: Entity, select: string, queryOptions?: QueryOptions): AxiosPromise<Entity> {
        if (select != null && ! /^[?]/.test(select)) {
            select = `?${select}`;
        }

        // set reprensetation
        if (queryOptions == null) {
            queryOptions = {};
        }

        queryOptions.representation = true;

        const config: AxiosRequestConfig = this.getRequestConfig("POST", entitySet + select, queryOptions);

        config.data = JSON.stringify(entity);

        return axios(config);
    }

    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param queryOptions Various query options for the query
     */
    public update(entitySet: string, id: Guid, entity: Entity, queryOptions?: QueryOptions): AxiosPromise<null> {
        const config: AxiosRequestConfig = this.getRequestConfig("PATCH", `${entitySet}(${id.value})`, queryOptions);

        config.data = JSON.stringify(entity);

        return axios(config);
    }

    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param select Select odata query parameter
     * @param queryOptions Various query options for the query
     */
    public updateWithReturnData(entitySet: string, id: Guid, entity: Entity, select: string, queryOptions?: QueryOptions): AxiosPromise<Entity> {
        if (select != null && ! /^[?]/.test(select)) {
            select = `?${select}`;
        }

        // set representation
        if (queryOptions == null) {
            queryOptions = {};
        }

        queryOptions.representation = true;

        const config: AxiosRequestConfig = this.getRequestConfig("PATCH", `${entitySet}(${id.value})${select}`, queryOptions);

        config.data = JSON.stringify(entity);

        return axios(config);
    }

    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     * @param queryOptions Various query options for the query
     */
    public updateProperty(entitySet: string, id: Guid, attribute: string, value: any, queryOptions?: QueryOptions): AxiosPromise<null> {
        const config: AxiosRequestConfig = this.getRequestConfig("PUT", `${entitySet}(${id.value})/${attribute}`, queryOptions);

        config.data = JSON.stringify({ value: value });
        
        return axios(config);
    }

    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    public delete(entitySet: string, id: Guid): AxiosPromise<null> {
        const config: AxiosRequestConfig = this.getRequestConfig("DELETE", `${entitySet}(${id.value})`, null);

        return axios(config);
    }

    /**
     * Delete a property from a record in CRM. Non navigation properties only
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    public deleteProperty(entitySet: string, id: Guid, attribute: string): AxiosPromise<null> {
        let queryString: string = `/${attribute}`;

        const config: AxiosRequestConfig = this.getRequestConfig("DELETE", `${entitySet}(${id.value})${queryString}`, null);

        return axios(config);
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
    public associate(entitySet: string, id: Guid, relationship: string, relatedEntitySet: string, relatedEntityId: Guid, queryOptions?: QueryOptions): AxiosPromise<null> {
        const config: AxiosRequestConfig = this.getRequestConfig("POST", `${entitySet}(${id.value})/${relationship}/$ref`, queryOptions);

        const related: object = {
            "@odata.id": this.getClientUrl(`${relatedEntitySet}(${relatedEntityId.value})`)
        };

        config.data = JSON.stringify(related);

        return axios(config);
    }

    /**
     * Disassociate two records
     * @param entitySet Type of entity for primary record
     * @param id  Id of primary record
     * @param property Schema name of property or relationship
     * @param relatedEntityId Id of secondary record. Only needed for collection-valued navigation properties
     */
    public disassociate(entitySet: string, id: Guid, property: string, relatedEntityId?: Guid): AxiosPromise<null> {
        let queryString: string = property;

        if (relatedEntityId != null) {
            queryString += `(${relatedEntityId.value})`;
        }

        queryString += "/$ref";

        const config: AxiosRequestConfig = this.getRequestConfig("DELETE", `${entitySet}(${id.value})/${queryString}`, null);

        return axios(config);
    }

    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public boundAction(entitySet: string, id: Guid, actionName: string, inputs?: Object, queryOptions?: QueryOptions): AxiosPromise<any> {
        const config: AxiosRequestConfig = this.getRequestConfig("POST", `${entitySet}(${id.value})/Microsoft.Dynamics.CRM.${actionName}`, queryOptions);

        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }

        return axios(config);
    }

    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public unboundAction(actionName: string, inputs?: Object, queryOptions?: QueryOptions): AxiosPromise<any> {
        const config: AxiosRequestConfig = this.getRequestConfig("POST", actionName, queryOptions);

        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }

        return axios(config);
    }

    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public boundFunction(entitySet: string, id: Guid, functionName: string, inputs?: FunctionInput[], queryOptions?: QueryOptions): AxiosPromise<any> {
        let queryString: string = `${entitySet}(${id.value})/Microsoft.Dynamics.CRM.${functionName}(`;
        queryString = this.getFunctionInputs(queryString, inputs);

        const config: AxiosRequestConfig = this.getRequestConfig("GET", queryString, queryOptions);

        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }

        return axios(config);
    }

    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    public unboundFunction(functionName: string, inputs?: FunctionInput[], queryOptions?: QueryOptions): AxiosPromise<any> {
        let queryString: string = `${functionName}(`;
        queryString = this.getFunctionInputs(queryString, inputs);

        const config: AxiosRequestConfig = this.getRequestConfig("GET", queryString, queryOptions);

        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }

        return axios(config);
    }

    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     * @param queryOptions Various query options for the query
     */
    public batchOperation(batchId: string, changeSetId: string, changeSets: ChangeSet[], batchGets: string[], queryOptions?: QueryOptions): AxiosPromise<any> {
        const config: AxiosRequestConfig = this.getRequestConfig("POST", "$batch", queryOptions, `multipart/mixed;boundary=batch_${batchId}`);

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
            body.push(`POST ${this.getClientUrl(changeSets[i].queryString)} HTTP/1.1`);
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
            body.push(`GET ${this.getClientUrl(get)} HTTP/1.1`);
            body.push("Accept: application/json");
        }

        if (batchGets.length > 0) {
            body.push("");
        }

        body.push(`--batch_${batchId}--`);

        config.data = body.join("\r\n");

        return axios(config);
    }

    private getRequestConfig(method: string, queryString: string, queryOptions: QueryOptions,
        contentType: string = "application/json; charset=utf-8", needsUrl: boolean = true): AxiosRequestConfig {
        let url: string;

        if (needsUrl) {
            url = this.getClientUrl(queryString);
        } else {
            url = queryString;
        }

        // Get axios config
        const config: AxiosRequestConfig = {
            url: url,
            method: method,
            headers: {
                "Content-Type": contentType
            }
        };

        if (queryOptions != null && typeof(queryOptions) !== "undefined") {
            config.headers["Prefer"] = this.getPreferHeader(queryOptions);

            if (queryOptions.impersonateUser != null) {
                config.headers["MSCRMCallerID"] = queryOptions.impersonateUser.value;
            }
        }

        return config;
    }

    private getPreferHeader(queryOptions: QueryOptions): string {
        let prefer: string[] = [];

        // add max page size to prefer request header
        if (queryOptions.maxPageSize) {
            prefer.push(`odata.maxpagesize=${queryOptions.maxPageSize}`);
        }

        // add formatted values to prefer request header
        if (queryOptions.representation) {
            prefer.push("return=representation");
        } else if (queryOptions.includeFormattedValues && queryOptions.includeLookupLogicalNames &&
            queryOptions.includeAssociatedNavigationProperties) {
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

    private getFunctionInputs(queryString: string, inputs: FunctionInput[]): string {
        if (inputs == null) {
            return queryString + ")";
        }

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

        queryString = queryString.substr(0, queryString.length - 1) + ")";

        if (aliases !== "?") {
            queryString += aliases;
        }

        return queryString;
    }
}
