import { AxiosPromise } from 'axios';
export interface FunctionInput {
    name: string;
    value: string;
    alias?: string;
}
export declare class Guid {
    value: string;
    constructor(value: string);
    areEqual(compare: Guid): boolean;
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
export declare class WebApi {
    private version;
    private accessToken;
    private url;
    /**
     * Constructor
     * @param version Version must be 8.0, 8.1 or 8.2
     * @param accessToken Optional access token if using from outside Dynamics 365
     */
    constructor(version: string, accessToken?: string, url?: string);
    /**
     * Get the OData URL
     * @param queryString Query string to append to URL. Defaults to a blank string
     */
    getClientUrl(queryString?: string): string;
    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    retrieve(entitySet: string, id: Guid, queryString?: string, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    retrieveMultiple(entitySet: string, queryString?: string, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Retrieve next page from a retrieveMultiple request
     * @param query Query from the @odata.nextlink property of a retrieveMultiple
     * @param queryOptions Various query options for the query
     */
    getNextPage(query: string, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param impersonateUser Impersonate another user
     */
    create(entitySet: string, entity: object, queryOptions?: QueryOptions): AxiosPromise<CreatedEntity>;
    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param select Select odata query parameter
     * @param impersonateUser Impersonate another user
     */
    createWithReturnData(entitySet: string, entity: object, select: string, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param impersonateUser Impersonate another user
     */
    update(entitySet: string, id: Guid, entity: object, queryOptions?: QueryOptions): Promise<any>;
    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     * @param impersonateUser Impersonate another user
     */
    updateProperty(entitySet: string, id: Guid, attribute: string, value: any, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    delete(entitySet: string, id: Guid): AxiosPromise<any>;
    /**
     * Delete a property from a record in CRM. Non navigation properties only
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    deleteProperty(entitySet: string, id: Guid, attribute: string): AxiosPromise<any>;
    /**
     * Associate two records
     * @param entitySet Type of entity for primary record
     * @param id Id of primary record
     * @param relationship Schema name of relationship
     * @param relatedEntitySet Type of entity for secondary record
     * @param relatedEntityId Id of secondary record
     * @param impersonateUser Impersonate another user
     */
    associate(entitySet: string, id: Guid, relationship: string, relatedEntitySet: string, relatedEntityId: Guid, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Disassociate two records
     * @param entitySet Type of entity for primary record
     * @param id  Id of primary record
     * @param property Schema name of property or relationship
     * @param relatedEntityId Id of secondary record. Only needed for collection-valued navigation properties
     */
    disassociate(entitySet: string, id: Guid, property: string, relatedEntityId?: Guid): AxiosPromise<any>;
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    boundAction(entitySet: string, id: Guid, actionName: string, inputs?: Object, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    unboundAction(actionName: string, inputs?: Object, queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    boundFunction(entitySet: string, id: Guid, functionName: string, inputs?: FunctionInput[], queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    unboundFunction(functionName: string, inputs?: FunctionInput[], queryOptions?: QueryOptions): AxiosPromise<any>;
    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     * @param impersonateUser Impersonate another user
     */
    batchOperation(batchId: string, changeSetId: string, changeSets: ChangeSet[], batchGets: string[], queryOptions?: QueryOptions): AxiosPromise<any>;
    private getRequestConfig(method, queryString, queryOptions, contentType?, needsUrl?);
    private getPreferHeader(queryOptions);
    private getFunctionInputs(queryString, inputs);
}
