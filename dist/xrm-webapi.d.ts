import { Promise } from "es6-promise";
export interface FunctionInput {
    name: string;
    value: string;
    alias?: string;
}
export declare class Guid {
    value: string;
    constructor(value: string);
}
export interface Attribute {
    name: string;
    value?: any;
}
export interface Entity {
    id?: Guid;
    attributes: Array<Attribute>;
}
export interface CreatedEntity {
    id: Guid;
    uri: string;
}
export interface ChangeSet {
    queryString: string;
    object: Entity;
}
export declare class WebApi {
    private version;
    /**
     * Constructor. Version should be 8.0, 8.1 or 8.2
     */
    constructor(version: string);
    private getRequest(method, queryString, contentType?);
    private getFunctionInputs(queryString, inputs);
    private getPreferHeader(formattedValues, lookupLogicalNames, associatedNavigationProperties, maxPageSize?);
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
     * @param includeFormattedValues Include formatted values in results
     * @param includeLookupLogicalNames Include lookup logical names in results
     * @param includeAssociatedNavigationProperty Include associated navigation property in results
     */
    retrieve(entityType: string, id: Guid, queryString?: string, includeFormattedValues?: boolean, includeLookupLogicalNames?: boolean, includeAssociatedNavigationProperties?: boolean): Promise<any>;
    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param includeFormattedValues Include formatted values in results
     * @param includeLookupLogicalNames Include lookup logical names in results
     * @param includeAssociatedNavigationProperty Include associated navigation property in results
     * @param maxPageSize Records per page to return
     */
    retrieveMultiple(entitySet: string, queryString?: string, includeFormattedValues?: boolean, includeLookupLogicalNames?: boolean, includeAssociatedNavigationProperties?: boolean, maxPageSize?: number): Promise<any>;
    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     */
    create(entitySet: string, entity: Entity): Promise<CreatedEntity>;
    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     */
    update(entitySet: string, id: Guid, entity: Entity): Promise<any>;
    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     */
    updateProperty(entitySet: string, id: Guid, attribute: Attribute): Promise<any>;
    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    delete(entitySet: string, id: Guid): Promise<any>;
    /**
     * Delete a property from a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    deleteProperty(entitySet: string, id: Guid, attribute: Attribute, isNavigationProperty: boolean): Promise<any>;
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    boundAction(entitySet: string, id: Guid, actionName: string, inputs?: Object): Promise<any>;
    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    unboundAction(actionName: string, inputs?: Object): Promise<any>;
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    boundFunction(entitySet: string, id: Guid, functionName: string, inputs?: Array<FunctionInput>): Promise<any>;
    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    unboundFunction(functionName: string, inputs?: Array<FunctionInput>): Promise<any>;
    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     */
    batchOperation(batchId: string, changeSetId: string, changeSets: Array<ChangeSet>, batchGets: Array<string>): Promise<any>;
}
