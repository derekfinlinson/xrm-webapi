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
export declare class WebApi {
    private static getRequest(method, queryString);
    private static getFunctionInputs(queryString, inputs);
    private static getPreferHeader(formattedValues, lookupLogicalNames, associatedNavigationProperties, maxPageSize?);
    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     * @param includeFormattedValues Include formatted values in results
     * @param includeLookupLogicalNames Include lookup logical names in results
     * @param includeAssociatedNavigationProperty Include associated navigation property in results
     */
    static retrieve(entityType: string, id: Guid, queryString?: string, includeFormattedValues?: boolean, includeLookupLogicalNames?: boolean, includeAssociatedNavigationProperties?: boolean): Promise<{}>;
    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param includeFormattedValues Include formatted values in results
     * @param includeLookupLogicalNames Include lookup logical names in results
     * @param includeAssociatedNavigationProperty Include associated navigation property in results
     * @param maxPageSize Records per page to return
     */
    static retrieveMultiple(entitySet: string, queryString?: string, includeFormattedValues?: boolean, includeLookupLogicalNames?: boolean, includeAssociatedNavigationProperties?: boolean, maxPageSize?: number): Promise<{}>;
    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     */
    static create(entitySet: string, entity: Entity): Promise<CreatedEntity>;
    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     */
    static update(entitySet: string, id: Guid, entity: Entity): Promise<{}>;
    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     */
    static updateProperty(entitySet: string, id: Guid, attribute: Attribute): Promise<{}>;
    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    static delete(entitySet: string, id: Guid): Promise<{}>;
    /**
     * Delete a property from a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    static deleteProperty(entitySet: string, id: Guid, attribute: Attribute): Promise<{}>;
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    static boundAction(entitySet: string, id: Guid, actionName: string, inputs?: Object): Promise<{}>;
    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    static unboundAction(actionName: string, inputs?: Object): Promise<{}>;
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    static boundFunction(entitySet: string, id: Guid, functionName: string, inputs?: Array<FunctionInput>): Promise<{}>;
    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    static unboundFunction(functionName: string, inputs?: Array<FunctionInput>): Promise<{}>;
}
