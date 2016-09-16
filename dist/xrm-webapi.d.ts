/// <reference path="../src/typings/index.d.ts" />
export declare class WebApi {
    private static request;
    private static getRequest(method, entitySet, queryString?);
    static retrieve(entityType: string, id: string, queryString?: string): Promise<{}>;
    static retrieveMultiple(entitySet: string, queryString?: string): Promise<{}>;
    static create(entitySet: string, entity: Object): Promise<{}>;
    static update(entitySet: string, id: string, entity: Object): Promise<{}>;
    static updateProperty(entitySet: string, id: string, attribute: string, value: any): Promise<{}>;
    static delete(entitySet: string, id: string): Promise<{}>;
    static deleteProperty(entitySet: string, id: string, attribute: string): Promise<{}>;
    static executeAction(entitySet: string, id: string, actionName: string, inputs: Object): Promise<{}>;
}
