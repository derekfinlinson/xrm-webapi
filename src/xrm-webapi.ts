/// <reference path="./typings/index.d.ts" />
import {Promise} from "es6-promise";

export class WebApi {
    private static request: XMLHttpRequest;

    private static getRequest(method: string, entitySet: string, queryString?: string) {
        const context = typeof GetGlobalContext != "undefined" ? GetGlobalContext() : Xrm.Page.context;
        let url = context.getClientUrl() + "/api/data/v8.0/" + entitySet;

        if (queryString) {
            url += queryString;
        }

        this.request = new XMLHttpRequest();
        this.request.open(method, url, true);
        this.request.setRequestHeader("Accept", "application/json");
        this.request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        this.request.setRequestHeader("OData-MaxVersion", "4.0");
        this.request.setRequestHeader("OData-Version", "4.0");
    }

    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     */
    static retrieve(entityType: string, id: string, queryString?: string, includeFormattedValues?: boolean) {
        if (queryString != null && ! /^[?]/.test(queryString)) queryString = `?${queryString}`;
        id = id.replace(/[{}]/g, "");

        this.getRequest("GET", entityType, `(${id})${queryString}`);

        if (includeFormattedValues) {
          this.request.setRequestHeader("Prefer", 'odata.include-annotations="OData.Community.Display.V1.FormattedValue');
        }

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 200) {
                        resolve(JSON.parse(this.request.response));
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            this.request.send();
        });
    }

    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     */
    static retrieveMultiple(entitySet: string, queryString?: string, includeFormattedValues?: boolean, maxPageSize?: number) {
        if (queryString != null && ! /^[?]/.test(queryString)) queryString = `?${queryString}`;

        this.getRequest("GET", entitySet, queryString);

        if (includeFormattedValues || maxPageSize) {
          this.request.setRequestHeader("Prefer", [
            includeFormattedValues ? 'odata.include-annotations="OData.Community.Display.V1.FormattedValue' : "",
            maxPageSize ? `odata.maxpagesize=${maxPageSize}` : ""
          ].join(","));
        }

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 200) {
                        resolve(JSON.parse(this.request.response));
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            this.request.send();
        });
    }

    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     */
    static create(entitySet: string, entity: Object) {
        this.getRequest("POST", entitySet);

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 200) {
                        resolve(this.request.getResponseHeader("OData-EntityId"));
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            this.request.send(entity);
        });
    }

    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     */
    static update(entitySet: string, id: string, entity: Object) {
        id = id.replace(/[{}]/g, "");
        this.getRequest("PATCH", entitySet, `(${id})`);

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            this.request.send(JSON.stringify(entity));
        });
    }

    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute logical name
     * @param value Update value
     */
    static updateProperty(entitySet: string, id: string, attribute: string, value: any) {
        id = id.replace(/[{}]/g, "");
        this.getRequest("PUT", entitySet, `(${id})`);

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            this.request.send(JSON.stringify({ attribute: value }));
        });
    }

    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    static delete(entitySet: string, id: string) {
        id = id.replace(/[{}]/g, "");
        this.getRequest("DELETE", entitySet, `(${id})`);

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            this.request.send();
        });
    }

    /**
     * Delete a property from a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    static deleteProperty(entitySet: string, id: string, attribute: string) {
        id = id.replace(/[{}]/g, "");
        this.getRequest("DELETE", entitySet, `(${id})/${attribute}`);

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            this.request.send();
        });
    }

    /**
     * Execute a default or custom action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id or record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     */
    static executeAction(entitySet: string, id: string, actionName: string, inputs?: Object) {
        id = id.replace(/[{}]/g, "");
        this.getRequest("POST", entitySet, `(${id})/Microsoft.Dynamics.CRM.${actionName}`);

        return new Promise((resolve, reject) => {
            this.request.onreadystatechange = () => {
                if (this.request.readyState === 4 /* complete */) {
                    this.request.onreadystatechange = null;
                    if (this.request.status === 200) {
                        resolve(JSON.parse(this.request.response));
                    } else if (this.request.status === 204) {
                        resolve();
                    } else {
                        reject(JSON.parse(this.request.response).error);
                    }
                }
            };

            inputs != null ? this.request.send(JSON.stringify(inputs)) : this.request.send();
        });
    }
}
