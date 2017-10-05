"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var Guid = /** @class */ (function () {
    function Guid(value) {
        value = value.replace(/[{}]/g, "");
        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value)) {
            this.value = value.toUpperCase();
        }
        else {
            throw Error("Id " + value + " is not a valid GUID");
        }
    }
    Guid.prototype.areEqual = function (compare) {
        if (this === null || compare === null || this === undefined || compare === undefined) {
            return false;
        }
        return this.value.toLowerCase() === compare.value.toLowerCase();
    };
    return Guid;
}());
exports.Guid = Guid;
var WebApiBase = /** @class */ (function () {
    /**
     * Constructor
     * @param version Version must be 8.0, 8.1 or 8.2
     * @param accessToken Optional access token if using from outside Dynamics 365
     */
    function WebApiBase(version, accessToken, url) {
        this.version = version;
        this.accessToken = accessToken;
        this.url = url;
    }
    /**
     * Get the OData URL
     * @param queryString Query string to append to URL. Defaults to a blank string
     */
    WebApiBase.prototype.getClientUrl = function (queryString) {
        if (queryString === void 0) { queryString = ""; }
        if (this.url != null) {
            return this.url + "/api/data/v" + this.version + "/" + queryString;
        }
        var context = typeof GetGlobalContext !== "undefined" ? GetGlobalContext() : Xrm.Page.context;
        var url = context.getClientUrl() + "/api/data/v" + this.version + "/" + queryString;
        return url;
    };
    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    WebApiBase.prototype.retrieve = function (entitySet, id, queryString, queryOptions) {
        if (queryString != null && !/^[?]/.test(queryString)) {
            queryString = "?" + queryString;
        }
        var query = (queryString != null) ? entitySet + "(" + id.value + ")" + queryString : entitySet + "(" + id.value + ")";
        var req = this.getRequest("GET", query, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send();
        });
    };
    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    WebApiBase.prototype.retrieveMultiple = function (entitySet, queryString, queryOptions) {
        if (queryString != null && !/^[?]/.test(queryString)) {
            queryString = "?" + queryString;
        }
        var query = (queryString != null) ? entitySet + queryString : entitySet;
        var req = this.getRequest("GET", query, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send();
        });
    };
    /**
     * Retrieve next page from a retrieveMultiple request
     * @param query Query from the @odata.nextlink property of a retrieveMultiple
     * @param queryOptions Various query options for the query
     */
    WebApiBase.prototype.getNextPage = function (query, queryOptions) {
        var req = this.getRequest("GET", query, queryOptions, null, false);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send();
        });
    };
    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.create = function (entitySet, entity, queryOptions) {
        var req = this.getRequest("POST", entitySet, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        var uri = req.getResponseHeader("OData-EntityId");
                        var start = uri.indexOf("(") + 1;
                        var end = uri.indexOf(")", start);
                        var id = uri.substring(start, end);
                        var createdEntity = {
                            id: new Guid(id),
                            uri: uri,
                        };
                        resolve(createdEntity);
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send(JSON.stringify(entity));
        });
    };
    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param select Select odata query parameter
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.createWithReturnData = function (entitySet, entity, select, queryOptions) {
        if (select != null && !/^[?]/.test(select)) {
            select = "?" + select;
        }
        // set reprensetation
        if (queryOptions == null) {
            queryOptions = {};
        }
        queryOptions.representation = true;
        var req = this.getRequest("POST", entitySet + select, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 201) {
                        resolve(JSON.parse(req.response));
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send(JSON.stringify(entity));
        });
    };
    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.update = function (entitySet, id, entity, queryOptions) {
        var req = this.getRequest("PATCH", entitySet + "(" + id.value + ")", queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send(JSON.stringify(entity));
        });
    };
    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.updateProperty = function (entitySet, id, attribute, value, queryOptions) {
        var req = this.getRequest("PUT", entitySet + "(" + id.value + ")/" + attribute, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send(JSON.stringify({ value: value }));
        });
    };
    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    WebApiBase.prototype.delete = function (entitySet, id) {
        var req = this.getRequest("DELETE", entitySet + "(" + id.value + ")", null);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send();
        });
    };
    /**
     * Delete a property from a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    WebApiBase.prototype.deleteProperty = function (entitySet, id, attribute, isNavigationProperty) {
        var queryString = "/" + attribute;
        if (isNavigationProperty) {
            queryString += "/$ref";
        }
        var req = this.getRequest("DELETE", entitySet + "(" + id.value + ")" + queryString, null);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send();
        });
    };
    /**
     * Associate two records
     * @param entitySet Type of entity for primary record
     * @param id Id of primary record
     * @param relationship Schema name of relationship
     * @param relatedEntitySet Type of entity for secondary record
     * @param relatedEntityId Id of secondary record
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.associate = function (entitySet, id, relationship, relatedEntitySet, relatedEntityId, queryOptions) {
        var _this = this;
        var req = this.getRequest("POST", entitySet + "(" + id.value + ")/" + relationship + "/$ref", queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            var related = {
                "@odata.id": _this.getClientUrl(relatedEntitySet + "(" + relatedEntityId.value + ")")
            };
            req.send(JSON.stringify(related));
        });
    };
    /**
     * Disassociate two records
     * @param entitySet Type of entity for primary record
     * @param id  Id of primary record
     * @param relationship Schema name of relationship
     * @param relatedEntitySet Type of entity for secondary record
     * @param relatedEntityId Id of secondary record. Only needed for collection-valued navigation properties
     */
    WebApiBase.prototype.disassociate = function (entitySet, id, relationship, relatedEntitySet, relatedEntityId) {
        var queryString;
        if (relatedEntityId != null) {
            queryString = relationship + "(" + relatedEntityId.value + ")/$ref";
        }
        else {
            queryString = relationship + "/$ref";
        }
        var req = this.getRequest("DELETE", entitySet + "(" + id.value + ")/" + queryString, null);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send();
        });
    };
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.boundAction = function (entitySet, id, actionName, inputs, queryOptions) {
        var req = this.getRequest("POST", entitySet + "(" + id.value + ")/Microsoft.Dynamics.CRM." + actionName, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    }
                    else if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    };
    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.unboundAction = function (actionName, inputs, queryOptions) {
        var req = this.getRequest("POST", actionName, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    }
                    else if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    };
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.boundFunction = function (entitySet, id, functionName, inputs, queryOptions) {
        var queryString = entitySet + "(" + id.value + ")/Microsoft.Dynamics.CRM." + functionName + "(";
        queryString = this.getFunctionInputs(queryString, inputs);
        var req = this.getRequest("GET", queryString, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    }
                    else if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    };
    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.unboundFunction = function (functionName, inputs, queryOptions) {
        var queryString = functionName + "(";
        queryString = this.getFunctionInputs(queryString, inputs);
        var req = this.getRequest("GET", queryString, queryOptions);
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(JSON.parse(req.response));
                    }
                    else if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            inputs != null ? req.send(JSON.stringify(inputs)) : req.send();
        });
    };
    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     * @param impersonateUser Impersonate another user
     */
    WebApiBase.prototype.batchOperation = function (batchId, changeSetId, changeSets, batchGets, queryOptions) {
        var req = this.getRequest("POST", "$batch", queryOptions, "multipart/mixed;boundary=batch_" + batchId);
        // build post body
        var body = [];
        if (changeSets.length > 0) {
            body.push("--batch_" + batchId);
            body.push("Content-Type: multipart/mixed;boundary=changeset_" + changeSetId);
            body.push("");
        }
        // push change sets to body
        for (var i = 0; i < changeSets.length; i++) {
            body.push("--changeset_" + changeSetId);
            body.push("Content-Type: application/http");
            body.push("Content-Transfer-Encoding:binary");
            body.push("Content-ID: " + (i + 1));
            body.push("");
            body.push("POST " + this.getClientUrl(changeSets[i].queryString) + " HTTP/1.1");
            body.push("Content-Type: application/json;type=entry");
            body.push("");
            body.push(JSON.stringify(changeSets[i].entity));
        }
        if (changeSets.length > 0) {
            body.push("--changeset_" + changeSetId + "--");
            body.push("");
        }
        // push get requests to body
        for (var _i = 0, batchGets_1 = batchGets; _i < batchGets_1.length; _i++) {
            var get = batchGets_1[_i];
            body.push("--batch_" + batchId);
            body.push("Content-Type: application/http");
            body.push("Content-Transfer-Encoding:binary");
            body.push("");
            body.push("GET " + this.getClientUrl(get) + " HTTP/1.1");
            body.push("Accept: application/json");
        }
        if (batchGets.length > 0) {
            body.push("");
        }
        body.push("--batch_" + batchId + "--");
        return new Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 200) {
                        resolve(req.response);
                    }
                    else if (req.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            req.send(body.join("\r\n"));
        });
    };
    WebApiBase.prototype.getRequest = function (method, queryString, queryOptions, contentType, needsUrl) {
        if (contentType === void 0) { contentType = "application/json; charset=utf-8"; }
        if (needsUrl === void 0) { needsUrl = true; }
        var url;
        if (needsUrl) {
            url = this.getClientUrl(queryString);
        }
        else {
            url = queryString;
        }
        // build XMLHttpRequest
        var request = new XMLHttpRequest();
        request.open(method, url, true);
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", contentType);
        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");
        request.setRequestHeader("Cache-Control", "no-cache");
        if (queryOptions != null && typeof (queryOptions) !== "undefined") {
            request.setRequestHeader("Prefer", this.getPreferHeader(queryOptions));
            if (queryOptions.impersonateUser != null) {
                request.setRequestHeader("MSCRMCallerID", queryOptions.impersonateUser.value);
            }
        }
        if (this.accessToken != null) {
            request.setRequestHeader("Authorization", "Bearer " + this.accessToken);
        }
        return request;
    };
    WebApiBase.prototype.getPreferHeader = function (queryOptions) {
        var prefer = [];
        // add max page size to prefer request header
        if (queryOptions.maxPageSize) {
            prefer.push("odata.maxpagesize=" + queryOptions.maxPageSize);
        }
        // add formatted values to prefer request header
        if (queryOptions.includeFormattedValues && queryOptions.includeLookupLogicalNames &&
            queryOptions.includeAssociatedNavigationProperties) {
            prefer.push("odata.include-annotations=\"*\"");
        }
        else {
            var preferExtra = [
                queryOptions.includeFormattedValues ? "OData.Community.Display.V1.FormattedValue" : "",
                queryOptions.includeLookupLogicalNames ? "Microsoft.Dynamics.CRM.lookuplogicalname" : "",
                queryOptions.includeAssociatedNavigationProperties ? "Microsoft.Dynamics.CRM.associatednavigationproperty" : "",
            ].filter(function (v, i) {
                return v !== "";
            }).join(",");
            prefer.push("odata.include-annotations=\"" + preferExtra + "\"");
        }
        return prefer.join(",");
    };
    WebApiBase.prototype.getFunctionInputs = function (queryString, inputs) {
        if (inputs == null) {
            return queryString + ")";
        }
        var aliases = "?";
        for (var i = 0; i < inputs.length; i++) {
            queryString += inputs[i].name;
            if (inputs[i].alias) {
                queryString += "=@" + inputs[i].alias + ",";
                aliases += "@" + inputs[i].alias + "=" + inputs[i].value;
            }
            else {
                queryString += "=" + inputs[i].value + ",";
            }
        }
        queryString = queryString.substr(0, queryString.length - 1) + ")";
        if (aliases !== "?") {
            queryString += aliases;
        }
        return queryString;
    };
    return WebApiBase;
}());
exports.WebApiBase = WebApiBase;
var WebApi = /** @class */ (function (_super) {
    __extends(WebApi, _super);
    function WebApi() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return WebApi;
}(WebApiBase));
exports.WebApi = WebApi;
var WebApiNode = /** @class */ (function (_super) {
    __extends(WebApiNode, _super);
    function WebApiNode() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return WebApiNode;
}(WebApiBase));
