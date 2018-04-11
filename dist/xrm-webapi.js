"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var axios_1 = require("axios");
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
var WebApi = /** @class */ (function () {
    /**
     * Constructor
     * @param version Version must be 8.0, 8.1, 8.2 or 9.0
     * @param accessToken Optional access token if using from outside Dynamics 365
     * @param url Optional url if using from outside Dynamics 365
     */
    function WebApi(version, accessToken, url) {
        this.version = version;
        this.accessToken = accessToken;
        this.url = url;
        axios_1.default.defaults.headers.common["Accept"] = "application/json";
        axios_1.default.defaults.headers.common["OData-MaxVersion"] = "4.0";
        axios_1.default.defaults.headers.common["OData-Version"] = "4.0";
        if (this.accessToken != null) {
            axios_1.default.defaults.headers.common["Authorization"] = "Bearer " + this.accessToken;
        }
    }
    /**
     * Get the OData URL
     * @param queryString Query string to append to URL. Defaults to a blank string
     */
    WebApi.prototype.getClientUrl = function (queryString) {
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
    WebApi.prototype.retrieve = function (entitySet, id, queryString, queryOptions) {
        if (queryString != null && !/^[?]/.test(queryString)) {
            queryString = "?" + queryString;
        }
        var query = (queryString != null) ? entitySet + "(" + id.value + ")" + queryString : entitySet + "(" + id.value + ")";
        var config = this.getRequestConfig("GET", query, queryOptions);
        return axios_1.default(config);
    };
    /**
     * Retrieve multiple records from CRM
     * @param entitySet Type of entity to retrieve
     * @param queryString OData query string parameters
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.retrieveMultiple = function (entitySet, queryString, queryOptions) {
        if (queryString != null && !/^[?]/.test(queryString)) {
            queryString = "?" + queryString;
        }
        var query = (queryString != null) ? entitySet + queryString : entitySet;
        var config = this.getRequestConfig("GET", query, queryOptions);
        return axios_1.default(config);
    };
    /**
     * Retrieve next page from a retrieveMultiple request
     * @param query Query from the @odata.nextlink property of a retrieveMultiple
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.getNextPage = function (query, queryOptions) {
        var config = this.getRequestConfig("GET", query, queryOptions, null, false);
        return axios_1.default(config);
    };
    /**
     * Create a record in CRM
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.create = function (entitySet, entity, queryOptions) {
        var config = this.getRequestConfig("POST", entitySet, queryOptions);
        config.transformResponse = function (data, headers) {
            var uri = headers["odata-entityid"];
            var start = uri.indexOf("(") + 1;
            var end = uri.indexOf(")", start);
            var id = uri.substring(start, end);
            data = {
                id: new Guid(id),
                uri: uri
            };
            return data;
        };
        config.data = JSON.stringify(entity);
        return axios_1.default(config);
    };
    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param entity Entity to create
     * @param select Select odata query parameter
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.createWithReturnData = function (entitySet, entity, select, queryOptions) {
        if (select != null && !/^[?]/.test(select)) {
            select = "?" + select;
        }
        // set reprensetation
        if (queryOptions == null) {
            queryOptions = {};
        }
        queryOptions.representation = true;
        var config = this.getRequestConfig("POST", entitySet + select, queryOptions);
        config.data = JSON.stringify(entity);
        return axios_1.default(config);
    };
    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.update = function (entitySet, id, entity, queryOptions) {
        var config = this.getRequestConfig("PATCH", entitySet + "(" + id.value + ")", queryOptions);
        config.data = JSON.stringify(entity);
        return axios_1.default(config);
    };
    /**
     * Create a record in CRM and return data
     * @param entitySet Type of entity to create
     * @param id Id of record to update
     * @param entity Entity fields to update
     * @param select Select odata query parameter
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.updateWithReturnData = function (entitySet, id, entity, select, queryOptions) {
        if (select != null && !/^[?]/.test(select)) {
            select = "?" + select;
        }
        // set representation
        if (queryOptions == null) {
            queryOptions = {};
        }
        queryOptions.representation = true;
        var config = this.getRequestConfig("PATCH", entitySet + "(" + id.value + ")" + select, queryOptions);
        config.data = JSON.stringify(entity);
        return axios_1.default(config);
    };
    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.updateProperty = function (entitySet, id, attribute, value, queryOptions) {
        var config = this.getRequestConfig("PUT", entitySet + "(" + id.value + ")/" + attribute, queryOptions);
        config.data = JSON.stringify({ value: value });
        return axios_1.default(config);
    };
    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    WebApi.prototype.delete = function (entitySet, id) {
        var config = this.getRequestConfig("DELETE", entitySet + "(" + id.value + ")", null);
        return axios_1.default(config);
    };
    /**
     * Delete a property from a record in CRM. Non navigation properties only
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to delete
     */
    WebApi.prototype.deleteProperty = function (entitySet, id, attribute) {
        var queryString = "/" + attribute;
        var config = this.getRequestConfig("DELETE", entitySet + "(" + id.value + ")" + queryString, null);
        return axios_1.default(config);
    };
    /**
     * Associate two records
     * @param entitySet Type of entity for primary record
     * @param id Id of primary record
     * @param relationship Schema name of relationship
     * @param relatedEntitySet Type of entity for secondary record
     * @param relatedEntityId Id of secondary record
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.associate = function (entitySet, id, relationship, relatedEntitySet, relatedEntityId, queryOptions) {
        var config = this.getRequestConfig("POST", entitySet + "(" + id.value + ")/" + relationship + "/$ref", queryOptions);
        var related = {
            "@odata.id": this.getClientUrl(relatedEntitySet + "(" + relatedEntityId.value + ")")
        };
        config.data = JSON.stringify(related);
        return axios_1.default(config);
    };
    /**
     * Disassociate two records
     * @param entitySet Type of entity for primary record
     * @param id  Id of primary record
     * @param property Schema name of property or relationship
     * @param relatedEntityId Id of secondary record. Only needed for collection-valued navigation properties
     */
    WebApi.prototype.disassociate = function (entitySet, id, property, relatedEntityId) {
        var queryString = property;
        if (relatedEntityId != null) {
            queryString += "(" + relatedEntityId.value + ")";
        }
        queryString += "/$ref";
        var config = this.getRequestConfig("DELETE", entitySet + "(" + id.value + ")/" + queryString, null);
        return axios_1.default(config);
    };
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.boundAction = function (entitySet, id, actionName, inputs, queryOptions) {
        var config = this.getRequestConfig("POST", entitySet + "(" + id.value + ")/Microsoft.Dynamics.CRM." + actionName, queryOptions);
        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }
        return axios_1.default(config);
    };
    /**
     * Execute a default or custom unbound action in CRM
     * @param actionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.unboundAction = function (actionName, inputs, queryOptions) {
        var config = this.getRequestConfig("POST", actionName, queryOptions);
        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }
        return axios_1.default(config);
    };
    /**
     * Execute a default or custom bound action in CRM
     * @param entitySet Type of entity to run the action against
     * @param id Id of record to run the action against
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.boundFunction = function (entitySet, id, functionName, inputs, queryOptions) {
        var queryString = entitySet + "(" + id.value + ")/Microsoft.Dynamics.CRM." + functionName + "(";
        queryString = this.getFunctionInputs(queryString, inputs);
        var config = this.getRequestConfig("GET", queryString, queryOptions);
        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }
        return axios_1.default(config);
    };
    /**
     * Execute an unbound function in CRM
     * @param functionName Name of the action to run
     * @param inputs Any inputs required by the action
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.unboundFunction = function (functionName, inputs, queryOptions) {
        var queryString = functionName + "(";
        queryString = this.getFunctionInputs(queryString, inputs);
        var config = this.getRequestConfig("GET", queryString, queryOptions);
        if (inputs != null) {
            config.data = JSON.stringify(inputs);
        }
        return axios_1.default(config);
    };
    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     * @param queryOptions Various query options for the query
     */
    WebApi.prototype.batchOperation = function (batchId, changeSetId, changeSets, batchGets, queryOptions) {
        var config = this.getRequestConfig("POST", "$batch", queryOptions, "multipart/mixed;boundary=batch_" + batchId);
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
        config.data = body.join("\r\n");
        return axios_1.default(config);
    };
    WebApi.prototype.getRequestConfig = function (method, queryString, queryOptions, contentType, needsUrl) {
        if (contentType === void 0) { contentType = "application/json; charset=utf-8"; }
        if (needsUrl === void 0) { needsUrl = true; }
        var url;
        if (needsUrl) {
            url = this.getClientUrl(queryString);
        }
        else {
            url = queryString;
        }
        // Get axios config
        var config = {
            url: url,
            method: method,
            headers: {
                "Content-Type": contentType
            }
        };
        if (queryOptions != null && typeof (queryOptions) !== "undefined") {
            config.headers["Prefer"] = this.getPreferHeader(queryOptions);
            if (queryOptions.impersonateUser != null) {
                config.headers["MSCRMCallerID"] = queryOptions.impersonateUser.value;
            }
        }
        return config;
    };
    WebApi.prototype.getPreferHeader = function (queryOptions) {
        var prefer = [];
        // add max page size to prefer request header
        if (queryOptions.maxPageSize) {
            prefer.push("odata.maxpagesize=" + queryOptions.maxPageSize);
        }
        // add formatted values to prefer request header
        if (queryOptions.representation) {
            prefer.push("return=representation");
        }
        else if (queryOptions.includeFormattedValues && queryOptions.includeLookupLogicalNames &&
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
    WebApi.prototype.getFunctionInputs = function (queryString, inputs) {
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
    return WebApi;
}());
exports.WebApi = WebApi;
