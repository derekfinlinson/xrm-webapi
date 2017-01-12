"use strict";
var es6_promise_1 = require("es6-promise");
var Guid = (function () {
    function Guid(value) {
        value = value.replace(/[{}]/g, "");
        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value)) {
            this.value = value.toUpperCase();
        }
        else {
            throw Error("Id " + value + " is not a valid GUID");
        }
    }
    return Guid;
}());
exports.Guid = Guid;
var WebApi = (function () {
    function WebApi() {
    }
    WebApi.getRequest = function (method, queryString) {
        var context = typeof GetGlobalContext != "undefined" ? GetGlobalContext() : Xrm.Page.context;
        var url = context.getClientUrl() + "/api/data/v8.1/" + queryString;
        var request = new XMLHttpRequest();
        request.open(method, url, true);
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");
        return request;
    };
    WebApi.getFunctionInputs = function (queryString, inputs) {
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
        queryString += queryString.substr(0, queryString.length - 1) + ")";
        if (aliases != "?") {
            queryString += aliases;
        }
        return queryString;
    };
    WebApi.getPreferHeader = function (formattedValues, lookupLogicalNames, associatedNavigationProperties, maxPageSize) {
        var prefer = [];
        if (maxPageSize) {
            prefer.push("odata.maxpagesize=" + maxPageSize);
        }
        if (formattedValues && lookupLogicalNames & associatedNavigationProperties) {
            prefer.push('odata.include-annotations="*"');
        }
        else {
            var preferExtra = [
                formattedValues ? "OData.Community.Display.V1.FormattedValue" : "",
                lookupLogicalNames ? "Microsoft.Dynamics.CRM.lookuplogicalname" : "",
                associatedNavigationProperties ? "Microsoft.Dynamics.CRM.associatednavigationproperty" : ""
            ].filter(function (v, i) { return v != ""; }).join(",");
            prefer.push('odata.include-annotations="' + preferExtra + '"');
        }
        return prefer.join(",");
    };
    /**
     * Retrieve a record from CRM
     * @param entityType Type of entity to retrieve
     * @param id Id of record to retrieve
     * @param queryString OData query string parameters
     * @param includeFormattedValues Include formatted values in results
     * @param includeLookupLogicalNames Include lookup logical names in results
     * @param includeAssociatedNavigationProperty Include associated navigation property in results
     */
    WebApi.retrieve = function (entityType, id, queryString, includeFormattedValues, includeLookupLogicalNames, includeAssociatedNavigationProperties) {
        if (includeFormattedValues === void 0) { includeFormattedValues = false; }
        if (includeLookupLogicalNames === void 0) { includeLookupLogicalNames = false; }
        if (includeAssociatedNavigationProperties === void 0) { includeAssociatedNavigationProperties = false; }
        if (queryString != null && !/^[?]/.test(queryString))
            queryString = "?" + queryString;
        var query = (queryString != null) ? entityType + "(" + id.value + ")" + queryString : entityType + "(" + id.value + ")";
        var req = this.getRequest("GET", query);
        if (includeFormattedValues || includeLookupLogicalNames || includeAssociatedNavigationProperties) {
            req.setRequestHeader("Prefer", this.getPreferHeader(includeFormattedValues, includeLookupLogicalNames, includeAssociatedNavigationProperties));
        }
        return new es6_promise_1.Promise(function (resolve, reject) {
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
     * @param includeFormattedValues Include formatted values in results
     * @param includeLookupLogicalNames Include lookup logical names in results
     * @param includeAssociatedNavigationProperty Include associated navigation property in results
     * @param maxPageSize Records per page to return
     */
    WebApi.retrieveMultiple = function (entitySet, queryString, includeFormattedValues, includeLookupLogicalNames, includeAssociatedNavigationProperties, maxPageSize) {
        if (includeFormattedValues === void 0) { includeFormattedValues = false; }
        if (includeLookupLogicalNames === void 0) { includeLookupLogicalNames = false; }
        if (includeAssociatedNavigationProperties === void 0) { includeAssociatedNavigationProperties = false; }
        if (queryString != null && !/^[?]/.test(queryString))
            queryString = "?" + queryString;
        var query = (queryString != null) ? entitySet + queryString : entitySet;
        var req = this.getRequest("GET", query);
        if (includeFormattedValues || includeLookupLogicalNames || includeAssociatedNavigationProperties) {
            req.setRequestHeader("Prefer", this.getPreferHeader(includeFormattedValues, includeLookupLogicalNames, includeAssociatedNavigationProperties, maxPageSize));
        }
        return new es6_promise_1.Promise(function (resolve, reject) {
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
     */
    WebApi.create = function (entitySet, entity) {
        var req = this.getRequest("POST", entitySet);
        return new es6_promise_1.Promise(function (resolve, reject) {
            req.onreadystatechange = function () {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;
                    if (req.status === 204) {
                        var uri = req.getResponseHeader("OData-EntityId");
                        var start = uri.indexOf('(') + 1;
                        var end = uri.indexOf(')', start);
                        var id = uri.substring(start, end);
                        var createdEntity = {
                            id: new Guid(id),
                            uri: uri
                        };
                        resolve(createdEntity);
                    }
                    else {
                        reject(JSON.parse(req.response).error);
                    }
                }
            };
            var attributes = {};
            entity.attributes.forEach(function (attribute) {
                attributes[attribute.name] = attribute.value;
            });
            req.send(JSON.stringify(attributes));
        });
    };
    /**
     * Update a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param entity Entity fields to update
     */
    WebApi.update = function (entitySet, id, entity) {
        var req = this.getRequest("PATCH", entitySet + "(" + id.value + ")");
        return new es6_promise_1.Promise(function (resolve, reject) {
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
            var attributes = {};
            entity.attributes.forEach(function (attribute) {
                attributes[attribute.name] = attribute.value;
            });
            req.send(JSON.stringify(attributes));
        });
    };
    /**
     * Update a single property of a record in CRM
     * @param entitySet Type of entity to update
     * @param id Id of record to update
     * @param attribute Attribute to update
     */
    WebApi.updateProperty = function (entitySet, id, attribute) {
        var req = this.getRequest("PUT", entitySet + "(" + id.value + ")/" + attribute.name);
        return new es6_promise_1.Promise(function (resolve, reject) {
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
            req.send(JSON.stringify({ "value": attribute.value }));
        });
    };
    /**
     * Delete a record from CRM
     * @param entitySet Type of entity to delete
     * @param id Id of record to delete
     */
    WebApi.delete = function (entitySet, id) {
        var req = this.getRequest("DELETE", entitySet + "(" + id.value + ")");
        return new es6_promise_1.Promise(function (resolve, reject) {
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
    WebApi.deleteProperty = function (entitySet, id, attribute) {
        var req = this.getRequest("DELETE", entitySet + "(" + id.value + ")/" + attribute.name);
        return new es6_promise_1.Promise(function (resolve, reject) {
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
     */
    WebApi.boundAction = function (entitySet, id, actionName, inputs) {
        var req = this.getRequest("POST", entitySet + "(" + id.value + ")/Microsoft.Dynamics.CRM." + actionName);
        return new es6_promise_1.Promise(function (resolve, reject) {
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
     */
    WebApi.unboundAction = function (actionName, inputs) {
        var req = this.getRequest("POST", actionName);
        return new es6_promise_1.Promise(function (resolve, reject) {
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
     */
    WebApi.boundFunction = function (entitySet, id, functionName, inputs) {
        var queryString = entitySet + "(" + id.value + ")/Microsoft.Dynamics.CRM." + functionName + "(";
        queryString = this.getFunctionInputs(queryString, inputs);
        var req = this.getRequest("GET", queryString);
        return new es6_promise_1.Promise(function (resolve, reject) {
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
     */
    WebApi.unboundFunction = function (functionName, inputs) {
        var queryString = functionName + "(";
        queryString = this.getFunctionInputs(queryString, inputs);
        var req = this.getRequest("GET", queryString);
        return new es6_promise_1.Promise(function (resolve, reject) {
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
    return WebApi;
}());
exports.WebApi = WebApi;
