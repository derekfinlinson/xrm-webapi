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
    /**
     * Constructor. Version should be 8.0, 8.1 or 8.2
     */
    function WebApi(version) {
        this.version = version;
    }
    WebApi.prototype.getRequest = function (method, queryString, contentType) {
        if (contentType === void 0) { contentType = "application/json; charset=utf-8"; }
        var url = this.getClientUrl(queryString);
        var request = new XMLHttpRequest();
        request.open(method, url, true);
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", contentType);
        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");
        request.setRequestHeader("Cache-Control", "no-cache");
        return request;
    };
    WebApi.prototype.getFunctionInputs = function (queryString, inputs) {
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
    WebApi.prototype.getPreferHeader = function (formattedValues, lookupLogicalNames, associatedNavigationProperties, maxPageSize) {
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
     * Get the OData URL
     * @param queryString Query string to append to URL. Defaults to a blank string
     */
    WebApi.prototype.getClientUrl = function (queryString) {
        if (queryString === void 0) { queryString = ""; }
        var context = typeof GetGlobalContext != "undefined" ? GetGlobalContext() : Xrm.Page.context;
        var url = context.getClientUrl() + ("/api/data/v" + this.version + "/") + queryString;
        return url;
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
    WebApi.prototype.retrieve = function (entityType, id, queryString, includeFormattedValues, includeLookupLogicalNames, includeAssociatedNavigationProperties) {
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
    WebApi.prototype.retrieveMultiple = function (entitySet, queryString, includeFormattedValues, includeLookupLogicalNames, includeAssociatedNavigationProperties, maxPageSize) {
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
    WebApi.prototype.create = function (entitySet, entity) {
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
    WebApi.prototype.update = function (entitySet, id, entity) {
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
    WebApi.prototype.updateProperty = function (entitySet, id, attribute) {
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
    WebApi.prototype.delete = function (entitySet, id) {
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
    WebApi.prototype.deleteProperty = function (entitySet, id, attribute, isNavigationProperty) {
        var queryString = "/" + attribute.name;
        if (isNavigationProperty) {
            queryString += "/$ref";
        }
        var req = this.getRequest("DELETE", entitySet + "(" + id.value + ")" + queryString);
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
    WebApi.prototype.boundAction = function (entitySet, id, actionName, inputs) {
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
    WebApi.prototype.unboundAction = function (actionName, inputs) {
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
    WebApi.prototype.boundFunction = function (entitySet, id, functionName, inputs) {
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
    WebApi.prototype.unboundFunction = function (functionName, inputs) {
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
    /**
     * Execute a batch operation in CRM
     * @param batchId Unique batch id for the operation
     * @param changeSetId Unique change set id for any changesets in the operation
     * @param changeSets Array of change sets (create or update) for the operation
     * @param batchGets Array of get requests for the operation
     */
    WebApi.prototype.batchOperation = function (batchId, changeSetId, changeSets, batchGets) {
        var req = this.getRequest("POST", "$batch", "multipart/mixed;boundary=batch_" + batchId);
        // Build post body
        var body = [
            "--batch_" + batchId,
            "Content-Type: multipart/mixed;boundary=changeset_" + changeSetId,
            ""
        ];
        var _loop_1 = function (i) {
            body.push("--changeset_" + changeSetId);
            body.push("Content-Type: application/http");
            body.push("Content-Transfer-Encoding:binary");
            body.push("Content-ID: " + (i + 1));
            body.push("");
            body.push("POST " + this_1.getClientUrl(changeSets[i].queryString) + " HTTP/1.1");
            body.push("Content-Type: application/json;type=entry");
            body.push("");
            var attributes = {};
            changeSets[i].object.attributes.forEach(function (attribute) {
                attributes[attribute.name] = attribute.value;
            });
            body.push(JSON.stringify(attributes));
        };
        var this_1 = this;
        // Push change sets to body
        for (var i = 0; i < changeSets.length; i++) {
            _loop_1(i);
        }
        if (changeSets.length > 0) {
            body.push("--changeset_" + changeSetId + "--");
            body.push("");
        }
        // Push get requests to body
        for (var _i = 0, batchGets_1 = batchGets; _i < batchGets_1.length; _i++) {
            var get = batchGets_1[_i];
            body.push("--batch_" + batchId);
            body.push("Content-Type: application/http");
            body.push("Content-Transfer-Encoding:binary");
            body.push("");
            body.push("GET " + this.getClientUrl(get) + " HTTP/1.1");
            body.push("Accept: application/json");
        }
        body.push("--batch_" + batchId + "--");
        return new es6_promise_1.Promise(function (resolve, reject) {
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
    return WebApi;
}());
exports.WebApi = WebApi;
