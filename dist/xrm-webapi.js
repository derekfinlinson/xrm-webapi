"use strict";
var es6_promise_1 = require("es6-promise");
var WebApi = (function () {
    function WebApi() {
    }
    WebApi.getRequest = function (method, entitySet, queryString) {
        var context = typeof GetGlobalContext != "undefined" ? GetGlobalContext() : Xrm.Page.context;
        var url = context.getClientUrl() + "/api/data/v8.0/" + entitySet + queryString;
        this.request = new XMLHttpRequest();
        this.request.open(method, url, true);
        this.request.setRequestHeader("Accept", "application/json");
        this.request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        this.request.setRequestHeader("OData-MaxVersion", "4.0");
        this.request.setRequestHeader("OData-Version", "4.0");
    };
    WebApi.retrieve = function (entityType, id, queryString) {
        var _this = this;
        if (queryString != null && !/^[?]/.test(queryString))
            queryString = "?" + queryString;
        id = id.replace(/[{}]/g, "");
        this.getRequest("GET", entityType, "(" + id + ")" + queryString);
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 200) {
                        resolve(JSON.parse(_this.request.response));
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            _this.request.send();
        });
    };
    WebApi.retrieveMultiple = function (entitySet, queryString) {
        var _this = this;
        if (queryString != null && !/^[?]/.test(queryString))
            queryString = "?" + queryString;
        this.getRequest("GET", entitySet, queryString);
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 200) {
                        resolve(JSON.parse(_this.request.response));
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            _this.request.send();
        });
    };
    WebApi.create = function (entitySet, entity) {
        var _this = this;
        this.getRequest("POST", entitySet);
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 200) {
                        resolve(_this.request.getResponseHeader("OData-EntityId"));
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            _this.request.send(entity);
        });
    };
    WebApi.update = function (entitySet, id, entity) {
        var _this = this;
        id = id.replace(/[{}]/g, "");
        this.getRequest("PATCH", entitySet, "(" + id + ")");
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            _this.request.send(JSON.stringify(entity));
        });
    };
    WebApi.updateProperty = function (entitySet, id, attribute, value) {
        var _this = this;
        id = id.replace(/[{}]/g, "");
        this.getRequest("PUT", entitySet, "(" + id + ")");
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            _this.request.send(JSON.stringify({ attribute: value }));
        });
    };
    WebApi.delete = function (entitySet, id) {
        var _this = this;
        id = id.replace(/[{}]/g, "");
        this.getRequest("DELETE", entitySet, "(" + id + ")");
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            _this.request.send();
        });
    };
    WebApi.deleteProperty = function (entitySet, id, attribute) {
        var _this = this;
        id = id.replace(/[{}]/g, "");
        this.getRequest("DELETE", entitySet, "(" + id + ")/" + attribute);
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            _this.request.send();
        });
    };
    WebApi.executeAction = function (entitySet, id, actionName, inputs) {
        var _this = this;
        id = id.replace(/[{}]/g, "");
        this.getRequest("POST", entitySet, "(" + id + ")/Microsoft.Dynamics.CRM." + actionName);
        return new es6_promise_1.Promise(function (resolve, reject) {
            _this.request.onreadystatechange = function () {
                if (_this.request.readyState === 4) {
                    _this.request.onreadystatechange = null;
                    if (_this.request.status === 200) {
                        resolve(JSON.parse(_this.request.response));
                    }
                    else if (_this.request.status === 204) {
                        resolve();
                    }
                    else {
                        reject(JSON.parse(_this.request.response).error);
                    }
                }
            };
            inputs != null ? _this.request.send(JSON.stringify(inputs)) : _this.request.send();
        });
    };
    return WebApi;
}());
exports.WebApi = WebApi;
