import { WebApiConfig } from "./webapi";
import { QueryOptions } from "./models";
import { RequestOptions } from "https";
import { request } from "https";
import { URL } from "url";
import { ClientRequest } from "http";

export interface WebApiRequestResult {
    error: boolean;
    response: any;
    headers?: any;
}

export interface WebApiRequestConfig {
    method: string;
    contentType: string;
    body?: any;
    queryString: string;
}

export class WebApiRequest {
    private _config: WebApiConfig;

    constructor(config: WebApiConfig) {
        this._config = config;
    }

    public submitRequest(config: WebApiRequestConfig, queryOptions: QueryOptions, callback: (result: WebApiRequestResult) => void): void {
        if (typeof window !== "undefined" && typeof window.document !== "undefined") {
            const req: XMLHttpRequest = new XMLHttpRequest();

            req.open(config.method, encodeURI(`${this._config.url}/${config.queryString}`), true);

            const headers: any = this.getHeaders(config, queryOptions);

            for (let header in headers) {
                if (headers.hasOwnProperty(header)) {
                    req.setRequestHeader(header, headers[header]);
                }
            }

            req.onreadystatechange = () => {
                if (req.readyState === 4 /* complete */) {
                    req.onreadystatechange = null;

                    if ((req.status >= 200) && (req.status < 300)) {
                        callback({ error: false, response: req.response, headers: req.getAllResponseHeaders() });
                    } else {
                        callback({ error: true, response: req.response, headers: req.getAllResponseHeaders() });
                    }
                }
            };

            if (config.body != null) {
                req.send(config.body);
            } else {
                req.send();
            }
        } else {
            this.submitRequestNode(config, queryOptions, callback);
        }
    }

    private submitRequestNode(config: WebApiRequestConfig, queryOptions: QueryOptions,
        callback: (result: WebApiRequestResult) => void): void {
        const url: URL = new URL(`${this._config.url}/${config.queryString}`);

        const headers: any = this.getHeaders(config, queryOptions);

        const options: RequestOptions = {
            hostname: url.hostname,
            path: `${url.pathname}${url.search}`,
            method: config.method,
            headers: headers
        };

        if (config.body) {
            options.headers["Content-Length"] = config.body.length;
        }

        const req: ClientRequest = request(options,
            (result) => {
                let body: string = "";

                result.setEncoding("utf8");

                result.on("data", (chunk) => {
                    body += chunk;
                });

                result.on("end", () => {
                    if ((result.statusCode >= 200) && (result.statusCode < 300)) {
                        callback({ error: false, response: body, headers: result.headers });
                    } else {
                        callback({ error: true, response: body, headers: result.headers });
                    }
                });
            }
        );

        req.on("error", (error) => {
			callback({ error: true, response: error });
        });

		if (config.body != null) {
			req.write(config.body);
        }

		req.end();
    }

    private getHeaders(config: WebApiRequestConfig, queryOptions?: QueryOptions): any {
        // Get axios config
        const headers: any = {};

        headers.Accept = "application/json";
        headers["OData-MaxVersion"] = "4.0";
        headers["OData-Version"] = "4.0";
        headers["Content-Type"] = config.contentType;

        if (this._config.accessToken != null) {
            headers.Authorization = `Bearer ${this._config.accessToken}`;
        }

        if (queryOptions != null && typeof(queryOptions) !== "undefined") {
            headers.Prefer = this.getPreferHeader(queryOptions);

            if (queryOptions.impersonateUser != null) {
                headers.MSCRMCallerID = queryOptions.impersonateUser.value;
            }
        }

        return headers;
    }

    private getPreferHeader(queryOptions: QueryOptions): string {
        let prefer: string[] = [];

        // add max page size to prefer request header
        if (queryOptions.maxPageSize) {
            prefer.push(`odata.maxpagesize=${queryOptions.maxPageSize}`);
        }

        // add formatted values to prefer request header
        if (queryOptions.representation) {
            prefer.push("return=representation");
        } else if (queryOptions.includeFormattedValues && queryOptions.includeLookupLogicalNames &&
            queryOptions.includeAssociatedNavigationProperties) {
            prefer.push("odata.include-annotations=\"*\"");
        } else {
            const preferExtra: string = [
                queryOptions.includeFormattedValues ? "OData.Community.Display.V1.FormattedValue" : "",
                queryOptions.includeLookupLogicalNames ? "Microsoft.Dynamics.CRM.lookuplogicalname" : "",
                queryOptions.includeAssociatedNavigationProperties ? "Microsoft.Dynamics.CRM.associatednavigationproperty" : "",
            ].filter((v) => {
                return v !== "";
            }).join(",");

            prefer.push("odata.include-annotations=\"" + preferExtra + "\"");
        }

        return prefer.join(",");
    }
}