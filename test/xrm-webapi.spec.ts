import { WebApi, WebApiConfig } from "../src/xrm-webapi";

describe("WebApi", () => {
    let api: WebApi;

    beforeEach(() => {
        const config: WebApiConfig = {
            version: "8.2"
        };

        api = new WebApi(config);
    });

    test("create an account", async () => {
    });

    test("create an account with returned data", async () => {
    });
});