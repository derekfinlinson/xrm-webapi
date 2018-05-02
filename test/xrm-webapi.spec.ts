import { WebApi } from "../src/xrm-webapi";

describe('WebApi', () => {
    let api: WebApi;

    beforeEach(() => {
        api = new WebApi("8.2", "", "");
    });

    test('create an account', async () => {
        const account = {
            name: 'Test account'
        };

        const result = await api.create('accounts', account);

        expect(result).toBeDefined();
        expect(result.data).toBeDefined();
        expect(result.data.id).toBeDefined();
        expect(result.data.uri).toBeDefined();
    });

    test('create an account with returned data', async () => {
        const account = {
            name: 'Test account'
        };

        const result = await api.createWithReturnData('accounts', account, "$select=name");

        expect(result).toBeDefined();
        expect(result.data).toBeDefined();
        expect(result.data.name).toBeDefined();
        expect(result.data.name).toBe(account.name);
    });
});