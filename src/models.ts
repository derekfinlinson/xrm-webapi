export class Guid {
    public value: string;

    constructor(value: string) {
        value = value.replace(/[{}]/g, "");

        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value)) {
            this.value = value.toUpperCase();
        } else {
            throw Error(`Id ${value} is not a valid GUID`);
        }
    }

    areEqual(compare: Guid): boolean {
        if (this === null || compare === null || this === undefined || compare === undefined) {
            return false;
        }

        return this.value.toLowerCase() === compare.value.toLowerCase();
    }
}

export interface QueryOptions {
    includeFormattedValues?: boolean;
    includeLookupLogicalNames?: boolean;
    includeAssociatedNavigationProperties?: boolean;
    maxPageSize?: number;
    impersonateUser?: Guid;
    representation?: boolean;
}

export interface Entity {
    [propName: string]: any;
}

export interface RetrieveMultipleResponse {
    value: Entity[];
    "@odata.nextlink": string;
}

export interface ChangeSet {
    queryString: string;
    entity: object;
}

export interface FunctionInput {
    name: string;
    value: string;
    alias?: string;
}