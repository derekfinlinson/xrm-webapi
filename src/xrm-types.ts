export interface FunctionInput {
    name: string;
    value: string;
    alias?: string;
}

export class Guid {
    value: string;

    constructor(value: string) {
        value.replace(/[{}]/g, "");
        
        if (/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(value)) {
            this.value = value.toUpperCase();
        } else {
            throw Error(`Id ${value} is not a valid GUID`);
        }
    }
}

export interface Attribute {
    name: string;
    value?: any
}

export interface Entity {
    id?: Guid;
    attributes: Array<Attribute>;
}