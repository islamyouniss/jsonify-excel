// See the discussion:
// https://github.com/catamphetamine/read-excel-file/issues/71

type BasicType =
	| string
	| number
	| boolean
	| typeof Date
	| 'Integer'
	| 'URL'
	| 'Email'

interface SchemaEntryBasic {
	prop: string;
	type: BasicType;
	oneOf?<T>: T[];
	required?: boolean;
	validate?<T>(value: T): void;
}

interface SchemaEntryParsed {
	prop: string;
	parse<T>: (value: string) => T?;
	oneOf?<T>: T[];
	required?: boolean;
	validate?<T>(value: T): void;
}

// Implementing recursive types in TypeScript:
// https://dev.to/busypeoples/notes-on-typescript-recursive-types-and-immutability-5ck1
interface SchemaEntryRecursive {
	prop: string;
	type: Record<string, SchemaEntry>;
	required?: boolean;
}

type SchemaEntry = SchemaEntryBasic | SchemaEntryParsed | SchemaEntryRecursive

export type Schema = Record<string, SchemaEntry>

export interface Error {
	error: string;
	row: number;
	column: number;
	value?: any;
	type?: SchemaEntry;
};

type Cell = string | number | boolean | typeof Date
export type Row = Cell[]

export interface ParsedObjectsResult {
	rows: object[];
	errors: Error[];
}

export interface ParseWithSchemaOptions {
	schema: Schema;
	transformData?: (rows: Row[]) => Row[];
	sheet?: number | string;
}

export interface ParseWithoutSchemaOptions {
	sheet?: number | string;
}