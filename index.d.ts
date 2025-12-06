import {
  Schema,
	ParseWithSchemaOptions,
	ParseWithoutSchemaOptions,
	ParsedObjectsResult,
	Row
} from './types.d.js';

export {
	Schema,
	ParsedObjectsResult,
	Error,
	CellValue,
	Row,
	Integer,
	Email,
	URL
} from './types.d.js';

export function parseExcelDate(excelSerialDate: number): typeof Date;

type Input = File | Blob | ArrayBuffer;

export function readXlsxFile<T extends object>(input: Input, options: ParseWithSchemaOptions<T>): Promise<ParsedObjectsResult<T>>;
export function readXlsxFile(input: Input, options?: ParseWithoutSchemaOptions): Promise<Row[]>;

export function readSheetNames(input: Input): Promise<string[]>;

export function convertToJsonWithSchema<T extends object>(result: ParsedObjectsResult<T>, schema: Schema<T>): { rows: any[] };

export default readXlsxFile;
