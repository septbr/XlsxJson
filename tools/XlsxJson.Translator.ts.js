let tsTemplate = `export namespace xlsx {
    export type u8 = number;
    export type u16 = number;
    export type u32 = number;
    export type u64 = number;
    export type i8 = number;
    export type i16 = number;
    export type i32 = number;
    export type i64 = number;
    export type bool = boolean;
    export type str = string;

__SHEETS_DATA__

    type Sheets = {
__SHEETS_DATA_MAP__
    };

    let hdrs: { [P in keyof Sheets]: [string, number, boolean][] } = {} as any;
    let rowss: { [P in keyof Sheets]: any[][] } = {} as any;

    /**
     * 加载json
     * @param json JSON字符串 或 JSON对象
     */
    export function load(json: string | object): Error | undefined {
        let _hdrs = hdrs, _rowss = rowss;
        try {
            hdrs = {} as any, rowss = typeof json == 'object' ? json : JSON.parse(json);
            for (let sheet in rowss) {
                let hdr = rowss[sheet as keyof Sheets].shift()!;
                hdrs[sheet as keyof Sheets] = hdr.map(index => [
                    index[0],
                    index[1].indexOf(':') >= 0 ? 5 : index[1][0] == '[' ? 4 : index[1].endsWith(']') ? 3 : index[1] == 'bool' ? 2 : 1,
                    index[2] == 1
                ]);
                /** check data */
                rowss[sheet as keyof Sheets].forEach(row => make(sheet as keyof Sheets, row));
            }
        } catch (err) {
            hdrs = _hdrs, rowss = _rowss;
            return err;
        }
    }

    function make<T extends keyof Sheets>(sheet: T, row: any[]): Sheets[T] {
        let hdr = hdrs[sheet], data: any = {};
        hdr.forEach(([name, kind], index) => {
            let value = kind == 1 ? row[index] : kind == 2 ? !!row[index] : kind == 3 || kind == 4 ? row[index].slice(0) : {};
            if (kind == 5) (row[index] as [any, any][]).forEach(v => value[v[0]] = v[1]);
            data[name] = value;
        });
        return data;
    }

__SHEET_DATA_GET__
    /** @implements */
    export function get<T extends keyof Sheets>(sheet: T, ...args: any[]): Sheets[T] | null {
        let hdr = hdrs[sheet], indexs = hdr && hdr.map<[boolean, number]>((v, i) => [v[2], i]).filter(v => v[0]);
        for (const row of rowss[sheet] || []) {
            if (indexs.every((v, i) => row[v[1]] == args[i])) {
                return make(sheet, row);
            }
        }
        return null;
    }

    export function forEach<S extends keyof Sheets>(sheet: S, iterator: (data: Sheets[S]) => void): void {
        (rowss[sheet] || [] as any[][]).forEach(row => iterator(make(sheet, row)));
    }
    export function filter<S extends keyof Sheets>(sheet: S, iterator: (data: Sheets[S]) => boolean): Sheets[S][] {
        return (rowss[sheet] || [] as any[][]).map(row => make(sheet, row)).filter(data => iterator(data));
    }
    export function map<T, S extends keyof Sheets>(sheet: S, iterator: (data: Sheets[S]) => T): T[] {
        return (rowss[sheet] || [] as any[][]).map(row => iterator(make(sheet, row)));
    }
    export function some<S extends keyof Sheets>(sheet: S, iterator: (data: Sheets[S]) => boolean): boolean {
        return (rowss[sheet] || [] as any[][]).some(row => iterator(make(sheet, row)));
    }
    export function every<S extends keyof Sheets>(sheet: S, iterator: (data: Sheets[S]) => boolean): boolean {
        return (rowss[sheet] || [] as any[][]).every(row => iterator(make(sheet, row)));
    }
}`;

/** @param {string} value */
const lCap = (value) => value[0].toLowerCase() + value.substr(1);
/** @param {string} value */
const uCap = value => value[0].toUpperCase() + value.substr(1);

/**
 * translate
 * @param {Object<string, Array<Array<any>>} sheets XlsxJson 导出的json对象
 * @returns {string}
 */
module.exports.translate = sheets => {
    /**
     * @param {string} types
     * @return {string}
     */
    const toType = (types) => {
        let type = types;
        if (types[0] == '[') {
            type = '[' + types.substring(1, types.length - 1).split(',').join(', ') + ']';
        } else if (types[types.length - 1] == '[') {
            type = types.substring(0, types.indexOf('[')) + '[]';
        } else if (types.indexOf(':') >= 0) {
            type = '{ ' + types.split(':').map((type, index) => {
                if (index == 0) {
                    switch (type) {
                        case 'str': type = '[_: string/*' + type + '*/]'; break;
                        case 'bool': type = '[_: number/*0 | 1*/]'; break;
                        default: type = '[_: number/*' + type + '*/]'; break;
                    }
                }
                return type;
            }).join(': ') + ' }';
        }
        return type;
    }
    let __SHEETS_DATA__ = '', __SHEETS_DATA_MAP__ = '', __SHEET_DATA_GET__ = '';
    for (const sheet in sheets) {
        /** @type [string, string, number, string][] */
        let header = sheets[sheet][0];
        /** @type [string, string, string][] */
        let primarys = [];

        __SHEETS_DATA__ += __SHEETS_DATA__.length > 0 ? '\n' : '';
        __SHEETS_DATA__ += `    export interface ${uCap(sheet)} {\n`;
        header.forEach(([name, types, primary, comment]) => {
            let type = toType(types);
            if (comment) __SHEETS_DATA__ += `        /** ${comment} */\n`;
            __SHEETS_DATA__ += `        ${name}: ${type};\n`;
            if (primary) primarys.push([name, type, comment]);
        });
        __SHEETS_DATA__ += `    }`;

        __SHEETS_DATA_MAP__ += __SHEETS_DATA_MAP__.length > 0 ? '\n' : '';
        __SHEETS_DATA_MAP__ += `        ${sheet}: ${uCap(sheet)};`;

        __SHEET_DATA_GET__ += __SHEET_DATA_GET__.length > 0 ? '\n' : '';
        if (primarys.length) {
            __SHEET_DATA_GET__ += `    /**\n${primarys.map(([name, _, comment]) => `     * @param _${name} ${comment}`).join('\n')}\n     */\n`;
            __SHEET_DATA_GET__ += `    export function get(sheet: '${sheet}', ${primarys.map(([name, type]) => `_${name}: ${type}`).join(', ')}): ${uCap(sheet)} | null;`;
        }
    }

    return tsTemplate
        .replace('__SHEETS_DATA__', __SHEETS_DATA__)
        .replace('__SHEETS_DATA_MAP__', __SHEETS_DATA_MAP__)
        .replace('__SHEET_DATA_GET__', __SHEET_DATA_GET__);
}
