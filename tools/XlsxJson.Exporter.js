let tsTemplate = `export namespace xlsx {
    export type u8 = number;
    export type u16 = number;
    export type u32 = number;
    export type u64 = number;
    export type i8 = number;
    export type i16 = number;
    export type i32 = number;
    export type i64 = number;
    export type bool = 0 | 1;
    export type str = string;

__SHEETS_DATA__

    type Sheets = {
__SHEETS_DATA_MAP__
    };

    let hdrs: { [sheet: string]: [string, number, boolean][] } = {};
    let rowss: { [sheet: string]: any[][] } = {};

    /**
     * 加载json
     * @param path json路径
     */ 
    export function load(path: string): Promise<boolean> {
        return new Promise<boolean>(resolve => {
            const init = (json: any) => {
                hdrs = {}, rowss = json;
                for (const sheet in rowss) {
                    let hdr = rowss[sheet].shift()!;
                    hdrs[sheet] = hdr.map(index => [
                        index[0],
                        index[1].indexOf(':') >= 0 ? 4 : index[1][0] == '[' ? 3 : index[1].endsWith(']') ? 2 : 1,
                        index[2] == 1
                    ]);
                }
            };
__TS_TEMPLATE__
        });
    }

    function make(sheet: keyof Sheets, row: any[]): any {
        let hdr = hdrs[sheet], data: any = {};
        hdr.forEach(([name, kind], index) => {
            let value = kind == 1 ? row[index] : kind == 2 || kind == 3 ? row[index].slice(0) : {};
            if (kind == 4) (row[index] as [any, any][]).forEach(v => value[v[0]] = v[1]);
            data[name] = value;
        });
        return data;
    }

__SHEET_DATA_GET__
    /** @implements */
    export function get(sheet: keyof Sheets, ...args: any[]): any {
        let hdr = hdrs[sheet], indexs = hdr && hdr.map<[boolean, number]>((v, i) => [v[2], i]).filter(v => v[0]);
        for (const row of rowss[sheet] || []) {
            if (indexs.every((v, i) => row[v[1]] == args[i])) {
                return make(sheet, row);
            }
        }
        return null;
    }

    export function forEach<S extends keyof Sheets, D = Sheets[S]>(sheet: S, iterator: (data: D) => void): void {
        (rowss[sheet] || []).forEach(row => iterator(make(sheet, row)));
    }
    export function filter<S extends keyof Sheets, D = Sheets[S]>(sheet: S, iterator: (data: D) => boolean): D[] {
        let datas: D[] = [];
        (rowss[sheet] || []).forEach(row => iterator(make(sheet, row)) && datas.push(make(sheet, row)));
        return datas;
    }
    export function map<T, S extends keyof Sheets, D = Sheets[S]>(sheet: S, iterator: (data: D) => T): T[] {
        return (rowss[sheet] || []).map(row => iterator(make(sheet, row)));
    }
    export function some<S extends keyof Sheets, D = Sheets[S]>(sheet: S, iterator: (data: D) => boolean): boolean {
        return (rowss[sheet] || []).some(row => iterator(make(sheet, row)));
    }
    export function every<S extends keyof Sheets, D = Sheets[S]>(sheet: S, iterator: (data: D) => boolean): boolean {
        return (rowss[sheet] || []).every(row => iterator(make(sheet, row)));
    }
}`;
function ts(template, sheets) {
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
        let SHEET = 'XD' + sheet[0].toUpperCase() + sheet.substr(1);
        /** @type [string, string, number, string][] */
        let header = sheets[sheet][0];
        /** @type [string, string, string][] */
        let primarys = [];

        __SHEETS_DATA__ += __SHEETS_DATA__.length > 0 ? '\n' : '';
        __SHEETS_DATA__ += `    export interface ${SHEET} {\n`;
        header.forEach(([name, types, primary, comment]) => {
            let type = toType(types);
            if (comment) __SHEETS_DATA__ += `        /** ${comment} */\n`;
            __SHEETS_DATA__ += `        ${name}: ${type};\n`;
            if (primary) primarys.push([name, type, comment]);
        });
        __SHEETS_DATA__ += `    }`;

        __SHEETS_DATA_MAP__ += __SHEETS_DATA_MAP__.length > 0 ? '\n' : '';
        __SHEETS_DATA_MAP__ += `        ${sheet}: ${SHEET};`;

        __SHEET_DATA_GET__ += __SHEET_DATA_GET__.length > 0 ? '\n' : '';
        if (primarys.length) {
            __SHEET_DATA_GET__ += `    /**\n${primarys.map(([name, _, comment]) => '     * @param ' + name + ' ' + comment).join('\n')}\n     */\n`;
            __SHEET_DATA_GET__ += `    export function get(sheet: '${sheet}', ${primarys.map(([name, type]) => name + ': ' + type).join(', ')}): ${SHEET};`;
        }
    }

    return template
        .replace('__SHEETS_DATA__', __SHEETS_DATA__)
        .replace('__SHEETS_DATA_MAP__', __SHEETS_DATA_MAP__)
        .replace('__SHEET_DATA_GET__', __SHEET_DATA_GET__);
}

module.exports.cc = function (sheets) {
    let template = tsTemplate.replace('__TS_TEMPLATE__',
        `            cc.loader.loadRes(path, cc.JsonAsset, (err: Error | null, data: cc.JsonAsset) => {
                if (err) { 
                    cc.error(err);
                    resolve(false);
                } else {
                    init(data.json);
                    resolve(true);
                }
            });`);
    return ts(template, sheets);
}

module.exports.node = function (sheets) {
    let template = tsTemplate.replace('__TS_TEMPLATE__',
        `            (require('fs') as typeof import('fs')).readFile.readFile(path, 'utf8', (err, data) => {
                if (err) { 
                    console.error(err);
                    resolve(false);
                } else {
                    init(JSON.parse(data));
                    resolve(true);
                }
            });`);
    return ts(template, sheets);
}