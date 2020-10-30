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
let goTemplate = `package xlsx

import (
	"encoding/json"
	"io/ioutil"
	"os"
)

type u8 = uint8
type u16 = uint16
type u32 = uint32
type u64 = uint64
type i8 = int8
type i16 = int16
type i32 = int32
type i64 = int64
type f32 = float32
type f64 = float64
type str = string
type any = interface{}

__SHEETS_DATA__

__SHEETS_FUN__

// 加载并解析JSON
// 若失败则返回错误，成功则覆盖数据
func Load(path string) (err interface{}) {
	_0, err := os.Open(path)
	if err != nil {
		return err
	}
	_1, err := ioutil.ReadAll(_0)
	if err != nil {
		return err
	}
	var _2 map[string][][]any
	err = json.Unmarshal(_1, &_2)
	if err != nil {
		return err
	}

	defer func() { err = recover() }()

__SHEETS_PARSE_0__
	for _3, _4 := range _2 {
__SHEETS_PARSE__
	}
__SHEETS_PARSE_1__

	return err
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
        /** @type [string, string, number, string][] */
        let header = sheets[sheet][0];
        /** @type [string, string, string][] */
        let primarys = [];

        __SHEETS_DATA__ += __SHEETS_DATA__.length > 0 ? '\n' : '';
        __SHEETS_DATA__ += `    export interface ${upperCaption(sheet)} {\n`;
        header.forEach(([name, types, primary, comment]) => {
            let type = toType(types);
            if (comment) __SHEETS_DATA__ += `        /** ${comment} */\n`;
            __SHEETS_DATA__ += `        ${name}: ${type};\n`;
            if (primary) primarys.push([name, type, comment]);
        });
        __SHEETS_DATA__ += `    }`;

        __SHEETS_DATA_MAP__ += __SHEETS_DATA_MAP__.length > 0 ? '\n' : '';
        __SHEETS_DATA_MAP__ += `        ${sheet}: ${upperCaption(sheet)};`;

        __SHEET_DATA_GET__ += __SHEET_DATA_GET__.length > 0 ? '\n' : '';
        if (primarys.length) {
            __SHEET_DATA_GET__ += `    /**\n${primarys.map(([name, _, comment]) => `     * @param _${name} ${comment}`).join('\n')}\n     */\n`;
            __SHEET_DATA_GET__ += `    export function get(sheet: '${sheet}', ${primarys.map(([name, type]) => `_${name}: ${type}`).join(', ')}): ${upperCaption(sheet)};`;
        }
    }

    return template
        .replace('__SHEETS_DATA__', __SHEETS_DATA__)
        .replace('__SHEETS_DATA_MAP__', __SHEETS_DATA_MAP__)
        .replace('__SHEET_DATA_GET__', __SHEET_DATA_GET__);
}

function go(template, sheets) {
    /** @param {string} value */
    const lowerCaption = (value) => '_' + value[0].toLowerCase() + value.substr(1);
    /** @param {string} value */
    const upperCaption = value => value[0].toUpperCase() + value.substr(1);

    let __SHEETS_DATA__ = '', __SHEETS_FUN__ = '', __SHEETS_PARSE_0__ = '', __SHEETS_PARSE__ = '', __SHEETS_PARSE_1__ = '';
    let count = 0;
    for (const sheet in sheets) {
        /** @type [string, string, number, string][] */
        let header = sheets[sheet][0];
        /** @type [string, string | string[]][] */
        let nameTypes = []
        header.forEach(v => {
            let name = v[0], type = v[1];
            if (type[0] == '[') {
                type = type.substr(1, type.length - 2).split(',');
            } else if (type[type.length - 1] == ']') {
                let count = type.substring(type.indexOf('[') + 1, type.length - 1);
                type = `[${count}]${type.substring(0, type.indexOf('['))}`;
            } else if (type.includes(':')) {
                type = `map[${type.split(':').join(']')}`;
            }
            nameTypes.push([name, type]);
        });

        __SHEETS_DATA__ += __SHEETS_DATA__.length > 0 ? '\n' : '';
        __SHEETS_DATA__ += `var ${lowerCaption(sheet)}s []${lowerCaption(sheet)}`;

        __SHEETS_FUN__ += __SHEETS_FUN__.length > 0 ? '\n\n' : '';
        __SHEETS_FUN__ += `type ${lowerCaption(sheet)} struct {\n`
        nameTypes.forEach(([name, type]) => {
            if (typeof type == 'string') __SHEETS_FUN__ += `    ${lowerCaption(name)} ${type}\n`;
            else type.forEach((type, index) => __SHEETS_FUN__ += `    ${lowerCaption(name)}${index} ${type}\n`)
        });
        __SHEETS_FUN__ += '}\n';
        __SHEETS_FUN__ += `type ${upperCaption(sheet)} interface {\n`;
        nameTypes.forEach(([name, type], index) => __SHEETS_FUN__ += `    ${upperCaption(name)}() ${typeof type == 'string' ? type : '(' + type.join(', ') + ')'} // ${header[index][3]}\n`);
        __SHEETS_FUN__ += '}\n\n';

        nameTypes.forEach(([name, type]) => {
            __SHEETS_FUN__ += `func (d *${lowerCaption(sheet)}) ${upperCaption(name)}() ${typeof type == 'string' ? type : '(' + type.join(', ') + ')'} {`;
            if (typeof type == 'string') {
                if (!type.startsWith('map')) __SHEETS_FUN__ += ` return d.${lowerCaption(name)} }\n`;
                else __SHEETS_FUN__ += `\n\t_0 := ${type}{}\n\tfor k, v := range d.${lowerCaption(name)} {\n\t\t_0[k] = v\n\t}\n\treturn _0\n}\n`;
            } else __SHEETS_FUN__ += ` return ${type.map((_, i) => `d.${lowerCaption(name)}${i}`).join(', ')} }\n`;
        });

        /** @type [string, string | string[]][] */
        let indexNameTypes = []; nameTypes.forEach((v, i) => header[i][2] == 1 && indexNameTypes.push([...v, i]));
        if (indexNameTypes.length) {
            __SHEETS_FUN__ += `// \n${indexNameTypes.map(([n, t, i]) => `// @param ${lowerCaption(n)} ${header[i][3]}`).join('\n')} \n`;
            __SHEETS_FUN__ += `func Get${upperCaption(sheet)}(${indexNameTypes.map(([n, t]) => `${lowerCaption(n)} ${t}`).join(', ')}) ${upperCaption(sheet)} {\n\t`;
            __SHEETS_FUN__ += `for _, _0 := range ${lowerCaption(sheet)}s {\n\t\tif ${indexNameTypes.map(([n, _]) => `_0.${lowerCaption(n)} == ${lowerCaption(n)}`).join(' && ')} {\n\t\t\treturn &_0\n\t\t}\n\t}\n\treturn nil\n}\n`
        }
        __SHEETS_FUN__ += `func Find${upperCaption(sheet)}(iterator func(data ${upperCaption(sheet)}) bool) ${upperCaption(sheet)} {\n\t`;
        __SHEETS_FUN__ += `for _, _0 := range ${lowerCaption(sheet)}s {\n\t\tif iterator(&_0) {\n\t\t\treturn &_0\n\t\t}\n\t}\n\treturn nil\n}\n`;
        __SHEETS_FUN__ += `func Each${upperCaption(sheet)}(iterator func(data ${upperCaption(sheet)})) {\n\t`;
        __SHEETS_FUN__ += `for _, _0 := range ${lowerCaption(sheet)}s {\n\t\titerator(&_0)\n\t}\n}\n`
        __SHEETS_FUN__ += `func Filter${upperCaption(sheet)}(iterator func(data ${upperCaption(sheet)}) bool) []${upperCaption(sheet)} {\n\t`;
        __SHEETS_FUN__ += `var _0 []${upperCaption(sheet)}\n\tfor _, _1 := range ${lowerCaption(sheet)}s {\n\t\tif iterator(&_1) {\n\t\t\t_0 = append(_0, &_1)\n\t\t}\n\t}\n\treturn _0\n}`;

        __SHEETS_PARSE_0__ += __SHEETS_PARSE_0__.length ? '\n' : '';
        __SHEETS_PARSE_0__ += `\tvar __${count} []${lowerCaption(sheet)}`;

        __SHEETS_PARSE__ += __SHEETS_PARSE__.length ? '\n' : '';
        __SHEETS_PARSE__ += `\t\tif _3 == "${sheet}" {\n\t\t\tfor i := 1; i < len(_4); i++ {\n\t\t\t\t_5 := ${lowerCaption(sheet)}{}\n\t\t\t\t`;
        const toValue = (type, value) => {
            if (type == 'str')
                return `${value}.(str)`;
            else if (type == 'bool')
                return `${value}.(f64) == 1`;
            else
                return `${type}(${value}.(f64))`;
        }
        nameTypes.forEach(([name, type], index) => {
            if (typeof type == 'string') {
                if (!type.startsWith('map')) __SHEETS_PARSE__ += `_5.${lowerCaption(name)} = ${toValue(type, `_4[i][${index}]`)}\n\t\t\t\t`;
                else {
                    __SHEETS_PARSE__ += `_5.${lowerCaption(name)} = ${type}{}\n\t\t\t\t`;
                    __SHEETS_PARSE__ += `for _, v := range _4[i][${index}].([]any) {\n\t\t\t\t\tif v, ok := v.([]any); ok {\n\t\t\t\t\t\t`;
                    let kt = type.substring(type.indexOf('[') + 1, type.indexOf(']')), vt = type.substring(type.indexOf(']') + 1);
                    __SHEETS_PARSE__ += `_5.${lowerCaption(name)}[${toValue(kt, 'v[0]')}] = ${toValue(vt, 'v[1]')}\n\t\t\t\t\t}\n\t\t\t\t}\n\t\t\t\t`;
                }
            }
            else type.forEach((type, i) => __SHEETS_PARSE__ += `_5.${lowerCaption(name)}${i} = ${toValue(type, `_4[i][${index}].([]any)[${i}]`)}\n\t\t\t\t`);
        });
        __SHEETS_PARSE__ += `__${count} = append(__${count}, _5)\n\t\t\t}\n\t\t}`;

        __SHEETS_PARSE_1__ += __SHEETS_PARSE_1__.length ? '\n' : '';
        __SHEETS_PARSE_1__ += `\t${lowerCaption(sheet)}s = __${count}`;

        count++;
    }

    return template
        .replace('__SHEETS_DATA__', __SHEETS_DATA__)
        .replace('__SHEETS_FUN__', __SHEETS_FUN__)
        .replace('__SHEETS_PARSE_0__', __SHEETS_PARSE_0__)
        .replace('__SHEETS_PARSE__', __SHEETS_PARSE__)
        .replace('__SHEETS_PARSE_1__', __SHEETS_PARSE_1__);
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
        `            (require('fs') as typeof import('fs')).readFile(path, 'utf8', (err, data) => {
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

module.exports.go = function (sheets) { return go(goTemplate, sheets); }