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
                    index[1].indexOf(':') >= 0 ? 4 : index[1][0] == '[' ? 3 : index[1].endsWith(']') ? 2 : 1,
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
            let value = kind == 1 ? row[index] : kind == 2 || kind == 3 ? row[index].slice(0) : {};
            if (kind == 4) (row[index] as [any, any][]).forEach(v => value[v[0]] = v[1]);
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

func parse(sheets map[string][][]interface{}) {
__SHEETS_PARSE_0__
	for sheet, rows := range sheets {
__SHEETS_PARSE__
	}
__SHEETS_PARSE_1__
}

func LoadByFile(path string) (err interface{}) {
	file, err := os.Open(path)
	if err != nil {
		return err
	}
	bytes, err := ioutil.ReadAll(file)
	if err != nil {
		return err
	}

	return LoadByData(bytes)
}
func LoadByData(bytes []byte) (err interface{}) {
	var data map[string][][]any
	err = json.Unmarshal(bytes, &data)
	if err != nil {
		return err
	}

	defer func() { err = recover() }()
	parse(data)

	return err
}
`;

/** @param {string} value */
const lCap = (value) => value[0].toLowerCase() + value.substr(1);
/** @param {string} value */
const uCap = value => value[0].toUpperCase() + value.substr(1);

module.exports.ts = sheets => {
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

module.exports.go = (sheets) => {
    const _lCap = str => '_' + lCap(str);
    let __SHEETS_DATA__ = '', __SHEETS_FUN__ = '', __SHEETS_PARSE_0__ = '', __SHEETS_PARSE__ = '', __SHEETS_PARSE_1__ = '';
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
        __SHEETS_DATA__ += `var ${_lCap(sheet)}s []${_lCap(sheet)}`;

        __SHEETS_FUN__ += __SHEETS_FUN__.length > 0 ? '\n\n' : '';
        __SHEETS_FUN__ += `type ${_lCap(sheet)} struct {\n`
        /** @type [string,string] */
        let members = [], maxDefineLength = 0;
        nameTypes.forEach(([name, type]) => {
            if (typeof type == 'string') {
                let start = `\t${_lCap(name)}`;
                if (start.length > maxDefineLength) maxDefineLength = start.length;
                members.push([start, type]);
            } else type.forEach((type, index) => {
                let start = `\t${_lCap(name)}${index}`;
                if (start.length > maxDefineLength) maxDefineLength = start.length;
                members.push([start, type]);
            });
        });
        members.forEach(v => {
            while (v[0].length < maxDefineLength) v[0] += ' ';
            __SHEETS_FUN__ += v[0] + ' ' + v[1] + '\n';
        });
        __SHEETS_FUN__ += '}\n';
        __SHEETS_FUN__ += `type ${uCap(sheet)} interface {\n`;
        members = [], maxDefineLength = 0;
        nameTypes.forEach(([name, type], index) => {
            let start = `\t${uCap(name)}() ${typeof type == 'string' ? type : '(' + type.join(', ') + ')'}`, end = `// ${header[index][3]}`;
            if (start.length > maxDefineLength) maxDefineLength = start.length;
            members.push([start, end]);
        });
        members.forEach(v => {
            while (v[0].length < maxDefineLength) v[0] += ' ';
            __SHEETS_FUN__ += v[0] + ' ' + v[1] + '\n';
        });
        __SHEETS_FUN__ += '}\n\n';

        members = [], maxDefineLength = 0;
        nameTypes.forEach(([name, type]) => {
            let isBreak = false;
            let start = `func (d *${_lCap(sheet)}) ${uCap(name)}() ${typeof type == 'string' ? type : '(' + type.join(', ') + ')'}`, end = '';
            if (typeof type == 'string') {
                if (!type.startsWith('map'))
                    end = `{ return d.${_lCap(name)} }`;
                else
                    isBreak = true, end = `{\n\tdata := ${type}{}\n\tfor k, v := range d.${_lCap(name)} {\n\t\tdata[k] = v\n\t}\n\treturn data\n}`;
            } else
                end = `{ return ${type.map((_, i) => `d.${_lCap(name)}${i}`).join(', ')} }`;
            if (isBreak) {
                members.forEach(v => {
                    while (v[0].length < maxDefineLength) v[0] += ' ';
                    __SHEETS_FUN__ += v[0] + ' ' + v[1] + '\n';
                });
                __SHEETS_FUN__ += start + ' ' + end + '\n';
                members = [], maxDefineLength = 0;
            } else {
                if (start.length > maxDefineLength) maxDefineLength = start.length;
                members.push([start, end]);
            }
        });
        members.forEach(v => {
            while (v[0].length < maxDefineLength) v[0] += ' ';
            __SHEETS_FUN__ += v[0] + ' ' + v[1] + '\n';
        });

        __SHEETS_FUN__ += '\n';
        /** @type [string, string | string[]][] */
        let indexNameTypes = []; nameTypes.forEach((v, i) => header[i][2] == 1 && indexNameTypes.push([...v, i]));
        if (indexNameTypes.length) {
            __SHEETS_FUN__ += `${indexNameTypes.map(([n, t, i]) => `// @param ${_lCap(n)} ${header[i][3]}`).join('\n')}\n`;
            __SHEETS_FUN__ += `func Get${uCap(sheet)}(${indexNameTypes.map(([n, t]) => `${_lCap(n)} ${t}`).join(', ')}) ${uCap(sheet)} {\n\t`;
            __SHEETS_FUN__ += `for _, v := range ${_lCap(sheet)}s {\n\t\tif ${indexNameTypes.map(([n, _]) => `v.${_lCap(n)} == ${_lCap(n)}`).join(' && ')} {\n\t\t\treturn &v\n\t\t}\n\t}\n\treturn nil\n}\n`
        }
        __SHEETS_FUN__ += `func Find${uCap(sheet)}(iterator func(data ${uCap(sheet)}) bool) ${uCap(sheet)} {\n\t`;
        __SHEETS_FUN__ += `for _, v := range ${_lCap(sheet)}s {\n\t\tif iterator(&v) {\n\t\t\treturn &v\n\t\t}\n\t}\n\treturn nil\n}\n`;
        __SHEETS_FUN__ += `func Each${uCap(sheet)}(iterator func(data ${uCap(sheet)})) {\n\t`;
        __SHEETS_FUN__ += `for _, v := range ${_lCap(sheet)}s {\n\t\titerator(&v)\n\t}\n}\n`
        __SHEETS_FUN__ += `func Filter${uCap(sheet)}(iterator func(data ${uCap(sheet)}) bool) []${uCap(sheet)} {\n\t`;
        __SHEETS_FUN__ += `var datas []${uCap(sheet)}\n\tfor _, v := range ${_lCap(sheet)}s {\n\t\tif iterator(&v) {\n\t\t\tdatas = append(datas, &v)\n\t\t}\n\t}\n\treturn datas\n}\n`;
        __SHEETS_FUN__ += `func Some${uCap(sheet)}(iterator func(data ${uCap(sheet)}) bool) bool {\n\t`;
        __SHEETS_FUN__ += `for _, v := range ${_lCap(sheet)}s {\n\t\tif iterator(&v) {\n\t\t\treturn true\n\t\t}\n\t}\n\treturn false\n}\n`;
        __SHEETS_FUN__ += `func Every${uCap(sheet)}(iterator func(data ${uCap(sheet)}) bool) bool {\n\t`;
        __SHEETS_FUN__ += `for _, v := range ${_lCap(sheet)}s {\n\t\tif !iterator(&v) {\n\t\t\treturn false\n\t\t}\n\t}\n\treturn true\n}`;

        __SHEETS_PARSE_0__ += __SHEETS_PARSE_0__.length ? '\n' : '';
        __SHEETS_PARSE_0__ += `\tvar _${_lCap(sheet)} []${_lCap(sheet)}`;

        __SHEETS_PARSE__ += __SHEETS_PARSE__.length ? '\n' : '';
        __SHEETS_PARSE__ += `\t\tif sheet == "${sheet}" {\n\t\t\tfor i := 1; i < len(rows); i++ {\n\t\t\t\tdata := ${_lCap(sheet)}{}\n\t\t\t\t`;
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
                if (!type.startsWith('map')) __SHEETS_PARSE__ += `data.${_lCap(name)} = ${toValue(type, `rows[i][${index}]`)}\n\t\t\t\t`;
                else {
                    __SHEETS_PARSE__ += `data.${_lCap(name)} = ${type}{}\n\t\t\t\t`;
                    __SHEETS_PARSE__ += `for _, v := range rows[i][${index}].([]any) {\n\t\t\t\t\tif v, ok := v.([]any); ok {\n\t\t\t\t\t\t`;
                    let kt = type.substring(type.indexOf('[') + 1, type.indexOf(']')), vt = type.substring(type.indexOf(']') + 1);
                    __SHEETS_PARSE__ += `data.${_lCap(name)}[${toValue(kt, 'v[0]')}] = ${toValue(vt, 'v[1]')}\n\t\t\t\t\t}\n\t\t\t\t}\n\t\t\t\t`;
                }
            }
            else type.forEach((type, i) => __SHEETS_PARSE__ += `data.${_lCap(name)}${i} = ${toValue(type, `rows[i][${index}].([]any)[${i}]`)}\n\t\t\t\t`);
        });
        __SHEETS_PARSE__ += `_${_lCap(sheet)} = append(_${_lCap(sheet)}, data)\n\t\t\t}\n\t\t}`;

        __SHEETS_PARSE_1__ += __SHEETS_PARSE_1__.length ? '\n' : '';
        __SHEETS_PARSE_1__ += `\t${_lCap(sheet)}s = _${_lCap(sheet)}`;
    }

    return goTemplate
        .replace('__SHEETS_DATA__', __SHEETS_DATA__)
        .replace('__SHEETS_FUN__', __SHEETS_FUN__)
        .replace('__SHEETS_PARSE_0__', __SHEETS_PARSE_0__)
        .replace('__SHEETS_PARSE__', __SHEETS_PARSE__)
        .replace('__SHEETS_PARSE_1__', __SHEETS_PARSE_1__);
}
