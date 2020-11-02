const fs = require('fs');
const exporter = require('./XlsxJson.Exporter')

/**
 * 解析参数
 * @param {string[]} argv 参数列表
 * @return {{file:string,outType:string,outFile:string}}
 */
function parseCommand(argv) {
    let file = '', outType = '', outFile = '.';

    for (let i = 0; i < argv.length; i++) {
        if (argv[i] == '-t') {
            if (!argv[i + 1] || argv[i + 1][0] == '-') {
                file = '';
                break;
            }
            outType = argv[++i].toLowerCase();
        } else if (argv[i] == '-o') {
            if (!argv[i + 1] || argv[i + 1][0] == '-') {
                file = '';
                break;
            }
            outFile = argv[++i].toLowerCase();
        } else if (argv[i][0] != '-' && !file) {
            file = argv[i];
        } else {
            file = '';
            break;
        }
    }

    if (!file || !outType || !outFile)
        console.error('usage: node exporter JSON文件 -t 导出语言 [-o 输出路径]');

    return { file, outType, outFile };
}

/**
 * 检查
 * @param {Object<string, any[][]>} sheets JSON数据
 * @returns {Boolean}
 */
function check(sheets) {
    let types = ['u8', 'u16', 'u32', 'u64', 'i8', 'i16', 'i32', 'i64', 'f32', 'f64', 'str', 'bool'];
    return typeof sheets == 'object' && Object.entries(sheets).every(([name, rows]) =>
        /^[a-zA-Z][a-zA-Z\d_]*$/.test(name) &&
        Array.isArray(rows) &&
        rows.length > 1 &&
        rows.every((values, row) =>
            Array.isArray(values) &&
            (row > 0 ? values.length == rows[0].length : true) &&
            values.every((value, col) => {
                if (row == 0) return Array.isArray(value) && value.length == 4 && value.every((v, i) => {
                    if (i == 0) return typeof v == 'string' && /^[a-zA-Z][a-zA-Z\d_]*$/.test(v);
                    if (i == 1) {
                        if (typeof v != 'string') return false;
                        let leftSplit = v.indexOf('[');
                        if (leftSplit == 0)
                            return v[v.length - 1] == ']' && v.substring(1, v.length - 1).split(',').every(sv => types.includes(sv));
                        if (leftSplit > 0) {
                            let cnt = v.substring(leftSplit + 1, v.length - 1);
                            return v[v.length - 1] == ']' && types.includes(v.substring(0, leftSplit)) && /^\d*$/.test(cnt) && +cnt > 0;
                        }
                        leftSplit = v.indexOf(':');
                        if (leftSplit > 0)
                            return v.split(':').every(sv => types.includes(sv))
                        return types.includes(v);
                    }
                    if (i == 2) return v === 0 || v === 1;
                    if (i == 3) return typeof v == 'string';
                });
                /** @type [string,string,number,string] */
                let header = rows[0];
                /**
                 * 
                 * 
                 * 
                 * 
                 */
                return true;
            })
        )
    );
}

function main() {
    const { file, outType, outFile } = parseCommand(process.argv.slice(2));
    if (!file) return -1;
    if (!fs.existsSync(file)) {
        console.error(file + ' 文件不存在.');
        return -1;
    }
    let sheets = JSON.parse(fs.readFileSync(file, 'utf8'));
    let exptr = exporter[outType];
    if (!exptr) {
        console.error('不支持' + outType + '输出');
        return -1;
    }

    try {
        fs.writeFileSync(outFile, exptr(sheets), 'utf8');
    } catch (err) {
        console.error(err);
        return -1;
    }
    return 0;
}
process.exitCode = main();