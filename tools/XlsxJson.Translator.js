const fs = require('fs');

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

function main() {
    const { file, outType, outFile } = parseCommand(process.argv.slice(2));
    if (!file) return -1;
    if (!fs.existsSync(file)) {
        console.error(file + ' 文件不存在.');
        return -1;
    }
    let sheets = JSON.parse(fs.readFileSync(file, 'utf8'));
    let exptr = require(`./XlsxJson.Translator.${outType}`);
    if (!exptr || !exptr.translate) {
        console.error('不支持' + outType + '输出');
        return -1;
    }

    try {
        fs.writeFileSync(outFile, exptr.translate(sheets), 'utf8');
        console.log('输出:', outFile);
    } catch (err) {
        console.error(err);
        return -1;
    }
    return 0;
}
process.exitCode = main();