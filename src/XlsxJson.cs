using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsxJson
{
    /// <summary>
    /// 值
    /// </summary>
    class Value
    {
        public string Text { get; protected set; }
        public short Error { get; protected set; }
        public string ErrorInfo => Error switch
        {
            1001 => "索引不能为空",
            1002 => "索引只能包含下划线字母和数字且必须以下划线或字母开头",
            1003 => "索引重复",
            1004 => "索引行不允许出现合并单元格",

            1101 => "类型不能为空",
            1102 => "类型定义错误",
            1103 => "复合元素类型定义错误",
            1104 => "数组元素类型定义错误",
            1105 => "数组长度必须大于0",
            1106 => "数组长度超出范围",
            1107 => "数组长度定义错误",
            1108 => "元组元素类型定义错误",
            1109 => "元组元素个数必须大于1",
            1110 => "字典元素类型定义错误",
            1111 => "主索引的类型必须时基础类型",

            1201 => "导出定义不能为合并单元格",
            1202 => "导出定义重复",
            1203 => "此处只能填写 \"-\" 或 空",
            1204 => "此处只能填写 \"*\"、\"-\" 或 空",

            1301 => "不是整数",
            1302 => "不是小数",
            1303 => "超出精度范围",
            1304 => "不是整数或超出精度范围",
            1305 => "不是小数或超出精度范围",
            1306 => @"复合类型中的字符串需用""包括，且字符串中包含""时需替换成\""，且字符串中包含\""时需替换成\\\""",
            1307 => "bool类型只能填0和1",
            1308 => "值数量溢出",
            1309 => "键值对的键值需用:分隔",
            1310 => "键重复",
            1350 => "该行主索引组合重复",
            1399 => "无法解析",

            2001 => "表名只能包含下划线字母和数字且必须以下划线或字母开头",

            _ => ""
        };
    }

    /// <summary>
    /// 索引
    /// </summary>
    class Index : Value
    {
        /// <summary>
        /// 主键
        /// </summary>
        public bool IsPrimary { get; }

        public Index(XlsxTextReader.Cell cell)
        {
            Text = cell.Value.Trim();
            IsPrimary = false;

            if (Text.Length == 0) Error = 1001;
            else if (Text[0] == '*')
            {
                IsPrimary = true;
                Text = Text.Substring(1);
            }
            if (!Regex.IsMatch(Text, @"^[a-zA-Z_][a-zA-Z\d_]*$"))
                Error = 1002;
        }
    }

    /// <summary>
    /// 类型<br/>
    /// 基本类型: uint8, uint16, uint32, uint64, int8, int16, int32, int64, float32, float64, string <br/>
    /// 数组: int8[], int16[5] ..., 数组元素必须是基本类型, 指定长度则为定长数组 <br/>
    /// 元组: [int8, int16], [float32, int16, string] ..., 元组元素必须是基本类型 <br/>
    /// 字典: int8:string, string:int32 ..., 字典元素必须是基本类型
    /// </summary>
    class Type : Value
    {
        public const string __u8 = "u8";
        public const string __u16 = "u16";
        public const string __u32 = "u32";
        public const string __u64 = "u64";
        public const string __i8 = "i8";
        public const string __i16 = "i16";
        public const string __i32 = "i32";
        public const string __i64 = "i64";
        public const string __f32 = "f32";
        public const string __f64 = "f64";
        public const string __bool = "bool";
        public const string __str = "str";
        private static bool Has(string type) => type == __u8 || type == __u16 || type == __u32 || type == __u64
                                            || type == __i8 || type == __i16 || type == __i32 || type == __i64
                                            || type == __f32 || type == __f64
                                            || type == __bool || type == __str;
        public enum TypeKind { Simple, Array, Tuple, Dictionary }

        public TypeKind Kind { get; }
        public ReadOnlyCollection<string> Items { get; }
        public int Count { get; }

        public Type(XlsxTextReader.Cell cell)
        {
            var kind = TypeKind.Simple;
            var items = new List<string>();
            var count = -1;

            var text = cell.Value.Trim();
            if (text.Length == 0) Error = 1101;
            else if (Has(text)) items.Add(text);
            else
            {
                var values = new string[0];
                if (text[text.Length - 1] == ']')    // 数组、元组
                {
                    var leftIndex = text.IndexOf('[');
                    if (leftIndex > 0) // 数组
                    {
                        kind = TypeKind.Array;
                        values = new string[1] { text.Substring(0, leftIndex) };
                        var countValue = text.Substring(leftIndex + 1, text.Length - 2 - leftIndex).Trim();
                        if (countValue.Length > 0)
                        {
                            if (long.TryParse(countValue, out long num))
                            {
                                if (num <= 0) Error = 1105;
                                else if (num > int.MaxValue) Error = 1106;
                                else count = (int)num;
                            }
                            else Error = 1107;
                        }
                    }
                    else if (leftIndex == 0) // 元组
                    {
                        kind = TypeKind.Tuple;
                        values = text.Substring(leftIndex + 1, text.Length - 2).Split(',');
                        count = values.Length;
                        if (count < 2) Error = 1109;
                    }
                    else Error = 1102;
                }
                else
                {
                    var splitCount = 0;
                    foreach (var ch in text) splitCount += ch == ':' ? 1 : 0;
                    if (splitCount != 1) Error = splitCount > 0 || text.IndexOf(',') != -1 ? 1103 : 1102;
                    else
                    {
                        kind = TypeKind.Dictionary;
                        values = text.Split(':');
                        if (values.Length != 2) Error = 1110;
                    }
                }

                if (Error == 0)
                {
                    for (int i = 0; i < values.Length; i++) values[i] = values[i].Trim();
                    foreach (var value in values)
                    {
                        if (Has(value)) items.Add(value);
                        else
                        {
                            Error = kind == TypeKind.Array ? 1104 : kind == TypeKind.Tuple ? 1108 : 1110;
                            break;
                        }
                    }
                }
            }
            Text = text.Replace(" ", "");
            Kind = kind;
            Items = new ReadOnlyCollection<string>(items);
            Count = count;
        }

        public (short error, string value) Parse(string value)
        {
            (short error, string value) outValue = (0, "");
            if (Kind == TypeKind.Simple)
            {
                if (Items[0] != __str) outValue = Parse(Items[0], value);
                else outValue.value = '"' + (value ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n") + '"';
            }
            else
            {
                // 字符串需用"包括，且字符串中包含"时需替换成\"，且字符串中包含\"时需替换成\\\"
                var values = new List<string>();
                for (int i = 0, li = 0, isIn = 0; i < value.Length; i++)
                {
                    if (value[i] == '"' && (i == 0 || value[i - 1] != '\\')) isIn = isIn == 0 ? 1 : 0;
                    if (isIn == 0 && (value[i] == ',' || i == value.Length - 1))
                    {
                        values.Add(value.Substring(li, i - li + (i == value.Length - 1 ? 1 : 0)));
                        li = i + 1;
                    }
                }

                if (Kind == TypeKind.Array || Kind == TypeKind.Tuple)
                {
                    while (Count != -1 && values.Count < Count) values.Add("");
                    if (Count != -1 && values.Count > Count) outValue.error = 1308;
                    else
                    {
                        for (int i = 0; i < values.Count; i++)
                        {
                            (short error, string value) res = Parse(Items[Kind == TypeKind.Array ? 0 : i], values[i]);
                            if (res.error == 0) outValue.value += (i > 0 ? ',' : "") + res.value;
                            else
                            {
                                outValue.error = res.error;
                                outValue.value = "[]";
                                break;
                            }
                        }
                        if (outValue.error == 0) outValue.value = '[' + outValue.value + ']';
                    }
                }
                else if (Kind == TypeKind.Dictionary)
                {
                    var keyValues = new List<(string k, string v)>();
                    foreach (var keyValue in values)
                    {
                        (string k, string v) item = ("", "");
                        for (int i = 0, isIn = 0; i < keyValue.Length; i++)
                        {
                            if (keyValue[i] == '"' && (i == 0 || keyValue[i - 1] != '\\')) isIn = isIn == 0 ? 1 : 0;
                            if (isIn == 0 && keyValue[i] == ':')
                            {
                                (short error, string value) keyRes = Parse(Items[0], keyValue.Substring(0, i));
                                (short error, string value) valueRes = Parse(Items[1], keyValue.Substring(i + 1));
                                if (keyRes.error == 0 && valueRes.error == 0) item = (keyRes.value, valueRes.value);
                                else outValue.error = keyRes.error == 0 ? keyRes.error : valueRes.error;
                                break;
                            }
                            if (i == keyValue.Length - 1) outValue.error = 1309;
                        }
                        if (outValue.error == 0)
                        {
                            if (keyValues.Exists(kv => kv.Item1 == item.k)) outValue.error = 1310;
                            else keyValues.Add(item);
                        }
                        if (outValue.error != 0) break;
                    }
                    if (outValue.error != 0) outValue.value = "[]";
                    else
                    {
                        foreach ((var k, var v) in keyValues)
                            outValue.value += (outValue.value.Length == 0 ? '[' : ",[") + k + ',' + v + ']';
                        outValue.value = '[' + outValue.value + ']';
                    }
                }
            }
            return outValue;
        }
        private (short error, string value) Parse(string type, string value)
        {
            (short error, string value) outValue = (0, "");
            value = (value ?? "").Trim();
            if (type == __str)
            {
                if (value == "") value = "\"\"";
                outValue.error = 1306;

                // 字符串需用"包括，且字符串中包含"时需替换成\"，且字符串中包含\"时需替换成\\\"
                if (value.Length > 1 && value[0] == '"' && value[value.Length - 2] != '\\' && value[value.Length - 1] == '"')
                {
                    var builder = new StringBuilder(value.Substring(1, value.Length - 2));
                    var isOk = true;
                    for (int i = 1; i < builder.Length; i++)
                    {
                        if (builder[i] == '\\')
                        {
                            if (i + 1 < builder.Length && builder[i + 1] == '"') i += 1;
                            else if (i + 3 < builder.Length && builder[i + 1] == '\\' && builder[i + 2] == '\\' && builder[i + 1] == '"') i += 3;
                            else builder.Insert(i++, '\\');
                        }
                        else if (builder[i] == '\r' || builder[i] == '\n')
                        {
                            builder.Insert(i++, '\\');
                            builder[i] = builder[i] == '\r' ? 'r' : 'n';
                        }
                        else if (builder[i] == '"')
                        {
                            isOk = false;
                            break;
                        }
                    }
                    if (isOk) outValue = (0, builder.Insert(0, '"').Append('"').ToString());
                }
            }
            else
            {
                if (value == "") value = "0";
                if (type == __u8 || type == __u16 || type == __u32 || type == __i8 || type == __i16 || type == __i32 || type == __i64)
                {
                    if (!long.TryParse(value, out long num)) outValue.error = 1304;
                    else
                    {
                        long min = type == __i8 ? sbyte.MinValue : type == __i16 ? short.MinValue : type == __i32 ? int.MinValue : type == __i64 ? long.MinValue : 0;
                        long max = type == __i8 ? sbyte.MaxValue : type == __i16 ? short.MaxValue : type == __i32 ? int.MaxValue : type == __i64 ? long.MaxValue : type == __u8 ? byte.MaxValue : type == __u16 ? ushort.MaxValue : uint.MaxValue;
                        if (num < min || num > max) outValue.error = 1303;
                        else outValue.value = num + "";
                    }

                }
                else if (type == __u64)
                {
                    if (!ulong.TryParse(value, out ulong num)) outValue.error = 1304;
                    else outValue.value = num + "";
                }
                else if (type == __f32)
                {
                    if (!float.TryParse(value, out float num)) outValue.error = 1305;
                    else outValue.value = num + "";
                }
                else if (type == __f64)
                {
                    if (!double.TryParse(value, out double num)) outValue.error = 1305;
                    else outValue.value = num + "";
                }
                else if (type == __bool)
                {
                    if (!int.TryParse(value, out int num) && num != 0 && num != 1) outValue.error = 1307;
                    else outValue.value = num == 1 ? "true" : "false";
                }
                else
                    outValue.error = 1399;
            }
            return outValue;
        }
    }

    class Sheet : Value
    {
        public ReadOnlyCollection<Index> Indexs { get; private set; }
        public ReadOnlyCollection<Type> Types { get; private set; }
        public ReadOnlyDictionary<string, ReadOnlyCollection<bool>> Outputs { get; private set; }
        public ReadOnlyCollection<string> Comments { get; private set; }
        public ReadOnlyCollection<ReadOnlyCollection<string>> Rows { get; private set; }

        public string Reference { get; private set; }

        public void Read(XlsxTextReader.Worksheet worksheet)
        {
            var indexs = new List<(short column, Index index)>();
            var types = new List<Type>();
            var outputs = new Dictionary<string, ReadOnlyCollection<bool>>();
            var comments = new List<string>();
            var rows = new List<ReadOnlyCollection<string>>();
            var step = 1;

            Error = 0;
            var rowIndexs = new List<string>();
            foreach (List<XlsxTextReader.Cell> cells in worksheet.Read())
            {
                if (step == 1/*indexs*/)
                {
                    if (cells[0].Reference.Row != 1) break; // 空表
                    foreach (var cell in cells)
                    {
                        if (cell is XlsxTextReader.MergeCell)
                        {
                            Error = 1004;
                            Reference = cell.Reference.Value;
                            break;
                        }
                        if (cell.Reference.Column == 1 || cell.Value.Trim() == "") continue;

                        var index = new Index(cell);
                        if (index.Error != 0)
                        {
                            Error = index.Error;
                            Reference = cell.Reference.Value;
                            break;
                        }
                        indexs.Add((cell.Reference.Column, index));
                    }
                    if (indexs.Count == 0) break; // 无输出内容
                    worksheet.MaxColumn = indexs[indexs.Count - 1].column;
                    step++;
                }
                else if (step == 2/*types*/)
                {
                    for (int i = 0, j = 0; i < indexs.Count; i++)
                    {
                        while (j < cells.Count && cells[j].Reference.Column < indexs[i].column) j++;
                        if (j >= cells.Count || cells[j].Reference.Column > indexs[i].column)
                        {
                            Error = 1101;
                            Reference = new XlsxTextReader.Reference(cells[0].Reference.Row, indexs[i].column).Value;
                            break;
                        }
                        var type = new Type(cells[j]);
                        if (type.Error == 0)
                        {
                            if (indexs[i].index.IsPrimary && type.Kind != Type.TypeKind.Simple)
                            {
                                Error = 1111;
                                Reference = cells[j].Reference.Value;
                                break;
                            }
                        }
                        else
                        {
                            Error = type.Error;
                            Reference = cells[j].Reference.Value;
                            break;
                        }
                        types.Add(type);
                    }
                    step++;
                }
                else if (step == 3/*outputs*/)
                {
                    var flag = cells[0].Value.Trim();
                    if (cells[0].Reference.Column != 1 || flag == "*" || flag == "" || flag == "-")
                    {
                        if (outputs.Count == 0)   // 无导出内容，空表
                        {
                            indexs.Clear();
                            break;
                        }
                        step++;
                    }
                    else
                    {
                        if (cells[0] is XlsxTextReader.MergeCell)
                        {
                            Error = 1201;
                            Reference = cells[0].Reference.Value;
                            break;
                        }
                        for (int i = 1; i < cells.Count; i++)
                        {
                            var value = cells[i].Value.Trim();
                            if (value != "" && value != "-")
                            {
                                Error = 1203;
                                Reference = cells[i].Reference.Value;
                                break;
                            }
                        }
                        if (Error != 0) break;

                        var isOuts = indexs.ConvertAll(_ => true);
                        for (int i = 0, j = 0; i < indexs.Count; i++)
                        {
                            while (j < cells.Count && cells[j].Reference.Column < indexs[i].column) j++;
                            if (j < cells.Count && cells[j].Reference.Column == indexs[i].column) isOuts[i] = cells[j].Value.Trim() != "-";
                        }
                        if (outputs.ContainsKey(flag))
                        {
                            Error = 1202;
                            Reference = cells[0].Reference.Value;
                            break;
                        }
                        if (isOuts.Contains(true)) outputs[flag] = new ReadOnlyCollection<bool>(isOuts);
                    }
                }
                if (step == 4/*comments*/)
                {
                    var flag = cells[0].Value.Trim();
                    comments = indexs.ConvertAll(_ => "");
                    if (cells[0].Reference.Column != 1 || flag == "" || flag == "-") step++;
                    else
                    {
                        if (cells[0].Value.Trim() != "*")
                        {
                            Error = 1204;
                            Reference = cells[0].Reference.Value;
                            break;
                        }
                        for (int i = 0, j = 0; i < indexs.Count; i++)
                        {
                            while (j < cells.Count && cells[j].Reference.Column < indexs[i].column) j++;
                            comments[i] = j < cells.Count && cells[j].Reference.Column == indexs[i].column ? cells[j].Value.Trim() : "";
                        }
                        step++;
                        continue;
                    }
                }
                if (step == 5/*rows*/)
                {
                    var flag = cells[0].Value.Trim();
                    if (cells[0].Reference.Column == 1)
                    {
                        if (flag == "-") continue;
                        if (flag != "")
                        {
                            Error = 1203;
                            Reference = cells[0].Reference.Value;
                            break;
                        }
                        cells.RemoveAt(0);
                    }
                    if (cells.Count > 0)
                    {
                        var row = new List<string>();
                        var rowIndex = "";
                        for (int i = 0, j = 0; i < indexs.Count; i++)
                        {
                            while (j < cells.Count && cells[j].Reference.Column < indexs[i].column) j++;
                            var cellValue = j < cells.Count && cells[j].Reference.Column == indexs[i].column ? cells[j].Value.Trim() : "";
                            var (error, value) = types[i].Parse(cellValue);
                            if (error != 0)
                            {
                                Error = error;
                                Reference = new XlsxTextReader.Reference(cells[0].Reference.Row, indexs[i].column).Value;
                                break;
                            }
                            if (indexs[i].index.IsPrimary) rowIndex += value;
                            row.Add(value);
                        }
                        if (Error != 0) break;
                        if (rowIndexs.IndexOf(rowIndex) != -1)
                        {
                            Error = 1350;
                            Reference = new XlsxTextReader.Reference(cells[0].Reference.Row, 1).Value;
                            break;
                        }
                        rows.Add(new ReadOnlyCollection<string>(row));
                        rowIndexs.Add(rowIndex);
                    }
                }
            }

            Text = worksheet.Name;
            if (indexs.Count > 0 && !Regex.IsMatch(Text, @"^[a-zA-Z_][a-zA-Z\d_]*$")) Error = 2001;
            if (Error == 0)
            {
                Indexs = new ReadOnlyCollection<Index>(indexs.ConvertAll(index => index.index));
                Types = new ReadOnlyCollection<Type>(types);
                Outputs = new ReadOnlyDictionary<string, ReadOnlyCollection<bool>>(outputs);
                Comments = new ReadOnlyCollection<string>(comments);
                Rows = new ReadOnlyCollection<ReadOnlyCollection<string>>(rows);
            }
        }
    }
}