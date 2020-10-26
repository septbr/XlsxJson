using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

namespace XlsxJson
{
    class Program
    {
        struct Config
        {
            public List<string> xlsxs;
            public string outDir;
        }
        static bool ParseCommand(string[] args, out Config config)
        {
            config = new Config { xlsxs = new List<string>(), outDir = "." };

            int[] flags = { 0, 0, 0 };
            List<string> dirs = new List<string>();
            List<string> files = new List<string>();
            if (args.Length == 0) args = new string[1] { "-h" };
            foreach (string arg in args)
            {
                if (flags[0] == 0 && arg.Equals("-f"))
                {
                    flags[0] = 1;
                    if (flags[1] == 1) flags[1] = 2;
                    if (flags[2] == 1) flags[2] = 2;
                    continue;
                }
                if (flags[1] == 0 && arg.Equals("-d"))
                {
                    flags[1] = 1;
                    if (flags[0] == 1) flags[0] = 2;
                    if (flags[2] == 1) flags[2] = 2;
                    continue;
                }
                if (flags[2] == 0 && arg.Equals("-o"))
                {
                    flags[2] = 1;
                    if (flags[0] == 1) flags[0] = 2;
                    if (flags[1] == 1) flags[1] = 2;
                    continue;
                }
                if (arg.Equals("-h") || arg.StartsWith("-"))
                {
                    Console.WriteLine("repos: https://github.com/septeer/XlsxJson.git");
                    Console.WriteLine("Usage: XlsxJson [-d 文件夹1 文件夹2 ...文件夹n] [-f 文件1 文件2 ...文件n] [-o 输出文件夹]");
                    return false;
                }

                if (flags[0] == 1) files.Add(arg);
                else if (flags[1] == 1) dirs.Add(arg);
                else if (flags[2] == 1)
                {
                    config.outDir = arg;
                    flags[2] = 2;
                }
            }
            if (dirs.Count == 0 && files.Count == 0) return false;

            foreach (string path in files)
            {
                FileInfo file = new FileInfo(path);
                if (file.Exists && file.Extension.ToLower().Equals(".xlsx"))
                    config.xlsxs.Add(file.FullName);
            }
            foreach (string path in dirs)
            {
                if (!Directory.Exists(path)) continue;

                DirectoryInfo directory = new DirectoryInfo(path);
                foreach (FileInfo file in directory.GetFiles("*.xlsx"))
                    config.xlsxs.Add(file.FullName);
            }

            if (config.xlsxs.Count == 0)
            {
                Console.WriteLine("未找到文件！");
                return false;
            }

            return true;
        }

        /**
         * -f 指定文件
         * -d 指定文件夹
         * -o 输出文件夹
         * -h 查看帮助
         */
        static int Main(string[] args)
        {
            Console.Title = "XlsxJson https://github.com/septbr/XlsxJson.git";
            Console.OutputEncoding = new UTF8Encoding(false);

            Config config;
            if (!ParseCommand(args, out config)) return -1;

            var res = 0;
            var sheets = new List<(string xlsx, Sheet sheet)>();
            for (int i = 0; res == 0 && i < config.xlsxs.Count; i++)
            {
                var xlsx = config.xlsxs[i];

                Console.WriteLine("  ({0}/{1}) {2}", i + 1, config.xlsxs.Count, xlsx);
                Console.ForegroundColor = ConsoleColor.Red;
                try
                {
                    using (var workbook = XlsxTextReader.Workbook.Open(xlsx))
                    {
                        foreach (var worksheet in workbook.Read())
                        {
                            var sheet = new Sheet();
                            sheet.Read(worksheet);
                            if (sheet.Error != 0)
                            {
                                Console.WriteLine("      错误: {0}:{1} {2}", sheet.Text, sheet.Reference, sheet.ErrorInfo);
                                res = -1;
                                break;
                            }
                            if (sheet.Indexs.Count > 0)
                            {
                                var xlsx2 = sheets.Find(st => st.sheet.Text == sheet.Text).xlsx;
                                if (xlsx2 != null)
                                {
                                    Console.WriteLine("      错误: {0} 已存在于{1}", sheet.Text, xlsx2);
                                    res = -1;
                                    break;
                                }
                                sheets.Add((xlsx, sheet));
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    Console.WriteLine("      错误: " + ex.Message);
                    res = -1;
                }
                Console.ResetColor();
            }
            if (res != 0) return res;

            var outSheets = new Dictionary<string, List<(string sheetName, List<string> rows)>>();
            foreach ((var xlsx, var sheet) in sheets)
            {
                foreach ((var name, var isOuts) in sheet.Outputs)
                {
                    var rows = new List<string>();
                    var row = new StringBuilder();
                    for (int i = 0; i < sheet.Indexs.Count; i++)
                    {
                        (var index, var type, var comment) = (sheet.Indexs[i], sheet.Types[i], sheet.Comments[i]);
                        if (isOuts[i])
                        {
                            if (row.Length > 0) row.Append(", ");
                            row.Append("[\"").Append(index.Text).Append("\",\"").Append(type.Text).Append("\",").Append(index.IsPrimary ? 1 : 0).Append(",\"").Append(comment).Append("\"]");
                        }
                    }
                    rows.Add(row.Insert(0, '[').Append(']').ToString());

                    foreach (var datas in sheet.Rows)
                    {
                        row.Clear();
                        for (int i = 0; i < datas.Count; i++)
                        {
                            if (isOuts[i])
                            {
                                if (row.Length > 0) row.Append(", ");
                                row.Append(datas[i]);
                            }
                        }
                        rows.Add(row.Insert(0, '[').Append(']').ToString());
                    }

                    outSheets.TryAdd(name, new List<(string sheetName, List<string> rows)>());
                    outSheets[name].Add((sheet.Text, rows));
                }
            }

            var jsons = new Dictionary<string, StringBuilder>();
            foreach ((var name, var sheet) in outSheets)
            {
                jsons.TryAdd(name, new StringBuilder());
                var json = jsons[name];
                foreach ((var sheetName, var rows) in sheet)
                {
                    if (json.Length > 0) json.Append(",\n");
                    json.Append("    \"").Append(sheetName).Append("\": [");
                    var more = false;
                    foreach (var row in rows)
                    {
                        if (more) json.Append(',');
                        json.Append("\n        ").Append(row);
                        more = true;
                    }
                    json.Append("\n    ]");
                }
            }
            foreach ((_, var json) in jsons)
            {
                json.Insert(0, "{\n");
                json.Append("\n}");
            }

            Directory.CreateDirectory(config.outDir);
            foreach ((var name, var json) in jsons)
                File.WriteAllText(Path.Combine(config.outDir, name) + ".json", json.ToString());

            return res;
        }
    }
}