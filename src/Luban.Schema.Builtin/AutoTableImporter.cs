﻿using Luban.Defs;
using Luban.RawDefs;
using Luban.Utils;
using System.Text.RegularExpressions;
using ExcelDataReader;
using System.Data;
using System.Globalization;

namespace Luban.Schema.Builtin;

[TableImporter("auto")]
public class AutoTableImporter : ITableImporter
{
    private static readonly NLog.Logger s_logger = NLog.LogManager.GetCurrentClassLogger();



    static string CapitalizeFirstLetter(string str)
    {
        if (string.IsNullOrEmpty(str))
        {
            return str;
        }
        // 使用 CultureInfo 处理首字母大写
        TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
        return textInfo.ToTitleCase(str.ToLower());  // 将整个字符串变为标题形式，首字母大写
    }
    public List<RawTable> LoadImportTables()
    {
        string dataDir = GenerationContext.GlobalConf.InputDataDir;
        string tableNamespaceFormatStr = EnvManager.Current.GetOptionOrDefault("tableImporter", "tableNamespaceFormat", false, "{0}");
        string tableNameFormatStr = EnvManager.Current.GetOptionOrDefault("tableImporter", "tableNameFormat", false, "Tb{0}");
        string valueTypeNameFormatStr = EnvManager.Current.GetOptionOrDefault("tableImporter", "valueTypeNameFormat", false, "{0}");
        var fileNamePattern = new Regex(EnvManager.Current.GetOptionOrDefault("tableImporter", "filePattern", false, "^[A-Za-z0-9]+$"));
        var sheetNamePattern = new Regex(EnvManager.Current.GetOptionOrDefault("tableImporter", "sheetPattern", false, @"^[A-Za-z0-9_]+(?:\|([A-Za-z0-9_]+))?$"));

        var excelExts = new HashSet<string> { "xlsx", "xls", "xlsm", "csv" };
        var tables = new List<RawTable>();
        foreach (string file in Directory.GetFiles(dataDir, "*", SearchOption.AllDirectories))
        {
            if (FileUtil.IsIgnoreFile(dataDir, file))
            {
                s_logger.Info("Ignore file {0}", file);
                continue;
            }
            string fileName = Path.GetFileName(file);
            string ext = Path.GetExtension(fileName).TrimStart('.');
            if (!excelExts.Contains(ext))
            {
                continue;
            }
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            if (!fileNamePattern.IsMatch(fileNameWithoutExt))
            {
                s_logger.Info("Ignore file {0}", fileName);
                continue;
            }

            string relativePath = file.Substring(dataDir.Length + 1).TrimStart('\\').TrimStart('/');
            string namespaceFromRelativePath = Path.GetDirectoryName(relativePath).Replace('/', '.').Replace('\\', '.');
            string rawTableFullName = fileNameWithoutExt;
            string rawTableNamespace = TypeUtil.GetNamespace(rawTableFullName);
            string rawTableName = TypeUtil.GetName(rawTableFullName);
            string tableNamespace = TypeUtil.MakeFullName(namespaceFromRelativePath, string.Format(tableNamespaceFormatStr, rawTableNamespace));
            string tableName = string.Format(tableNameFormatStr, rawTableName);
            string valueTypeFullName = TypeUtil.MakeFullName(tableNamespace, string.Format(valueTypeNameFormatStr, rawTableName));
            // 检查是否需要分类

            // 打开 Excel 文件，遍历所有 Sheet
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var dataSet = reader.AsDataSet();
                Dictionary<string, List<string>> groups = new Dictionary<string, List<string>>();
                Dictionary<string, string> valueType = new Dictionary<string, string>();
                foreach (DataTable sheet in dataSet.Tables)
                {
                    var match = sheetNamePattern.Match(sheet.TableName); // # 开头的不导出
                    if (!match.Success)
                    {
                        s_logger.Info("Ignore file {0} @ sheet {1}", file, sheet.TableName);
                        continue;
                    }

                    var typeName = (match.Groups.Count > 1 && !string.IsNullOrEmpty(match.Groups[1].Value)) ? CapitalizeFirstLetter(match.Groups[1].Value) : CapitalizeFirstLetter(sheet.TableName);
                    var groupName = tableName + typeName;
                    if (!groups.ContainsKey(groupName))
                    {
                        groups[groupName] = new List<string>();  // 创建新的 List<string>
                        valueType[groupName] = valueTypeFullName + typeName;
                    }
                    groups[groupName].Add(sheet.TableName + "@" + relativePath);
                }

                foreach (var kvp in groups)
                {
                    var table = new RawTable()
                    {
                        Namespace = tableNamespace,
                        Name = kvp.Key,
                        ValueType = valueType[kvp.Key],
                        ReadSchemaFromFile = true,
                        Mode = TableMode.MAP,
                        Comment = "Import by auto",
                        InputFiles = kvp.Value,
                    };
                    s_logger.Info("Import {0} from {1}", kvp.Key, string.Join(" ", kvp.Value));
                    tables.Add(table);
                }
            }
        }
        return tables;
    }
}
