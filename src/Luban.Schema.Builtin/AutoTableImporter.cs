
using Luban.Defs;
using Luban.RawDefs;
using Luban.Utils;
using System.Text.RegularExpressions;
using ExcelDataReader;
using System.Data;

namespace Luban.Schema.Builtin;

[TableImporter("auto")]
public class AutoTableImporter : ITableImporter
{
    private static readonly NLog.Logger s_logger = NLog.LogManager.GetCurrentClassLogger();
    public List<RawTable> LoadImportTables()
    {
        string dataDir = GenerationContext.GlobalConf.InputDataDir;
        string fileNamePatternStr = EnvManager.Current.GetOptionOrDefault("tableImporter", "filePattern", false, "#(.*)");
        string tableNamespaceFormatStr = EnvManager.Current.GetOptionOrDefault("tableImporter", "tableNamespaceFormat", false, "{0}");
        string tableNameFormatStr = EnvManager.Current.GetOptionOrDefault("tableImporter", "tableNameFormat", false, "Tbl{0}");
        string valueTypeNameFormatStr = EnvManager.Current.GetOptionOrDefault("tableImporter", "valueTypeNameFormat", false, "{0}");
        var fileNamePattern = new Regex(fileNamePatternStr);
        var excelExts = new HashSet<string> { "xlsx", "xls", "xlsm", "csv" };
        var tables = new List<RawTable>();
        foreach (string file in Directory.GetFiles(dataDir, "*", SearchOption.AllDirectories))
        {
            if (FileUtil.IsIgnoreFile(dataDir, file))
            {
                continue;
            }
            string fileName = Path.GetFileName(file);
            string ext = Path.GetExtension(fileName).TrimStart('.');
            if (!excelExts.Contains(ext))
            {
                continue;
            }
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            var match = fileNamePattern.Match(fileNameWithoutExt);
            if (!match.Success || match.Groups.Count <= 1)
            {
                continue;
            }

            string relativePath = file.Substring(dataDir.Length + 1).TrimStart('\\').TrimStart('/');
            string namespaceFromRelativePath = Path.GetDirectoryName(relativePath).Replace('/', '.').Replace('\\', '.');
            string rawTableFullName = match.Groups[1].Value;
            string rawTableNamespace = TypeUtil.GetNamespace(rawTableFullName);
            string rawTableName = TypeUtil.GetName(rawTableFullName);
            string tableNamespace = TypeUtil.MakeFullName(namespaceFromRelativePath, string.Format(tableNamespaceFormatStr, rawTableNamespace));
            string tableName = string.Format(tableNameFormatStr, rawTableName);
            string valueTypeFullName = TypeUtil.MakeFullName(tableNamespace, string.Format(valueTypeNameFormatStr, rawTableName));

            // 检查是否需要分类

            List<string> list = new List<string>();
            // 打开 Excel 文件，遍历所有 Sheet
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var dataSet = reader.AsDataSet();
                foreach (DataTable sheet in dataSet.Tables)
                {
                    match = fileNamePattern.Match(sheet.TableName); // # 开头的不导出
                    if (!match.Success || match.Groups.Count <= 1)
                    {
                        list.Add(sheet.TableName + "@" + relativePath);
                        continue;
                    }


                    var sheetName = match.Groups[1].Value;

                    string input = sheet.TableName + "@" + relativePath;
                    var table = new RawTable()
                    {
                        Namespace = tableNamespace,
                        Name = tableName + sheetName,
                        Index = "",
                        ValueType = valueTypeFullName + sheetName,
                        ReadSchemaFromFile = true,
                        Mode = TableMode.MAP,
                        Comment = "Import by auto",
                        Groups = new List<string> { },
                        InputFiles = new List<string> { input },
                        OutputFile = "",
                    };
                    s_logger.Info("import table file:{@}", input);
                    tables.Add(table);
                }
            }

            if (list.Count > 0)
            { 
                var table = new RawTable()
                    {
                        Namespace = tableNamespace,
                        Name = tableName ,
                        Index = "",
                        ValueType = valueTypeFullName ,
                        ReadSchemaFromFile = true,
                        Mode = TableMode.MAP,
                        Comment = "Import by auto",
                        Groups = new List<string> { },
                        InputFiles = list,
                        OutputFile = "",
                    };
                    s_logger.Info("import table file:{@}", list);
                    tables.Add(table);
            }
            ;
        }
        return tables;
    }
}
