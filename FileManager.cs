using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace MyExcelMAUIApp
{
    public class FileManager
    {
        public string filePath { get; set; }

        public void SaveAs(Dictionary<string, cellProp> expressions, int CountColumn, int CountRow, string path)
        {
            filePath = path;
            var tableInfo = new TableInfo
            {

                Expressions = expressions,
                CountColumn = CountColumn,
                CountRow = CountRow
            };
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            var json = JsonSerializer.Serialize(tableInfo, options);
            File.WriteAllText(filePath, json);
        }

        public void Save(Dictionary<string, cellProp> expressions, int CountColumn, int CountRow)
        {
            SaveAs(expressions, CountColumn, CountRow, filePath);
        }

        public TableInfo Load(string path)
        {
            if (File.Exists(path))
            {
                var json = File.ReadAllText(path);
                return JsonSerializer.Deserialize<TableInfo>(json);
            }
            return null;
        }
    }

    public class TableInfo
    {
        public Dictionary<string, cellProp> Expressions { get; set; }
        public int CountColumn { get; set; }
        public int CountRow { get; set; }
    }
}
