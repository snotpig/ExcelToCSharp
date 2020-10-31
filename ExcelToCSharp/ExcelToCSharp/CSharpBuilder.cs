using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToCSharp
{
    public class CSharpBuilder
    {
        private readonly IEnumerable<Worksheet> _worksheets;
        private IEnumerable<IEnumerable<string>> _values;
        private bool _convertToPascalCase;
        public string ClassName { get; set; }
        public List<Column> Columns { get; set; }

        public CSharpBuilder(IEnumerable<Worksheet> worksheets)
        {
            _worksheets = worksheets;
        }

        public void OpenWorksheet(int index, bool convertToPascalCase, bool ignoreEmptyHeaders)
        {
            _convertToPascalCase = convertToPascalCase;
            var worksheet = _worksheets.ElementAt(index);
            _values = worksheet.Rows.Skip(1);
            Columns = worksheet.Rows.First()
                .Select((h, i) => new Column(convertToPascalCase ? h.ToPascalCase() : h, _values.All(v => (v.Count() <= i) || Regex.IsMatch(v.ElementAt(i), @"^\d{2}/\d{2}/\d{4}"))
                    ? "DateTime"
                    : _values.All(v => (v.Count() <= i) || int.TryParse(v.ElementAt(i), out var n))
                        ? "int"
                        : _values.All(v => (v.Count() <= i) || decimal.TryParse(v.ElementAt(i), out var d))
                            ? "decimal" 
                            : "string", i))
                .Where(c => !ignoreEmptyHeaders || !string.IsNullOrEmpty(c.Name)).ToList();
        }

        public string GetClassDefinition(BackgroundWorker worker)
        {
            worker.ReportProgress(50, "class");
            var values = _values.ToList();
            var selectedColumns = Columns.Where(c => c.Include).ToList();

            var cSharp = selectedColumns
                .Aggregate(new StringBuilder($"public class {ClassName}\n{{\n"), 
                    (sb, v) => sb.Append($"     public {v.Type} {(_convertToPascalCase ? v.Name.ToPascalCase() : v.Name)} {{ get; set; }}{Environment.NewLine}"))
                .Append("}");
            worker.ReportProgress(100, "class");
            return cSharp.ToString();
        }

        public string GetJson(BackgroundWorker worker)
        {
            worker.ReportProgress(50, "json");
            var values = _values.ToList();
            var selectedColumns = Columns.Where(c => c.Include).ToList();

            var json = values.Aggregate(new StringBuilder("[\n")
                , (sb, v) => sb.Append($"\t{{\n\t\t{GetObjects(selectedColumns, v)}\n\t}},\n"));
            json.Length -= 2;
            json.Append("\n]");
            worker.ReportProgress(100, "json");
            return json.ToString();
        }

        private string GetObjects(IEnumerable<Column> columns, IEnumerable<string> values)
        {
            var names = Columns.Select(c => c.Name).ToList();
            var objects = columns.Aggregate(new StringBuilder(), 
                (sb, c) => sb.Append($"\"{c.Name}\": {GetValue(c.Type, values.ElementAt(c.Index))}{(c.Name == names.Last() ? "" : $"\n\t\t")}"));

            return objects.ToString();
        }

        private string GetValue(string type, string value)
        {
            switch(type)
            {
                case "int":
                case "decimal":
                    return value;
                case "string":
                    return $"\"{value}\"";
                case "DateTime":
                    return $"{value.ToJsonDateTime()}";
                default:
                    throw new InvalidOperationException($"Unecognised type: {type}");
            }
        }
    }
}
