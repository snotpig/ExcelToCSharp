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
        public string ClassName { get; set; }
        public List<Column> Columns { get; set; }

        public CSharpBuilder(IEnumerable<Worksheet> worksheets)
        {
            _worksheets = worksheets;
        }

        public void OpenWorksheet(int index, bool convertToPascalCase, bool ignoreEmptyHeaders)
        {
            var worksheet = _worksheets.ElementAt(index);
            _values = worksheet.Rows.Skip(1);
            Columns = worksheet.Rows.First()
                .Select((h, i) => new Column(convertToPascalCase ? h.ToPascalCase() : h, GetType(i), i))
                .Where(c => !ignoreEmptyHeaders || !string.IsNullOrEmpty(c.Name)).ToList();
        }

        public string GetClassDefinition(BackgroundWorker worker)
        {
            worker.ReportProgress(50, "class");
            var values = _values.ToList();
            var selectedColumns = Columns.Where(c => c.Include).ToList();

            var cSharp = selectedColumns
                .Aggregate(new StringBuilder($"public class {ClassName}\n{{\n"), 
                    (sb, v) => sb.Append($"     public {v.Type} {v.Name} {{ get; set; }}{Environment.NewLine}"))
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
                case "int?":
                case "decimal?":
                    return string.IsNullOrWhiteSpace(value) ? "null" : value;
                case "int":
                case "decimal":
                    return string.IsNullOrWhiteSpace(value) ? "0" : value;
                case "string":
                    return $"\"{value}\"";
                case "DateTime":
                    return $"{value.ToJsonDateTime()}";
                default:
                    throw new InvalidOperationException($"Unecognised type: {type}");
            }
        }

        private string GetType(int i)
        {
            var values = _values.Where(v => v.Count() > i).Select(v => v.ElementAt(i));

            if (IsNullableDatetime(values))
                return "DateTime";
            if (IsDatetime(values))
                return "DateTime?";
            if (IsNullableInt(values))
                return "int";
            if (IsNullableInt(values))
                return "int?";
            if (IsNullableDecimal(values))
                return "decimal";
            if (IsDecimal(values))
                return "decimal?";
            return "string";
        }

        private bool IsDatetime(IEnumerable<string> values)
            => values.All(v => Regex.IsMatch(v, @"^\d{2}/\d{2}/\d{4}"));

        private bool IsNullableDatetime(IEnumerable<string> values)
            => values.Any(v => Regex.IsMatch(v, @"^\d{2}/\d{2}/\d{4}"))
                && values.All(v => string.IsNullOrEmpty(v) || v.ToLower() == "null" || Regex.IsMatch(v, @"^\d{2}/\d{2}/\d{4}"));

        private bool IsInt(IEnumerable<string> values)
            => values.All(v => int.TryParse(v, out var n));

        private bool IsNullableInt(IEnumerable<string> values)
            => values.Any(v => int.TryParse(v, out var n))
                && values.All(v => string.IsNullOrEmpty(v) || v.ToLower() == "null" || int.TryParse(v, out var n));

        private bool IsNullableDecimal(IEnumerable<string> values)
            => values.Any(v => decimal.TryParse(v, out var d))
                && values.All(v => string.IsNullOrEmpty(v) || v.ToLower() == "null" || decimal.TryParse(v, out var d));

        private bool IsDecimal(IEnumerable<string> values)
            => values.All(v => decimal.TryParse(v, out var d));
    }
}
