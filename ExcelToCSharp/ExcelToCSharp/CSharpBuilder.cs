using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToCSharp
{
    public class CSharpBuilder
    {
        private readonly IEnumerable<Worksheet> _worksheets;
        private IEnumerable<IEnumerable<string>> _values;
        public string ClassName { get; set; }
        public List<Column> Columns { get; set; }

        public CSharpBuilder(IEnumerable<Worksheet> worksheets, bool convertToPascalCase)
        {
            _worksheets = worksheets;
            OpenWorksheet(1, convertToPascalCase);
        }

        public void OpenWorksheet(int index, bool convertToPascalCase)
        {
            var worksheet = _worksheets.ElementAt(index - 1);
            _values = worksheet.Rows.Skip(1);
            Columns = worksheet.Rows.First()
                .Select((h, i) => new Column(convertToPascalCase ? h.ToPascalCase() : h, _values.All(v => (v.Count() <= i) || Regex.IsMatch(v.ElementAt(i), @"^\d{2}/\d{2}/\d{4}$"))
                    ? "DateTme"
                    : _values.All(v => (v.Count() <= i) || int.TryParse(v.ElementAt(i), out var n))
                        ? "int"
                        : _values.All(v => (v.Count() <= i) || decimal.TryParse(v.ElementAt(i), out var d))
                            ? "decimal" 
                            : "string")).ToList();
        }

        public string GetClassDefinition(BackgroundWorker worker)
        {
            var values = _values.ToList();
            var selectedColumns = Columns.Where(c => c.Include).ToList();

            var cSharp = selectedColumns
                .Aggregate(new StringBuilder($"public class {ClassName}\n{{\n"), 
                    (sb, v) => sb.Append($"     public {v.Type} {v.Name.ToPascalCase()} {{ get; set; }}{Environment.NewLine}"))
                .Append("}");
            worker.ReportProgress(100);
            return cSharp.ToString();
        }

        public string GetJson(BackgroundWorker worker)
        {
            return "";
        }
    }
}
