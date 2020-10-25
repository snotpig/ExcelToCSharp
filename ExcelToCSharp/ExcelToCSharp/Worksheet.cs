using System.Collections.Generic;

namespace ExcelToCSharp
{
    public class Worksheet
    {
        public string Title { get; set; }
        public IEnumerable<IEnumerable<string>> Rows { get; set; }
    }
}
