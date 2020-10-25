using System;
using System.Linq;

namespace ExcelToCSharp
{
	static class StringExtensions
	{
		public static string ToPascalCase(this string str)
		{
			return string.Join("", str.Split(new[] { "_" }, StringSplitOptions.RemoveEmptyEntries)
				.Select(s => $"{char.ToUpper(s[0])}{s.Substring(1)}"));
		}
	}
}
