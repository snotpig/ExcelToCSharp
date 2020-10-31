using System;
using System.Linq;

namespace ExcelToCSharp
{
	static class StringExtensions
	{
		public static string ToPascalCase(this string str)
		{
			return string.Join("", str.LTrimNonAlpha().Split(new[] { "_", " " }, StringSplitOptions.RemoveEmptyEntries)
				.Select(s => $"{char.ToUpper(s[0])}{s.Substring(1)}"));
		}

		public static string ToJsonDateTime(this string str)
		{
			var dateString = new DateTime(int.Parse(str.Substring(6, 4)), int.Parse(str.Substring(3, 2)), 
				int.Parse(str.Substring(0, 2)), int.Parse(str.Substring(11, 2)), int.Parse(str.Substring(14, 2)), 0)
					.ToString("yyyy-MM-ddTHH:mm:ss");

			return $"\"{dateString}\"";
		}

		private static string LTrimNonAlpha(this string str)
		{
			return string.Join("", str.Aggregate("", (t, v) => $"{t}{(!string.IsNullOrEmpty(t) || char.IsLetter(v) ? v.ToString() : "")}"));
		}
	}
}
