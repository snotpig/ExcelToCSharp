using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelToCSharp
{
	static class Extensions
	{
		// string extensions
		public static string ToPascalCase(this string str)
		{
			str = Regex.Replace(str.ReplaceSymbols().LTrimNonAlpha(), "[a-z][A-Z0-9]", m => $"{m.Value[0]}-{m.Value[1]}");

			return string.Join("", str
				.Split(new[] { "_", " ", "(", ")", "-", "|", ":", ".", "/" }, StringSplitOptions.RemoveEmptyEntries)
				.Select(s => $"{char.ToUpper(s[0])}{s.Substring(1).ToLower()}"));
		}

		public static string ToJsonDateTime(this string str)
		{
			if (string.IsNullOrEmpty(str) || str.ToLower() == "null")
				return null;

			var dateString = new DateTime(int.Parse(str.Substring(6, 4)), int.Parse(str.Substring(3, 2)),
				int.Parse(str.Substring(0, 2)), int.Parse(str.Substring(11, 2)), int.Parse(str.Substring(14, 2)), 0)
					.ToJsonDateTimeString();

			return $"\"{dateString}\"";
		}

		public static bool IsId(this string str)
		{
			return str.Length > 1 && str.Substring(str.Length - 2).ToLower() == "id";
		}

		private static string LTrimNonAlpha(this string str)
		{
			return string
				.Join("", str.Aggregate("", (t, v) => $"{t}{(!string.IsNullOrEmpty(t) || char.IsLetter(v) ? v.ToString() : "")}"));
		}

		private static string ReplaceSymbols(this string str)
		{
			return str.Replace("%", "Percent").Replace("£", "GBP").Replace("&", "And")
				.Replace("PDSID", "PdsId").Replace("ID", "Id");
		}

		// DateTime extensions
		public static string ToJsonDateTimeString(this DateTime dt)
		{
			return dt.ToString("yyyy-MM-ddTHH:mm:ss");
		}
	}
}
