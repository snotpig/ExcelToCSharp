﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToCSharp
{
	public static class FileReader
	{
		public delegate void MessageDelegate(string message);
		private static MessageDelegate _showMessage;

		public static IEnumerable<Worksheet> ReadFile(string filePath, MessageDelegate msgDelegate)
		{
			_showMessage = msgDelegate;
			var worksheets = filePath.Substring(filePath.Length - 4).ToLower() == ".csv"
				? ReadCsvFile(filePath)
				: ReadXlsxFile(filePath);

			foreach (var worksheet in worksheets)
			{
				var numCols = worksheet.Rows.First().ToList().IndexOf(null);
				if (numCols >= 0)
					worksheet.Rows = worksheet.Rows.Select(r => r.Where((c, i) => i < numCols));
			}
			return worksheets;
		}

		private static IEnumerable<Worksheet> ReadCsvFile(string filePath)
		{
			ICollection<string> fileLines;
			try
			{
				fileLines = File.ReadAllLines(filePath);
			}
			catch
			{
				return null;
			}
			if (!fileLines.Any()) return null;
			var rows = fileLines.Select(l => string.Join("", l.Split(new[] { '"' })
				.Select((s, i) => i % 2 == 0 ? s : s.Replace(',', (char)176)))
				.Split(new[] { ',' })
				.Select(s => s.Replace((char)176, ',').Replace("\'", "\'\'"))).ToList();

			return new List<Worksheet> { new Worksheet { Rows = TrimEmpty(rows) } };
		}

		private static IEnumerable<Worksheet> ReadXlsxFile(string filePath)
		{
			var xlApp = new Excel.Application();
			Excel.Workbook xlWorkBook;
			try
			{
				xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
			}
			catch
			{
				xlApp.Quit();
				ReleaseObject(xlApp);
				return null;
			}
			var numSheets = xlWorkBook.Sheets.Count;
			var worksheets = new List<Worksheet>();

			for (var s = 1; s <= numSheets; s++)
			{
				var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(s);

				var range = xlWorkSheet.UsedRange;
				var cells = range.Value;
				var formulae = range.FormulaR1C1;

				var numRows = cells?.GetLength(0) ?? 0;
				var numCols = cells?.GetLength(1) ?? 0;
				List<List<string>> rows = null;

				if (numRows > 0 && numCols > 0)
				{
					rows = new List<List<string>>();
					for (var i = 1; i <= numRows; i++)
					{
						var row = new List<string> { };
						for (var j = 1; j <= numCols; j++)
						{
							if (cells[i, j] is int error && Enum.GetValues(typeof(ExcelError)).Cast<int>().Contains(error))
								row.Add(((ExcelError)cells[i, j]).ToString());

							else
							{
								var dt = cells[i, j] as DateTime?;
								row.Add(dt == null
									? cells[i, j]?.ToString() ?? ""
									: dt.Value.ToString("dd/MM/yyyy HH:mm:ss"));
							}
						}
						rows.Add(row);
					}
					worksheets.Add(new Worksheet { Title = xlWorkSheet.Name, Rows = TrimEmpty(rows) });
				}
				ReleaseObject(xlWorkSheet);
			}

			xlWorkBook.Close(true, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
			xlApp.Quit();
			ReleaseObject(xlWorkBook);
			ReleaseObject(xlApp);

			return worksheets;
		}

		private static void ReleaseObject(object obj)
		{
			try
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex)
			{
				obj = null;
				_showMessage("Unable to release the Object " + ex.ToString());
			}
			finally
			{
				GC.Collect();
			}
		}

		private static List<IEnumerable<string>> TrimEmpty(IEnumerable<IEnumerable<string>> rows)
		{
			return rows.Where(r => !r.All(v => string.IsNullOrEmpty(v))).ToList();
		}
	}

	enum ExcelError
	{
		ErrDiv0 = -2146826281,
		ErrNA = -2146826246,
		ErrName = -2146826259,
		ErrNull = -2146826288,
		ErrNum = -2146826252,
		ErrRef = -2146826265,
		ErrValue = -2146826273,
	}
}
