using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Threading;

namespace ExcelToCSharp
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private readonly string[] _extensions = { ".xlsx", ".xls", ".csv" };
		private CSharpBuilder _cSharpBuilder;
		private string _filePath;
		private IEnumerable<Worksheet> _worksheets;
		private BackgroundWorker _backgroundWorker;
		private DispatcherTimer _timer;
		private string _code;

		public MainWindow()
		{
			InitializeComponent();
			_backgroundWorker = (BackgroundWorker)FindResource("backgroundWorker");
			_timer = (DispatcherTimer)FindResource("timer");
			_timer.Interval = TimeSpan.FromSeconds(0.3);

			if (App.Args != null)
			{
				_filePath = App.Args[0];
				LoadFile();
			}
		}

		private void LoadFile()
		{
			PanelWorksheet.Visibility = Visibility.Collapsed;
			panelTableName.Visibility = Visibility.Collapsed;
			dgColumns.Visibility = Visibility.Collapsed;
			ShowButtons(false);
			var extension = Path.GetExtension(_filePath).ToLower();
			if (!_extensions.Contains(extension))
			{
				ShowMessage($"Can't open {extension} files");
				return;
			}
			Spinner.Visibility = Visibility.Visible;
			_backgroundWorker.RunWorkerAsync(new[] { "file", $"{((cbPascal.IsChecked ?? false) ? "pascal" : "")}" });
		}

		private void PopulateGrid()
		{
			if (_worksheets == null)
			{
				ShowMessage($"Can't open {_filePath}");
				return;
			}
			_cSharpBuilder = new CSharpBuilder(_worksheets, cbPascal.IsChecked ?? false);
			if (_worksheets.Count() > 1)
			{
				var worksheetNames = _worksheets.Select((w, i) => $" {i + 1}   {w.Title}");
				ComboWorksheet.ItemsSource = worksheetNames;
				ComboWorksheet.SelectedItem = worksheetNames.First();
				PanelWorksheet.Visibility = Visibility.Visible;
			}
			dgColumns.ItemsSource = _cSharpBuilder.Columns;
			dgColumns.Visibility = Visibility.Visible;
			panelTableName.Visibility = Visibility.Visible;
			ShowButtons();
			Resize();
		}

		public void ShowMessage(string message)
		{
			MessageBox.Show(message, "Error!");
		}

		private void Resize()
		{
			SizeToContent = SizeToContent.Height;
			MaxHeight = 200 + _cSharpBuilder.Columns.Count() * 19 + (_worksheets.Count() > 1 ? 26 : 0);
		}

		private void GenerateCode(string type)
		{
			if (_cSharpBuilder.Columns.Any(c => c.Include))
			{
				btnClass.IsEnabled = false;
				progress.Value = 0;
				SizeToContent = SizeToContent.Manual;
				ShowButtons(false);
				progress.Visibility = Visibility.Visible;
				SizeToContent = SizeToContent.Height;
				_cSharpBuilder.ClassName = txtClassName.Text;
				_backgroundWorker.RunWorkerAsync(new[] { type });
			}
			else
				ShowMessage("You must include at least one column");
		}

		private void ShowButtons(bool show = true)
		{
			var visibility = show ? Visibility.Visible : Visibility.Collapsed;
			btnClass.Visibility = visibility;
			btnJson.Visibility = visibility;
		}

		private void EnableButtons(bool enable = true)
		{
			btnClass.IsEnabled = enable;
			btnJson.IsEnabled = enable;
		}
		private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
		{
			var args = e.Argument as string[];

			if (args[0] == "c#")
			{
				_code = _cSharpBuilder.GetClassDefinition(sender as BackgroundWorker);
				e.Result = "code";
			}
			else if (args[0] == "json")
			{
				_code = _cSharpBuilder.GetJson(sender as BackgroundWorker);
				e.Result = "code";
			}
			else
			{
				_worksheets = FileReader.ReadFile(_filePath, ShowMessage);
				e.Result = "file";
			}
		}

		private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			if (e.Result as string == "code")
			{
				Clipboard.SetText(_code);
				_timer.Start();
			}
			else
			{
				Spinner.Visibility = Visibility.Collapsed;
				PopulateGrid();
			}
		}

		private void progressChanged(object sender, ProgressChangedEventArgs e)
		{
			progress.Value = e.ProgressPercentage;
		}

		private void timer_Tick(object sender, EventArgs e)
		{
			SizeToContent = SizeToContent.Manual;
			progress.Visibility = Visibility.Collapsed;
			ShowButtons();
			SizeToContent = SizeToContent.Height;
			EnableButtons();
		}

		private void BtnOpen_Click(object sender, RoutedEventArgs e)
		{
			var openFileDialog = new OpenFileDialog();

			openFileDialog.Filter = $"Excel files (*{string.Join("; *", _extensions)})|*{string.Join("; *", _extensions)}";
			var result = openFileDialog.ShowDialog();
			if (result.Value)
			{
				_filePath = openFileDialog.FileName;
				LoadFile();
			}
		}

		private void DropPanel_Drop(object sender, DragEventArgs e)
		{
			var files = (string[])e.Data.GetData(DataFormats.FileDrop);
			if (files != null)
			{
				_filePath = files[0];
				LoadFile();
			}
		}

		private void CbAll_Checked(object sender, RoutedEventArgs e)
		{
			_cSharpBuilder.Columns.ForEach(c => c.Include = true);
		}

		private void CbAll_Unchecked(object sender, RoutedEventArgs e)
		{
			_cSharpBuilder.Columns.ForEach(c => c.Include = false);
		}

		private void ComboSheet_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
		{
			_cSharpBuilder.OpenWorksheet(ComboWorksheet.SelectedIndex + 1, cbPascal.IsChecked ?? false);
			dgColumns.ItemsSource = _cSharpBuilder.Columns;
			Resize();
		}

		private void BtnClass_Click(object sender, RoutedEventArgs e)
		{
			GenerateCode("c#");
		}
		private void BtnJson_Click(object sender, RoutedEventArgs e)
		{
			GenerateCode("json");
		}

		private void CbPascal_Click(object sender, RoutedEventArgs e)
		{
			if(!string.IsNullOrEmpty(_filePath))
				PopulateGrid();
		}
	}
}
