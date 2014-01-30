using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using CsvHelper;
using Microsoft.Win32;

namespace TableExtractor
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DocumentBox.SelectAll();
            DocumentBox.Focus();
        }

        private string ExtractText(TextElement element)
        {
            TextRange range = new TextRange(element.ContentStart, element.ContentEnd);
            return range.Text;
        }

        private bool InternalExtractTables()
        {
            int columnCount = 0;
            int.TryParse(ColumnCountBox.Text, out columnCount);
            if (columnCount <= 0)
            {
                throw new Exception("The column parameter should be a positive integer");
            }
            var regexes = Regex.Split(RegexesBox.Text, Environment.NewLine, RegexOptions.Multiline).Select(s => s.Trim()).ToList();
            if (regexes.Count > columnCount)
            {
                throw new Exception("There are more regular expressions than columns");
            }
            while (regexes.Count < columnCount)
            {
                regexes.Add(string.Empty);
            }
            // Select all rows from tables having the right number of columns.
            var rows = DocumentBox.Document.Blocks.OfType<Table>().
                Where(tb => tb.Columns.Count == columnCount).
                SelectMany(tb => tb.RowGroups.SelectMany(g => g.Rows)).
                Where(r => r.Cells.Count == columnCount);
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Comma separated values (*.csv)|*.csv";
            bool result = dialog.ShowDialog(this).GetValueOrDefault(false);
            if (result)
            {
                var stream = dialog.OpenFile();
                using (var textWriter = new StreamWriter(stream))
                {
                    CsvWriter writer = new CsvWriter(textWriter);
                    string[] data = new string[columnCount];
                    foreach (var row in rows)
                    {
                        bool validRow = true;
                        for (int i = 0; validRow && i < columnCount; ++i)
                        {
                            data[i] = ExtractText(row.Cells[i]);
                            validRow = regexes[i] == string.Empty || Regex.IsMatch(data[i], regexes[i]);
                        }
                        if (validRow)
                        {
                            for (int i = 0; i < columnCount; ++i)
                            {
                                writer.WriteField(data[i]);
                            }
                            writer.NextRecord();
                        }
                    }
                }
            }
            return result;
        }

        private void OnExtractTables(object sender, RoutedEventArgs e)
        {
            try
            {
                if (InternalExtractTables())
                {
                    MessageBox.Show("Extraction completed!", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, ex.Source, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}