using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Controls;
using Microsoft.Win32;
using ClosedXML.Excel;
// DocumentFormat.OpenXml namespaces
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = System.Windows.Documents.Paragraph; // Disambiguate from OpenXml Paragraph
using Run = System.Windows.Documents.Run;             // Disambiguate from OpenXml Run

namespace AbbreviationReplacer
{
    public partial class MainWindow : Window
    {
        // Dictionary: Full Form -> List of possible Abbreviations
        private Dictionary<string, List<string>> fullToAbbrDict = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        public MainWindow()
        {
            InitializeComponent();
            LoadAbbreviationsFromExcel();
        }

        #region Abbreviation Logic & Excel Management

        private string GetExcelPath()
        {
            string exeFolder = AppDomain.CurrentDomain.BaseDirectory;
            string dataPath = System.IO.Path.Combine(exeFolder, @"..\..\data\abbr.xlsx");
            return System.IO.Path.GetFullPath(dataPath);
        }

        private void LoadAbbreviationsFromExcel()
        {
            string path = GetExcelPath();
            if (!File.Exists(path))
            {
                // Note: Silent fail or placeholder until user uploads a file
                return;
            }

            fullToAbbrDict.Clear();
            try
            {
                using (var workbook = new XLWorkbook(path))
                {
                    var ws = workbook.Worksheets.First();
                    foreach (var row in ws.RowsUsed())
                    {
                        string shortForm = row.Cell(1).GetString().Trim();
                        // Iterate through columns 2 to end for full forms
                        foreach (var cell in row.Cells(2, ws.LastColumnUsed().ColumnNumber()))
                        {
                            string full = cell.GetString().Trim();
                            if (string.IsNullOrEmpty(full)) continue;

                            if (!fullToAbbrDict.ContainsKey(full))
                                fullToAbbrDict[full] = new List<string>();

                            if (!fullToAbbrDict[full].Contains(shortForm))
                                fullToAbbrDict[full].Add(shortForm);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel: {ex.Message}");
            }
        }

        private void UploadExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string dest = GetExcelPath();
                    string dir = Path.GetDirectoryName(dest);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                    File.Copy(openFileDialog.FileName, dest, true);
                    LoadAbbreviationsFromExcel();
                    MessageBox.Show("Abbreviation dictionary updated!");
                }
                catch (Exception ex) { MessageBox.Show($"Upload failed: {ex.Message}"); }
            }
        }

        #endregion

        #region Word Document (DOCX) Integration

        private void UploadDocx_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Word Documents (*.docx)|*.docx" };
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(openFileDialog.FileName, false))
                    {
                        var body = wordDoc.MainDocumentPart.Document.Body;
                        // Simple inner text extraction for the input box
                        InputTextBox.Text = body.InnerText;
                    }
                }
                catch (Exception ex) { MessageBox.Show($"Could not read Word file: {ex.Message}"); }
            }
        }

        private void DownloadDocx_Click(object sender, RoutedEventArgs e)
        {
            if (OutputRichTextBox.Document.Blocks.Count == 0) return;

            SaveFileDialog saveFileDialog = new SaveFileDialog { Filter = "Word Document (*.docx)|*.docx", FileName = "Processed_Document.docx" };
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(saveFileDialog.FileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                        mainPart.Document = new Document(new Body());
                        var body = mainPart.Document.Body;

                        // Create a Word paragraph
                        var wPara = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();

                        // Capture the current text in the RichTextBox (with user selections)
                        TextRange range = new TextRange(OutputRichTextBox.Document.ContentStart, OutputRichTextBox.Document.ContentEnd);
                        var wRun = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(range.Text));

                        wPara.AppendChild(wRun);
                        body.AppendChild(wPara);
                        mainPart.Document.Save();
                    }
                    MessageBox.Show("Word document saved successfully!");
                }
                catch (Exception ex) { MessageBox.Show($"Save failed: {ex.Message}"); }
            }
        }

        #endregion

        #region Replacement Engine

        private void ReplaceButton_Click(object sender, RoutedEventArgs e)
        {
            OutputRichTextBox.Document.Blocks.Clear();
            Paragraph para = new Paragraph();

            // Split into words while keeping spaces
            string[] words = InputTextBox.Text.Split(new[] { ' ' }, StringSplitOptions.None);

            foreach (var word in words)
            {
                if (string.IsNullOrWhiteSpace(word))
                {
                    para.Inlines.Add(new Run(" "));
                    continue;
                }

                // Clean for dictionary lookup (remove punctuation)
                string clean = new string(word.Where(c => char.IsLetterOrDigit(c)).ToArray());
                string punct = new string(word.Where(c => char.IsPunctuation(c)).ToArray());

                if (fullToAbbrDict.ContainsKey(clean))
                {
                    var options = fullToAbbrDict[clean];

                    // Create Quillbot-style interactive link
                    Run linkText = new Run(options[0]);
                    System.Windows.Documents.Hyperlink link = new System.Windows.Documents.Hyperlink(linkText) { NavigateUri = new Uri("http://internal") };
                    link.TextDecorations = TextDecorations.Underline;

                    // Style it based on your UI resources or direct code
                    link.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 196, 142));

                    link.PreviewMouseLeftButtonDown += (s, ev) =>
                    {
                        ev.Handled = true;
                        ShowPopup(link, options);
                    };

                    para.Inlines.Add(link);
                    if (!string.IsNullOrEmpty(punct)) para.Inlines.Add(new Run(punct));
                }
                else
                {
                    para.Inlines.Add(new Run(word));
                }
                para.Inlines.Add(new Run(" "));
            }

            OutputRichTextBox.Document.Blocks.Add(para);
        }

        private void ShowPopup(System.Windows.Documents.Hyperlink targetLink, List<string> options)
        {
            PopupStackPanel.Children.Clear();

            foreach (var opt in options)
            {
                Button btn = new Button
                {
                    Content = opt,
                    Style = (System.Windows.Style)FindResource("PopupButtonStyle") // Refers to the style in your XAML
                };

                btn.Click += (s, e) =>
                {
                    ((Run)targetLink.Inlines.FirstInline).Text = opt;
                    AbbrPopup.IsOpen = false;
                };

                PopupStackPanel.Children.Add(btn);
            }

            AbbrPopup.Placement = System.Windows.Controls.Primitives.PlacementMode.MousePoint;
            AbbrPopup.IsOpen = true;
        }

        #endregion
    }
}