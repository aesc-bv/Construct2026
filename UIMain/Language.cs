/*
 Language provides CSV-based runtime localization support.
 It loads translations from languageConstruct.csv and applies them to WPF controls based on Tag identifiers.
*/

using AESCConstruct2026.FrameGenerator.Utilities;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace AESCConstruct2026.Localization
{
    public static class Language
    {
        private static readonly DataTable _translations = new DataTable();
        private static string[] _columns;

        // Loads the translation CSV into an in-memory DataTable keyed by translation id.
        public static void LoadCSV()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            var path = Path.Combine(appData, "AESCConstruct", "Language", "languageConstruct.csv");

            // Logger.Log("[Language] Loading CSV from " + path);
            if (!File.Exists(path))
                throw new FileNotFoundException("[Language] Missing CSV at " + path);

            using (var sr = new StreamReader(path, Encoding.GetEncoding("Windows-1252")))
            {
                var header = sr.ReadLine();
                if (header == null)
                    throw new InvalidDataException("[Language] Empty CSV");

                _columns = header.Split(';').Select(c => c.Trim('\uFEFF', ' ', '"')).ToArray();
                foreach (var col in _columns) _translations.Columns.Add(col);
                _translations.PrimaryKey = new[] { _translations.Columns[_columns[0]] };

                int rowNum = 2;
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine() ?? string.Empty;
                    var cells = Regex.Split(
                       line,
                       // this pattern correctly splits on ; outside of quotes
                       @";(?=(?:[^""]*""[^""]*"")*[^""]*$)"
                    );
                    var row = _translations.NewRow();
                    for (int i = 0; i < _columns.Length; i++)
                        row[i] = i < cells.Length ? cells[i] : string.Empty;

                    try { _translations.Rows.Add(row); } catch (Exception ex) { Logger.Log($"[Language] Duplicate row at line {rowNum}: {ex.Message}"); }
                    rowNum++;
                }
            }
        }

        // Looks up a localized string for the given id using the configured language, with English and id fallback.
        public static string Translate(string id)
        {
            if (string.IsNullOrEmpty(id)) return string.Empty;
            if (_columns == null) LoadCSV();

            var lang = Properties.Settings.Default.Construct_Language;
            var row = _translations.Rows.Find(id);
            if (row != null)
            {
                if (_translations.Columns.Contains(lang))
                {
                    var txt = row[lang]?.ToString();
                    if (!string.IsNullOrWhiteSpace(txt)) return txt;
                }
                if (_translations.Columns.Contains("EN"))
                {
                    var txtEn = row["EN"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(txtEn)) return txtEn;
                }
            }
            return id;
        }

        // Recursively walks a WPF element tree and replaces content for tagged elements with translated text.
        public static void LocalizeFrameworkElement(FrameworkElement root)
        {
            if (root == null) return;

            foreach (var child in LogicalTreeHelper.GetChildren(root).OfType<FrameworkElement>())
            {
                if (child.Tag is string key && _translations.Rows.Find(key) is DataRow)
                {
                    var translation = Translate(key);

                    // TextBlock: update Text
                    if (child is TextBlock tb)
                    {
                        tb.Text = translation;
                    }
                    // HeaderedContentControl: update Header only (Expander, GroupBox, TabItem)
                    else if (child is HeaderedContentControl hcc)
                    {
                        hcc.Header = translation;
                    }
                    // ContentControl (Button, CheckBox, Label, but not Expander or ItemsControl)
                    else if (child is ContentControl cc &&
                             !(cc is HeaderedContentControl))
                    {
                        cc.Content = translation;
                    }
                }

                // Recurse
                LocalizeFrameworkElement(child);
            }
        }
    }
}
