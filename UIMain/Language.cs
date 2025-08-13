using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace AESCConstruct25.Localization
{
    public static class Language
    {
        private static readonly DataTable _translations = new DataTable();
        private static string[] _columns;

        // ControlId -> CSV translation key
        private static readonly Dictionary<string, string> _ribbonMap = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            // Frame Generator group
            ["AESCConstruct25.Group"] = "Ribbon.Group.FrameGenerator",

            // Copy Edges (label hidden in XML, but we translate anyway for consistency)
            ["AESC.Construct.CopyEdges"] = "Ribbon.Button.CopyEdges",

            // Export / BOM
            ["AESCConstruct25.ExportSTEPBtn"] = "Ribbon.Button.ExportSTEP",
            ["AESCConstruct25.ExportBOMBtn"] = "Ribbon.Button.GenerateBOM",
            ["AESCConstruct25.UpdateBOM"] = "Ribbon.Button.UpdateBOM",
            ["AESCConstruct25.ExportExcelBtn"] = "Ribbon.Button.ExportExcel",

            // Plate
            ["AESCConstruct25.PlateGroup"] = "Ribbon.Group.Plate",
            ["AESCConstruct25.Plate"] = "Ribbon.Button.Plate",

            // Fastener
            ["AESCConstruct25.FastenerGroup"] = "Ribbon.Group.Fastener",
            ["AESCConstruct25.Fastener"] = "Ribbon.Button.Fastener",

            // Rib Cut-Out
            ["AESCConstruct25.RibCutOutGroup"] = "Ribbon.Group.RibCutOut",
            ["AESCConstruct25.RibCutOut"] = "Ribbon.Button.RibCutOut",

            // Engraving
            ["AESCConstruct25.Engraving"] = "Ribbon.Group.Engraving",
            ["AESCConstruct25.EngravingBtn"] = "Ribbon.Button.Engraving",

            // Custom Properties
            ["AESCConstruct25.CustomProperties"] = "Ribbon.Group.CustomProperties",
            ["AESCConstruct25.CustomPropertiesBtn"] = "Ribbon.Button.CustomProperties",

            // Settings
            ["AESCConstruct25.SettingsSidebarBtn"] = "Ribbon.Button.Settings",

            // UI
            ["AESCConstruct25.FloatBtn"] = "Ribbon.Button.Float", // if you add it later
            ["AESCConstruct25.DockBtn"] = "Ribbon.Button.Dock",
        };

        static Language()
        {
            try
            {
                LoadCSV();
            }
            catch (Exception)
            {
                // Logger.Log("[Language] CSV load failed: " + ex.Message);
            }
        }

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

                    try { _translations.Rows.Add(row); } catch { }
                    rowNum++;
                }
            }
        }

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

        public static string RibbonGetLabel(string controlId)
        {
            if (string.IsNullOrEmpty(controlId)) return string.Empty;

            if (_ribbonMap.TryGetValue(controlId, out var key))
                return Translate(key);

            // Fallbacks:
            // 1) Allow direct id-based translation (if you put the id in the CSV)
            var direct = Translate(controlId);
            if (!string.IsNullOrWhiteSpace(direct) && !string.Equals(direct, controlId, StringComparison.Ordinal))
                return direct;

            // 2) Last resort: show the id (useful while wiring new controls)
            return controlId;
        }

        /// <summary>
        /// For explicit cases (e.g., legacy GetEngravingLabel/GetCustomPropertiesLabel),
        /// translate by a known CSV key.
        /// </summary>
        public static string RibbonGetLabelByKey(string csvKey)
            => Translate(csvKey);

        /// <summary>
        /// Called once from OnRibbonLoad to enable invalidation.
        /// </summary>
        //public static void AttachRibbon(IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        /// <summary>
        /// Re-queries all ribbon labels. Call this after the user changes language.
        /// </summary>
        //public static void InvalidateRibbon()
        //{
        //    try { _ribbon?.Invalidate(); } catch { /* ignore UI exceptions */ }
        //}




        /// <summary>
        /// Translates only the header/text properties, leaving all control content intact.
        /// </summary>
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
                             !(cc is HeaderedContentControl) &&
                             !(cc is ItemsControl))
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
