/*
 CustomComponentControl lets users apply, update and remove document/component custom
 properties in SpaceClaim based on CSV templates, for both the active document and components.
*/

using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Properties;
using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Application = SpaceClaim.Api.V242.Application;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct2026.UI
{
    public partial class CustomComponentControl : UserControl
    {
        // Map display name -> full path so button handlers can open the file later
        private readonly Dictionary<string, string> _templateMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // Resolves a CSV path that may be relative (under ProgramData\AESCConstruct) or absolute.
        private string ResolveCsvPath(string relOrAbs)
        {
            if (string.IsNullOrWhiteSpace(relOrAbs))
                return null;

            return Path.IsPathRooted(relOrAbs)
                ? relOrAbs
                : Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "AESCConstruct",
                    relOrAbs
                );
        }

        // Initializes the custom component control UI, localization and loads template options.
        public CustomComponentControl()
        {
            InitializeComponent();

            DataContext = this;
            Localization.Language.LocalizeFrameworkElement(this);

            TemplateOptions = new ObservableCollection<string>();

            LoadTemplateOptions();
        }

        // Templates dropdown backing collection and selected item.
        public ObservableCollection<string> TemplateOptions { get; }
        public string SelectedTemplate { get; set; }

        // Returns the full file path for the currently selected template name, if any.
        private string GetSelectedTemplateFullPath()
        {
            if (string.IsNullOrWhiteSpace(SelectedTemplate)) return null;
            return _templateMap.TryGetValue(SelectedTemplate, out var full) ? full : null;
        }

        // Scans the configured folder for CSV templates and populates the TemplateOptions list.
        private void LoadTemplateOptions()
        {
            _templateMap.Clear();
            TemplateOptions.Clear();

            // Determine search folder:
            // 1) If Settings.Default.CompProperties points to a file, use its folder
            // 2) Otherwise default to %ProgramData%\AESCConstruct
            string configuredPath = ResolveCsvPath(Settings.Default.CompProperties);
            string searchDir = null;

            if (!string.IsNullOrWhiteSpace(configuredPath) && File.Exists(configuredPath))
                searchDir = Path.GetDirectoryName(configuredPath);

            if (string.IsNullOrWhiteSpace(searchDir))
                searchDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "AESCConstruct");

            if (!Directory.Exists(searchDir))
            {
                Application.ReportStatus($"Template folder not found: {searchDir}", StatusMessageType.Warning, null);
                SelectedTemplate = null;
                return;
            }

            // Collect all *.csv files (top level only), sorted by file name
            var csvFiles = Directory.EnumerateFiles(searchDir, "*.csv", SearchOption.TopDirectoryOnly)
                                    .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
                                    .ToList();

            foreach (var full in csvFiles)
            {
                var name = Path.GetFileName(full);
                if (!_templateMap.ContainsKey(name))
                {
                    _templateMap[name] = full;
                    TemplateOptions.Add(name);
                }
            }

            SelectedTemplate = TemplateOptions.FirstOrDefault();

            Application.ReportStatus(
                TemplateOptions.Count > 0
                    ? $"Loaded {TemplateOptions.Count} template(s) from {searchDir}"
                    : $"No .csv templates found in {searchDir}",
                StatusMessageType.Information, null);
        }

        // Handles “Add document properties” button: merges CSV-defined props into Document.CustomProperties.
        private void AddDocPropsButton_Click(object sender, RoutedEventArgs e)
        {
            string runId = Guid.NewGuid().ToString("N");
            Logger.Log($"[AddDocProps:{runId}] START");

            var window = Window.ActiveWindow;
            var doc = window?.Document;
            if (doc == null)
            {
                Logger.Log($"[AddDocProps:{runId}] No active document. ABORT");
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            var csvFullPath = GetSelectedTemplateFullPath();
            Logger.Log($"[AddDocProps:{runId}] SelectedTemplate='{SelectedTemplate}', FullPath='{csvFullPath}'");
            if (string.IsNullOrWhiteSpace(csvFullPath) || !File.Exists(csvFullPath))
            {
                Logger.Log($"[AddDocProps:{runId}] Template file not found. ABORT");
                Application.ReportStatus("Template file not found.", StatusMessageType.Warning, null);
                return;
            }

            // Read CSV and SKIP the first data row (header: Property,Value)
            List<(string Key, string DefaultValue)> csvProps;
            try
            {
                csvProps = ReadTemplateProperties(csvFullPath, runId, true);
                Logger.Log($"[AddDocProps:{runId}] CSV properties read (after header skip): {csvProps.Count}");
                for (int i = 0; i < csvProps.Count; i++)
                    Logger.Log($"[AddDocProps:{runId}] CSV[{i}] Key='{csvProps[i].Key}', Default='{csvProps[i].DefaultValue}'");
            }
            catch (Exception ex)
            {
                Logger.Log($"[AddDocProps:{runId}] ERROR reading CSV: {ex}");
                Application.ReportStatus("Failed to read template CSV.", StatusMessageType.Error, null);
                return;
            }

            if (csvProps.Count == 0)
            {
                Logger.Log($"[AddDocProps:{runId}] CSV empty (after header skip). ABORT");
                Application.ReportStatus("Template contains no properties.", StatusMessageType.Information, null);
                return;
            }

            // Snapshot existing DOC properties (outside write block is safe for reading)
            List<KeyValuePair<string, string>> existing;
            try
            {
                existing = GetExistingDocProps(doc, runId);
                Logger.Log($"[AddDocProps:{runId}] Existing DOC props: {existing.Count}");
                for (int i = 0; i < existing.Count; i++)
                    Logger.Log($"[AddDocProps:{runId}] EXIST[{i}] '{existing[i].Key}' = '{existing[i].Value}'");
            }
            catch (Exception ex)
            {
                Logger.Log($"[AddDocProps:{runId}] ERROR reading existing DOC properties: {ex}");
                Application.ReportStatus("Failed to read existing document properties.", StatusMessageType.Error, null);
                return;
            }

            var csvKeys = new HashSet<string>(csvProps.Select(p => p.Key), StringComparer.OrdinalIgnoreCase);
            var tailExisting = existing.Where(kv => !csvKeys.Contains(kv.Key)).ToList();
            Logger.Log($"[AddDocProps:{runId}] Tail (existing not in CSV): {tailExisting.Count}");
            for (int i = 0; i < tailExisting.Count; i++)
                Logger.Log($"[AddDocProps:{runId}] TAIL[{i}] '{tailExisting[i].Key}' = '{tailExisting[i].Value}'");

            // Mutations must be in a write block
            try
            {
                WriteBlock.ExecuteTask("Update Document Custom Properties", () =>
                {
                    Logger.Log($"[AddDocProps:{runId}] WRITEBLOCK BEGIN");
                    RebuildDocPropertiesInOrder(doc, csvProps, existing, tailExisting, runId);
                    DumpDocProps(doc, "AFTER REBUILD", runId);
                    Logger.Log($"[AddDocProps:{runId}] WRITEBLOCK END");
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"[AddDocProps:{runId}] ERROR during rebuild (WriteBlock): {ex}");
                Application.ReportStatus("Failed while updating document properties.", StatusMessageType.Error, null);
                return;
            }

            Application.ReportStatus($"Document properties updated from {SelectedTemplate}", StatusMessageType.Information, null);
            Logger.Log($"[AddDocProps:{runId}] DONE");
        }

        // Handles “Delete document properties” button: removes Document.CustomProperties that are listed in the CSV template.
        private void DeleteDocPropsButton_Click(object sender, RoutedEventArgs e)
        {
            string runId = Guid.NewGuid().ToString("N");
            Logger.Log($"[DelDocProps:{runId}] START");

            var window = Window.ActiveWindow;
            var doc = window?.Document;
            if (doc == null)
            {
                Logger.Log($"[DelDocProps:{runId}] No active document. ABORT");
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            var csvFullPath = GetSelectedTemplateFullPath();
            Logger.Log($"[DelDocProps:{runId}] SelectedTemplate='{SelectedTemplate}', FullPath='{csvFullPath}'");
            if (string.IsNullOrWhiteSpace(csvFullPath) || !File.Exists(csvFullPath))
            {
                Logger.Log($"[DelDocProps:{runId}] Template file not found. ABORT");
                Application.ReportStatus("Template file not found.", StatusMessageType.Warning, null);
                return;
            }

            List<(string Key, string DefaultValue)> csvProps;
            try
            {
                // skip header here too
                csvProps = ReadTemplateProperties(csvFullPath, runId, skipHeaderRow: true);
                Logger.Log($"[DelDocProps:{runId}] CSV properties read (after header skip): {csvProps.Count}");
            }
            catch (Exception ex)
            {
                Logger.Log($"[DelDocProps:{runId}] ERROR reading CSV: {ex}");
                Application.ReportStatus("Failed to read template CSV.", StatusMessageType.Error, null);
                return;
            }
            if (csvProps.Count == 0)
            {
                Logger.Log($"[DelDocProps:{runId}] CSV empty (after header skip). ABORT");
                Application.ReportStatus("Template contains no properties.", StatusMessageType.Information, null);
                return;
            }

            var toDelete = new HashSet<string>(csvProps.Select(p => p.Key), StringComparer.OrdinalIgnoreCase);
            Logger.Log($"[DelDocProps:{runId}] Keys to delete: {string.Join(", ", toDelete.ToArray())}");

            try
            {
                WriteBlock.ExecuteTask("Delete Document Custom Properties", () =>
                {
                    Logger.Log($"[DelDocProps:{runId}] WRITEBLOCK BEGIN");

                    var dict = doc.CustomProperties;
                    var existingKeys = dict.Select(kv => kv.Key).ToList();
                    Logger.Log($"[DelDocProps:{runId}] Existing DOC keys count: {existingKeys.Count}");

                    int deleted = 0;
                    foreach (var key in existingKeys)
                    {
                        if (!toDelete.Contains(key)) continue;

                        try
                        {
                            if (TryDeleteCustomProperty(dict, key, runId)) deleted++;
                        }
                        catch (Exception exDel)
                        {
                            Logger.Log($"[DelDocProps:{runId}] ERROR deleting '{key}': {exDel}");
                        }
                    }

                    Logger.Log($"[DelDocProps:{runId}] Deleted count: {deleted}. Dump after delete:");
                    DumpDocProps(doc, "AFTER DELETE", runId);

                    Logger.Log($"[DelDocProps:{runId}] WRITEBLOCK END");
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"[DelDocProps:{runId}] ERROR during delete (WriteBlock): {ex}");
                Application.ReportStatus("Failed while deleting document properties.", StatusMessageType.Error, null);
                return;
            }

            Application.ReportStatus($"Deleted properties from {SelectedTemplate}", StatusMessageType.Information, null);
            Logger.Log($"[DelDocProps:{runId}] DONE");
        }


        // Dumps all current document custom properties to the log, with type and value information.
        private static void DumpDocProps(Document doc, string label, string runId)
        {
            Logger.Log($"[DumpDOC:{runId}] ---- {label} ----");
            int i = 0;
            foreach (var kv in doc.CustomProperties)
            {
                var key = kv.Key ?? string.Empty;
                var raw = GetCustomPropertyValueObject(kv.Value);
                var val = FormatValueAsString(raw);
                var typeName = raw == null ? "<null>" : raw.GetType().FullName;
                Logger.Log($"[DumpDOC:{runId}] [{i}] '{key}' (type {typeName}) = '{val}'");
                i++;
            }
            Logger.Log($"[DumpDOC:{runId}] ---- total {i} ----");
        }

        // Reads a template CSV into ordered (Key, DefaultValue) tuples, with flexible separators and comments.
        private static List<(string Key, string DefaultValue)> ReadTemplateProperties(string csvFullPath, string runId, bool skipHeaderRow)
        {
            Logger.Log($"[CSV:{runId}] Reading '{csvFullPath}'...");
            var result = new List<(string Key, string DefaultValue)>();
            string[] lines = File.ReadAllLines(csvFullPath);

            Logger.Log($"[CSV:{runId}] Lines read: {lines.Length}");
            bool headerSkipped = !skipHeaderRow; // if false, we still need to skip the first data row

            for (int lineIdx = 0; lineIdx < lines.Length; lineIdx++)
            {
                var raw = lines[lineIdx];
                var line = raw == null ? string.Empty : raw.Trim();
                if (string.IsNullOrEmpty(line)) { Logger.Log($"[CSV:{runId}] line {lineIdx}: empty -> skip"); continue; }
                if (line.StartsWith("#") || line.StartsWith("//")) { Logger.Log($"[CSV:{runId}] line {lineIdx}: comment -> skip"); continue; }

                var fields = SplitCsvLineFlexible(line);
                if (fields.Count == 0) { Logger.Log($"[CSV:{runId}] line {lineIdx}: no fields -> skip"); continue; }

                if (!headerSkipped)
                {
                    Logger.Log($"[CSV:{runId}] line {lineIdx}: HEADER skipped ('{string.Join("|", fields)}')");
                    headerSkipped = true;
                    continue;
                }

                var key = TrimQuotes(fields[0]).Trim();
                if (string.IsNullOrEmpty(key)) { Logger.Log($"[CSV:{runId}] line {lineIdx}: empty key -> skip"); continue; }

                string def = fields.Count > 1 ? TrimQuotes(fields[1]).Trim() : string.Empty;

                if (result.Any(t => t.Key.Equals(key, StringComparison.OrdinalIgnoreCase)))
                {
                    Logger.Log($"[CSV:{runId}] line {lineIdx}: duplicate key '{key}' -> keep first, skip");
                    continue;
                }

                result.Add((key, def));
                Logger.Log($"[CSV:{runId}] line {lineIdx}: add Key='{key}', Default='{def}'");
            }

            Logger.Log($"[CSV:{runId}] Parsed keys: {result.Count}");
            return result;
        }

        // Rebuilds Document.CustomProperties in CSV-driven order, preserving existing values and trailing keys.
        private static void RebuildDocPropertiesInOrder(
            Document doc,
            List<(string Key, string DefaultValue)> csvProps,
            List<KeyValuePair<string, string>> existing,
            List<KeyValuePair<string, string>> tailExisting,
            string runId)
        {
            Logger.Log($"[RebuildDOC:{runId}] BEGIN. CSV={csvProps.Count}, existing={existing.Count}, tail={tailExisting.Count}");

            var dict = doc.CustomProperties;

            // Remove ALL current DOC properties (we’ll re-add in desired order)
            var allExistingKeys = dict.Select(kv => kv.Key).ToList();
            Logger.Log($"[RebuildDOC:{runId}] Removing all current DOC properties ({allExistingKeys.Count})...");
            foreach (var key in allExistingKeys)
            {
                try
                {
                    if (!TryDeleteCustomProperty(dict, key, runId))
                        Logger.Log($"[RebuildDOC:{runId}] WARN: failed to delete '{key}'");
                }
                catch (Exception ex)
                {
                    Logger.Log($"[RebuildDOC:{runId}] ERROR deleting '{key}': {ex}");
                }
            }

            // CSV-ordered keys first
            Logger.Log($"[RebuildDOC:{runId}] Creating CSV-ordered DOC keys...");
            foreach (var entry in csvProps)
            {
                string key = entry.Key;
                string def = entry.DefaultValue;

                var match = existing.FirstOrDefault(kv => kv.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
                var valueString = match.Key != null ? match.Value : def;

                Logger.Log($"[RebuildDOC:{runId}] -> Create '{key}' (from {(match.Key != null ? "existing" : "CSV default")}) = '{valueString}'");
                CreateDocumentCustomProperty(doc, key, valueString, runId);
            }

            // Append the rest (not in CSV)
            Logger.Log($"[RebuildDOC:{runId}] Appending tail DOC keys...");
            foreach (var kv in tailExisting)
            {
                Logger.Log($"[RebuildDOC:{runId}] -> Append '{kv.Key}' = '{kv.Value}'");
                CreateDocumentCustomProperty(doc, kv.Key, kv.Value, runId);
            }

            Logger.Log($"[RebuildDOC:{runId}] END");
        }

        // Tries to delete a document custom property using either static APIs or instance Delete() via reflection.
        private static bool TryDeleteCustomProperty(CustomPropertyDictionary dict, string key, string runId)
        {
            try
            {
                // 1) Preferred: static Delete(Document, string)
                var cpType = typeof(Document).Assembly.GetType("SpaceClaim.Api.V242.CustomProperty")
                           ?? typeof(Document).Assembly.GetType("SpaceClaim.Api.V242.CustomDocumentProperty");
                if (cpType != null)
                {
                    var delStatic = cpType.GetMethod(
                        "Delete",
                        BindingFlags.Public | BindingFlags.Static,
                        null,
                        new[] { typeof(Document), typeof(string) },
                        null);

                    if (delStatic != null)
                    {
                        // Need the owning Document to use the static delete
                        // dict is doc.CustomProperties, so get the Document from any property on the dictionary
                        // There isn't a public backref, so instead use Value.Delete() below if we can't get Document.
                        // Try to locate the Document via reflection on the dictionary if available.
                        var docProp = dict.GetType().GetProperty("Document", BindingFlags.Public | BindingFlags.Instance);
                        var docObj = docProp != null ? docProp.GetValue(dict) as Document : null;
                        if (docObj != null)
                        {
                            Logger.Log($"[DelDOC:{runId}] CustomProperty.Delete(Document,'{key}')");
                            delStatic.Invoke(null, new object[] { docObj, key });
                            return true;
                        }
                    }
                }

                // 2) Fallback: locate entry and call instance Delete() on its value
                var kv = dict.FirstOrDefault(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
                if (kv.Key != null && kv.Value != null)
                {
                    var delInstance = kv.Value.GetType().GetMethod("Delete", BindingFlags.Public | BindingFlags.Instance);
                    if (delInstance != null)
                    {
                        Logger.Log($"[DelDOC:{runId}] Value.Delete() for '{key}'");
                        delInstance.Invoke(kv.Value, null);
                        return true;
                    }
                }

                Logger.Log($"[DelDOC:{runId}] No supported delete API found for '{key}'");
                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"[DelDOC:{runId}] ERROR deleting '{key}': {ex}");
                return false;
            }
        }

        // Creates a document-level custom property via the SpaceClaim API or via dictionary fallbacks.
        private static void CreateDocumentCustomProperty(Document doc, string key, string value, string runId)
        {
            try
            {
                // Preferred path: static Create on CustomProperty
                var apiAsm = typeof(Document).Assembly;
                var cpType = apiAsm.GetType("SpaceClaim.Api.V242.CustomProperty")
                          ?? apiAsm.GetType("SpaceClaim.Api.V242.CustomDocumentProperty"); // alternate name, just in case

                if (cpType != null)
                {
                    var create = cpType.GetMethod("Create", BindingFlags.Public | BindingFlags.Static, null,
                        new[] { typeof(Document), typeof(string), typeof(string) }, null);

                    if (create != null)
                    {
                        Logger.Log($"[CreateDOC:{runId}] Using CustomProperty.Create(Document, name, value) for '{key}'");
                        create.Invoke(null, new object[] { doc, key, value ?? string.Empty });
                        return;
                    }
                }

                // Fallback: try the dictionary
                var dict = doc.CustomProperties;
                var add = dict.GetType().GetMethod("Add", BindingFlags.Public | BindingFlags.Instance, null,
                    new[] { typeof(string), typeof(string) }, null);

                if (add != null)
                {
                    Logger.Log($"[CreateDOC:{runId}] Using CustomPropertyDictionary.Add('{key}','{value}')");
                    add.Invoke(dict, new object[] { key, value ?? string.Empty });
                    return;
                }

                // Last resort: indexer set via reflection (if available)
                var idxProp = dict.GetType().GetProperty("Item", new[] { typeof(string) });
                if (idxProp != null && idxProp.CanWrite)
                {
                    Logger.Log($"[CreateDOC:{runId}] Using CustomPropertyDictionary indexer for '{key}'");
                    idxProp.SetValue(doc.CustomProperties, value ?? string.Empty, new object[] { key });
                    return;
                }

                Logger.Log($"[CreateDOC:{runId}] ERROR: No known method to create DOC property '{key}'");
                throw new InvalidOperationException("No API to create document property.");
            }
            catch (Exception ex)
            {
                Logger.Log($"[CreateDOC:{runId}] ERROR creating DOC property '{key}': {ex}");
                throw;
            }
        }

        // Enumerates existing document custom properties and returns them as key/value strings.
        private static List<KeyValuePair<string, string>> GetExistingDocProps(Document doc, string runId)
        {
            Logger.Log($"[ExistDOC:{runId}] Enumerating Document.CustomProperties...");
            var list = new List<KeyValuePair<string, string>>();
            int i = 0;
            foreach (var kv in doc.CustomProperties)
            {
                var key = kv.Key ?? string.Empty;
                var raw = GetCustomPropertyValueObject(kv.Value);
                var val = FormatValueAsString(raw);

                string typeName = raw == null ? "<null>" : raw.GetType().FullName;
                Logger.Log($"[ExistDOC:{runId}] [{i}] '{key}' (type {typeName}) = '{val}'");
                list.Add(new KeyValuePair<string, string>(key, val));
                i++;
            }
            Logger.Log($"[ExistDOC:{runId}] Total enumerated: {list.Count}");
            return list;
        }

        // Uses reflection to extract the underlying Value object from a custom property entry.
        private static object GetCustomPropertyValueObject(object customPropEntry)
        {
            if (customPropEntry == null) return null;
            var p = customPropEntry.GetType().GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
            return p != null ? p.GetValue(customPropEntry) : null;
        }

        // Returns existing part-level custom properties for a Part as key/value strings.
        private static List<KeyValuePair<string, string>> GetExistingPartProps(Part part)
        {
            var list = new List<KeyValuePair<string, string>>();
            foreach (var kv in part.CustomProperties)
            {
                var key = kv.Key ?? string.Empty;
                var valObj = kv.Value?.Value;
                var val = valObj == null ? string.Empty : Convert.ToString(valObj, CultureInfo.InvariantCulture);
                list.Add(new KeyValuePair<string, string>(key, val));
            }
            return list;
        }

        // Normalizes an arbitrary value into a string (bools lowercased, numeric via invariant culture).
        private static string FormatValueAsString(object value)
        {
            if (value == null) return string.Empty;
            var b = value as bool?;
            if (b.HasValue) return b.Value.ToString().ToLowerInvariant();

            var formattable = value as IFormattable;
            if (formattable != null) return formattable.ToString(null, CultureInfo.InvariantCulture);

            return value.ToString();
        }

        // Creates or updates a part custom property as a string; rebuild logic controls overwrite semantics.
        private static void CreateCustomProperty(Part part, string key, object newValue)
        {
            string valueString;
            switch (newValue)
            {
                case bool b:
                    valueString = b.ToString().ToLowerInvariant();
                    break;
                case IFormattable formattable:
                    valueString = formattable.ToString(null, CultureInfo.InvariantCulture);
                    break;
                default:
                    valueString = newValue?.ToString() ?? string.Empty;
                    break;
            }

            var existingKV = part.CustomProperties
                .FirstOrDefault(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

            if (existingKV.Key != null)
                existingKV.Value.Value = valueString;
            else
                CustomPartProperty.Create(part, key, valueString);
        }

        // Deletes a part-level custom property if it exists (case-insensitively) on the Part.
        private static void DeletePartProperty(Part part, string key)
        {
            if (part == null || string.IsNullOrWhiteSpace(key)) return;

            // Find the entry (case-insensitive) in the Part’s custom properties
            var kv = part.CustomProperties
                         .FirstOrDefault(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

            // If found, delete the *instance*
            if (kv.Key != null && kv.Value != null)
                kv.Value.Delete(); // instance method per V242 API
        }

        // Rebuilds a Part's custom properties from CSV order while preserving non-CSV properties at the end.
        private static void RebuildPartPropertiesInOrder(
            Part part,
            List<(string Key, string DefaultValue)> csvProps)
        {
            var existing = GetExistingPartProps(part);
            var csvKeySet = new HashSet<string>(csvProps.Select(p => p.Key), StringComparer.OrdinalIgnoreCase);
            var tailExisting = existing.Where(kv => !csvKeySet.Contains(kv.Key)).ToList();

            // Remove all
            var allKeys = part.CustomProperties.Select(kv => kv.Key).ToList();
            foreach (var k in allKeys)
                DeletePartProperty(part, k);

            // Re-create CSV keys first, keeping existing values when they existed
            foreach (var (key, def) in csvProps)
            {
                var match = existing.FirstOrDefault(kv => kv.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
                var valueToUse = match.Key != null ? match.Value : def;
                CreateCustomProperty(part, key, valueToUse);
            }

            // Then append the rest (not in CSV), preserving their relative order/values
            foreach (var kv in tailExisting)
                CreateCustomProperty(part, kv.Key, kv.Value);
        }

        // Splits a single CSV line into fields, handling quotes and comma/semicolon separators.
        private static List<string> SplitCsvLineFlexible(string line)
        {
            var result = new List<string>();
            var cur = new StringBuilder();
            bool inQuotes = false;

            for (int idx = 0; idx < line.Length; idx++)
            {
                char ch = line[idx];
                if (ch == '\"') { inQuotes = !inQuotes; continue; }

                if (!inQuotes && (ch == ',' || ch == ';'))
                {
                    result.Add(cur.ToString());
                    cur.Clear();
                }
                else cur.Append(ch);
            }
            result.Add(cur.ToString());
            return result;
        }

        // Trims wrapping double quotes from a string if both first and last characters are quotes.
        private static string TrimQuotes(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            s = s.Trim();
            if (s.Length >= 2 && s[0] == '\"' && s[s.Length - 1] == '\"')
                return s.Substring(1, s.Length - 2);
            return s;
        }

        // Applies template properties from the selected CSV to all currently selected Components.
        private void AddToSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.ActiveWindow;
            var doc = window?.Document;
            if (doc == null)
            {
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            var csvFullPath = GetSelectedTemplateFullPath();
            if (string.IsNullOrWhiteSpace(csvFullPath) || !File.Exists(csvFullPath))
            {
                Application.ReportStatus("Template file not found.", StatusMessageType.Warning, null);
                return;
            }

            var comps = GetSelectedComponents(window);
            if (comps.Count == 0)
            {
                Application.ReportStatus("Select one or more components.", StatusMessageType.Warning, null);
                return;
            }

            var csvProps = ReadTemplateProperties(csvFullPath, Guid.NewGuid().ToString("N"), skipHeaderRow: true);
            if (csvProps.Count == 0)
            {
                Application.ReportStatus("Template contains no properties.", StatusMessageType.Information, null);
                return;
            }

            WriteBlock.ExecuteTask("Add Component Properties (Selection)", () =>
            {
                foreach (var comp in comps)
                    RebuildPartPropertiesInOrder(comp.Template, csvProps);
            });

            Application.ReportStatus("Component properties added to selection.", StatusMessageType.Information, null);
        }

        // Deletes template-defined properties from all currently selected Components based on the CSV keys.
        private void DeleteFromSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.ActiveWindow;
            var doc = window?.Document;
            if (doc == null)
            {
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            var csvFullPath = GetSelectedTemplateFullPath();
            if (string.IsNullOrWhiteSpace(csvFullPath) || !File.Exists(csvFullPath))
            {
                Application.ReportStatus("Template file not found.", StatusMessageType.Warning, null);
                return;
            }

            var comps = GetSelectedComponents(window);
            if (comps.Count == 0)
            {
                Application.ReportStatus("Select one or more components.", StatusMessageType.Warning, null);
                return;
            }

            var csvProps = ReadTemplateProperties(csvFullPath, Guid.NewGuid().ToString("N"), skipHeaderRow: true);
            if (csvProps.Count == 0)
            {
                Application.ReportStatus("Template contains no properties.", StatusMessageType.Information, null);
                return;
            }
            var keys = csvProps.Select(p => p.Key).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            WriteBlock.ExecuteTask("Delete Component Properties (Selection)", () =>
            {
                foreach (var comp in comps)
                    foreach (var k in keys)
                        DeletePartProperty(comp.Template, k);
            });

            Application.ReportStatus("Component properties deleted from selection.", StatusMessageType.Information, null);
        }

        // Applies template-defined properties to all components in the main part of the active document.
        private void AddToAllButton_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.ActiveWindow;
            var doc = window?.Document;
            var mainPart = doc?.MainPart;
            if (mainPart == null)
            {
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            var csvFullPath = GetSelectedTemplateFullPath();
            if (string.IsNullOrWhiteSpace(csvFullPath) || !File.Exists(csvFullPath))
            {
                Application.ReportStatus("Template file not found.", StatusMessageType.Warning, null);
                return;
            }

            var comps = mainPart.GetDescendants<IComponent>().OfType<Component>().ToList();
            if (comps.Count == 0)
            {
                Application.ReportStatus("No components found.", StatusMessageType.Information, null);
                return;
            }

            var csvProps = ReadTemplateProperties(csvFullPath, Guid.NewGuid().ToString("N"), skipHeaderRow: true);
            if (csvProps.Count == 0)
            {
                Application.ReportStatus("Template contains no properties.", StatusMessageType.Information, null);
                return;
            }

            WriteBlock.ExecuteTask("Add Component Properties (All)", () =>
            {
                foreach (var comp in comps)
                    RebuildPartPropertiesInOrder(comp.Template, csvProps);
            });

            Application.ReportStatus("Component properties added to all components.", StatusMessageType.Information, null);
        }

        // Deletes template-defined properties from all components in the main part of the active document.
        private void DeleteFromAllButton_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.ActiveWindow;
            var doc = window?.Document;
            var mainPart = doc?.MainPart;
            if (mainPart == null)
            {
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            var csvFullPath = GetSelectedTemplateFullPath();
            if (string.IsNullOrWhiteSpace(csvFullPath) || !File.Exists(csvFullPath))
            {
                Application.ReportStatus("Template file not found.", StatusMessageType.Warning, null);
                return;
            }

            var comps = mainPart.GetDescendants<IComponent>().OfType<Component>().ToList();
            if (comps.Count == 0)
            {
                Application.ReportStatus("No components found.", StatusMessageType.Information, null);
                return;
            }

            var csvProps = ReadTemplateProperties(csvFullPath, Guid.NewGuid().ToString("N"), skipHeaderRow: true);
            if (csvProps.Count == 0)
            {
                Application.ReportStatus("Template contains no properties.", StatusMessageType.Information, null);
                return;
            }
            var keys = csvProps.Select(p => p.Key).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            WriteBlock.ExecuteTask("Delete Component Properties (All)", () =>
            {
                foreach (var comp in comps)
                    foreach (var k in keys)
                        DeletePartProperty(comp.Template, k);
            });

            Application.ReportStatus("Deleted component properties from all components.", StatusMessageType.Information, null);
        }

        // Returns only Component instances from the current SpaceClaim selection, logging them for diagnostics.
        private static List<Component> GetSelectedComponents(Window window)
        {
            var sel = window?.ActiveContext?.Selection;
            if (sel == null || sel.Count == 0) return new List<Component>();

            // Only return items that are actually Components in the selection.
            var comps = sel.OfType<Component>().Distinct().ToList();

            Logger.Log($"[GetSelectedComponents] Selected components: {comps.Count}");
            for (int i = 0; i < comps.Count; i++)
                Logger.Log($"[GetSelectedComponents] [{i}] Name='{comps[i].Name}'");

            return comps;
        }
    }
}
