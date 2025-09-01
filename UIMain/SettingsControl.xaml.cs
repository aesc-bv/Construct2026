using AESCConstruct25.Licensing;
using AESCConstruct25.Properties;
using Microsoft.Win32;
using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Application = SpaceClaim.Api.V242.Application;


namespace AESCConstruct25.FrameGenerator.UI
{
    public partial class SettingsControl : UserControl
    {
        public ObservableCollection<string> UnitOptions { get; }
            = new ObservableCollection<string>(new[] { "mm", "cm", "m", "inch" });
        public ObservableCollection<string> AnchorOptions { get; }
            = new ObservableCollection<string>(new[] { "TopLeft", "TopRight", "BottomLeft", "BottomRight" });

        // model for each row
        public class TypeNamePair
        {
            public string Type { get; set; }
            public string Name { get; set; }
        }

        // backing collection
        private readonly ObservableCollection<TypeNamePair> _pairs = new ObservableCollection<TypeNamePair>();

        public ObservableCollection<string> CornerOptions { get; }
          = new ObservableCollection<string>(
              Enum.GetNames(typeof(LocationPoint))
          );

        public SettingsControl()
        {
            InitializeComponent();
            this.Loaded += SettingsControl_Loaded;
            DataContext = this;
        }
        private void SettingsControl_Loaded(object sender, RoutedEventArgs e)
        {
            // 1) populate language combo with codes
            ComboBox_Language.ItemsSource = new[] { "EN", "NL", "FR", "DE", "IT", "ES" };
            ComboBox_Language.SelectedItem = Settings.Default.Construct_Language;

            // Length unit
            Label_BOMLengthUnitComboBox.ItemsSource = new[] { "mm", "cm", "m", "inch" };
            Label_BOMLengthUnitComboBox.SelectedItem = Settings.Default.LengthUnit;

            // Drawing sheet anchor (document anchor)
            Label_BOMAnchorComboBox.ItemsSource = AnchorOptions;
            Label_BOMAnchorComboBox.SelectedItem = Settings.Default.DocumentAnchor;

            // Table corner (LocationPoint)
            CornerComboBox.ItemsSource = new[] { "TopLeftCorner", "TopRightCorner", "BottomLeftCorner", "BottomRightCorner" };
            CornerComboBox.SelectedItem = Settings.Default.TableLocationPoint;

            // 2) translate every tagged element in _this_ control
            Localization.Language.LocalizeFrameworkElement(this);

            UIMain.UIManager.UpdateCommandTexts();
            Construct25.UpdateCommandTexts();

            PopulateCsvFilePathFields();

            //Label_BOMLengthUnitComboBox.ItemsSource = UnitOptions;
            //var unit = Settings.Default.LengthUnit;
            //if (!UnitOptions.Contains(unit)) unit = "mm";
            //Label_BOMLengthUnitComboBox.SelectedItem = unit;

            //Label_BOMAnchorComboBox.ItemsSource = AnchorOptions;
            //var anchor = Settings.Default.DocumentAnchor;
            //if (!AnchorOptions.Contains(anchor)) anchor = "BottomRight";
            //Label_BOMAnchorComboBox.SelectedItem = anchor;
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            // 1) Load the Part naming textbox
            PartNameTextBox.Text = Settings.Default.NameString ?? "";

            // 2) Load and parse the TypeString into the ObservableCollection
            _pairs.Clear();
            var raw = Settings.Default.TypeString ?? "";
            var entries = raw.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

            Checkbox_ExcelMaterial.IsChecked = Settings.Default.MatInExcel;
            Checkbox_BOMMaterial.IsChecked = Settings.Default.MatInBOM;
            Checkbox_STEPMaterial.IsChecked = Settings.Default.MatInSTEP;

            Label_BOMLengthUnitComboBox.ItemsSource = UnitOptions;
            var unit = Settings.Default.LengthUnit;
            if (!UnitOptions.Contains(unit)) unit = "mm";
            Label_BOMLengthUnitComboBox.SelectedItem = unit;


            foreach (var entry in entries)
            {
                var parts = entry.Split(new[] { '@' }, 2);
                if (parts.Length == 2)
                {
                    _pairs.Add(new TypeNamePair { Type = parts[0], Name = parts[1] });
                }
                else
                {
                    _pairs.Add(new TypeNamePair { Type = entry, Name = "" });
                }
            }

            // bind it
            TypesDataGrid.ItemsSource = _pairs;

            // anchor coordinates
            AnchorXTextBox.Text = Settings.Default.TableAnchorX.ToString();
            AnchorYTextBox.Text = Settings.Default.TableAnchorY.ToString();

            // corner enum
            //var corner = Settings.Default.TableLocationPoint;
            //if (CornerComboBox.Items.Contains(corner))
            //    CornerComboBox.SelectedItem = corner;
            //else
            //    CornerComboBox.SelectedIndex = 0;

            // License block
            SerialNumberTextBox.Text = Settings.Default.SerialNumber ?? "";
            LicenseStatusValue.Text = ConstructLicenseSpot.Status;
            var expiration = ConstructLicenseSpot.CurrentLicense?.GetTimeLimit()?.EndDate;
            LicenseDetailsValue.Text = expiration.HasValue ? expiration.Value.ToShortDateString() : string.Empty;

            ComboBox_Language.ItemsSource = new[] { "EN", "NL", "FR", "DE", "IT", "ES" };
            ComboBox_Language.SelectedItem = Settings.Default.Construct_Language ?? "EN";
            // Length unit
            Label_BOMLengthUnitComboBox.ItemsSource = new[] { "mm", "cm", "m", "inch" };
            Label_BOMLengthUnitComboBox.SelectedItem = Settings.Default.LengthUnit;

            // Drawing sheet anchor (document anchor)
            Label_BOMAnchorComboBox.ItemsSource = new[] { "TopLeft", "TopRight", "BottomLeft", "BottomRight" };
            Label_BOMAnchorComboBox.SelectedItem = Settings.Default.DocumentAnchor;

            // Table corner (LocationPoint)
            CornerComboBox.ItemsSource = AnchorOptions;
            CornerComboBox.SelectedItem = Settings.Default.TableLocationPoint;


            // Localize all named elements
            Localization.Language.LocalizeFrameworkElement(this);

            UIMain.UIManager.UpdateCommandTexts();
            Construct25.UpdateCommandTexts();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Save Part naming
            Settings.Default.NameString = PartNameTextBox.Text.Trim();

            // Reserialize the list back into TypeString
            var serialized = string.Join("|",
                _pairs
                    .Where(p => !string.IsNullOrWhiteSpace(p.Type) || !string.IsNullOrWhiteSpace(p.Name))
                    .Select(p => $"{p.Type}@{p.Name}")
            );
            Settings.Default.TypeString = serialized;

            Settings.Default.MatInExcel = Checkbox_ExcelMaterial.IsChecked == true;
            Settings.Default.MatInBOM = Checkbox_BOMMaterial.IsChecked == true;
            Settings.Default.MatInSTEP = Checkbox_STEPMaterial.IsChecked == true;

            // Refresh license state from disk (no modal)
            ConstructLicenseSpot.CheckLicense();
            LicenseStatusValue.Text = ConstructLicenseSpot.Status;
            var expiration = ConstructLicenseSpot.CurrentLicense?.GetTimeLimit()?.EndDate;
            LicenseDetailsValue.Text = expiration.HasValue ? expiration.Value.ToShortDateString() : string.Empty;

            Construct25.RefreshLicenseUI();

            // --- Persist BOM placement settings ---
            if (double.TryParse(AnchorXTextBox.Text, out var ax))
                Settings.Default.TableAnchorX = ax;

            if (double.TryParse(AnchorYTextBox.Text, out var ay))
                Settings.Default.TableAnchorY = ay;

            if (CornerComboBox.SelectedItem is string cp)
                Settings.Default.TableLocationPoint = cp;

            if (SerialNumberTextBox.Text is string sn)
                Settings.Default.SerialNumber = sn;

            foreach (var row in CsvPathsPanel.Children.OfType<StackPanel>())
            {
                // find the TextBox inside that row:
                var txt = row.Children
                             .OfType<TextBox>()
                             .FirstOrDefault();
                if (txt == null || !(txt.Tag is string settingKey))
                    continue;

                var path = txt.Text.Trim();
                if (!string.IsNullOrWhiteSpace(path))
                {
                    // Logger.Log($"Saving setting [{settingKey}] = {path}");
                    Settings.Default[settingKey] = path;
                }
            }

            if (Label_BOMLengthUnitComboBox.SelectedItem is string selUnit && !string.IsNullOrWhiteSpace(selUnit))
                Settings.Default.LengthUnit = selUnit;

            if (Label_BOMAnchorComboBox.SelectedItem is string selAnchor && !string.IsNullOrWhiteSpace(selAnchor))
                Settings.Default.DocumentAnchor = selAnchor;

            Settings.Default.Save();
            // persist all settings
            try
            {
                Settings.Default.Save();
                //MessageBox.Show("Settings saved.", "Settings", MessageBoxButton.OK, MessageBoxImage.Information);
                Application.ReportStatus("Settings saved.", StatusMessageType.Information, null);
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"Error saving settings:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.ReportStatus($"Error saving settings:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        private void ActivateLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            var serial = SerialNumberTextBox.Text.Trim();
            if (string.IsNullOrEmpty(serial))
            {
                //MessageBox.Show("Please enter a serial number first.", "Activate", MessageBoxButton.OK, MessageBoxImage.Warning);
                Application.ReportStatus("Please enter a serial number first.", StatusMessageType.Warning, null);
                return;
            }

            bool ok = ConstructLicenseSpot.TryActivate(serial, out var msg);
            LicenseStatusValue.Text = msg;

            var expiration = ConstructLicenseSpot.CurrentLicense?.GetTimeLimit()?.EndDate;
            LicenseDetailsValue.Text = expiration.HasValue ? expiration.Value.ToShortDateString() : string.Empty;

            AESCConstruct25.Construct25.RefreshLicenseUI();

            if (ok)
                Application.ReportStatus("License activated.", StatusMessageType.Information, null);
            //Application.ReportStatus("Please enter a serial number first.", StatusMessageType.Warning, null);
        }

        private void ComboBox_Language_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComboBox_Language.SelectedItem is string langCode)
            {
                Settings.Default.Construct_Language = langCode;
                Settings.Default.Save();

                // re-translate this panel
                Localization.Language.LocalizeFrameworkElement(this);

                //AESCConstruct25.Localization.Language.InvalidateRibbon();                // ribbon text
                UIMain.UIManager.UpdateCommandTexts();                   // sidebar/floating commands
                Construct25.UpdateCommandTexts();
            }
        }

        private void PopulateCsvFilePathFields()
        {
            // Logger.Log("=== PopulateCsvFilePathFields (Using Settings Keys) ===");

            CsvPathsPanel.Children.Clear();

            //string baseDir = Path.Combine(
            //    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            //    "AESCConstruct"
            //);
            // Logger.Log($"Base directory: {baseDir}");

            // Mapping: Label → Setting Name
            var settingMap = new Dictionary<string, string>
            {
                { "Connector properties", "ConnectorProperties" },
                { "Component properties", "CompProperties" },
                { "Bolts", "Bolt" },
                { "Nuts", "Nut" },
                { "Washers", "Washer" },
                { "Plate properties", "PlatesProperties" },
                { "Profiles - circular", "Profiles_Circular" },
                { "Profiles - H", "Profiles_H" },
                { "Profiles - L", "Profiles_L" },
                { "Profiles - rectangular", "Profiles_Rectangular" },
                { "Profiles - T", "Profiles_T" },
                { "Profiles - U", "Profiles_U" },
                { "Custom profiles", "profiles" }
            };

            foreach (var kvp in settingMap)
            {
                string label = kvp.Key;
                string settingKey = kvp.Value;

                string relPath = (string)Settings.Default[settingKey];
                string fullPath = relPath;// Path.IsPathRooted(relPath) ? relPath : Path.Combine(baseDir, relPath);

                // Logger.Log($"Adding UI for '{label}' from setting '{settingKey}' → {fullPath}");

                CsvPathsPanel.Children.Add(new Label
                {
                    Content = label,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 5, 0, 0)
                });

                var row = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 2, 0, 0)
                };

                // the path TextBox
                var txt = new TextBox
                {
                    Text = fullPath,
                    Tag = settingKey,
                    Width = 195
                };

                // the browse‐icon Button
                var btn = new Button
                {
                    Width = 16,
                    Height = 16,
                    Margin = new Thickness(5, 0, 0, 0),
                    Tag = settingKey,
                    Background = System.Windows.Media.Brushes.Transparent,
                    BorderBrush = System.Windows.Media.Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(0),
                };
                // load your browse.png (put it in Resources and mark “Resource” in the project)
                var icon = new System.Windows.Controls.Image
                {
                    Source = new System.Windows.Media.Imaging.BitmapImage(
                        new Uri("pack://application:,,,/AESCConstruct25;component/FrameGenerator/UI/Images/Icon_Upload.png")
                    )
                };
                btn.Content = icon;
                btn.Click += BrowseCsvButton_Click;

                // assemble
                row.Children.Add(txt);
                row.Children.Add(btn);
                CsvPathsPanel.Children.Add(row);
            }
        }

        private void BrowseCsvButton_Click(object sender, RoutedEventArgs e)
        {
            var btn = (Button)sender;
            var key = (string)btn.Tag;
            var dlg = new OpenFileDialog
            {
                Title = $"Select CSV for {key}",
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true)
                return;

            // locate the TextBox whose Tag matches
            var textBox = CsvPathsPanel.Children
                .OfType<StackPanel>()
                .SelectMany(sp => sp.Children.OfType<TextBox>())
                .First(tb => (string)tb.Tag == key);

            textBox.Text = dlg.FileName;
        }

        // ─────────────────────────────────────────────────────────────────────────────
        // EXPORT (System.Text.Json)
        // ─────────────────────────────────────────────────────────────────────────────
        private void ExportSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var map = new Dictionary<string, object>();
                var settings = AESCConstruct25.Properties.Settings.Default;

                foreach (SettingsProperty prop in settings.Properties)
                {
                    var name = prop.Name;
                    var value = settings[name];

                    if (value is StringCollection sc)
                        map[name] = sc.Cast<string>().ToArray();
                    else if (value is Enum)
                        map[name] = value.ToString();
                    else
                        map[name] = value; // string, bool, numeric, etc.
                }

                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(map, options);

                var dlg = new SaveFileDialog
                {
                    Title = "Export settings to JSON",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    FileName = "Construct25.Settings.json",
                    AddExtension = true,
                    DefaultExt = ".json",
                    OverwritePrompt = true
                };

                if (dlg.ShowDialog() == true)
                {
                    File.WriteAllText(dlg.FileName, json);
                    //MessageBox.Show("Settings exported:\n" + dlg.FileName, "Export Settings",
                    //    MessageBoxButton.OK, MessageBoxImage.Information);
                    Application.ReportStatus("Settings exported:\n" + dlg.FileName, StatusMessageType.Information, null);
                }
            }
            catch (System.Exception ex)
            {
                //MessageBox.Show("Export failed:\n" + ex.Message, "Export Settings",
                //    MessageBoxButton.OK, MessageBoxImage.Error);
                Application.ReportStatus("Export failed:\n" + ex.Message, StatusMessageType.Error, null);
            }
        }

        // Keep your existing usings (you already have System.Text.Json, etc.)

        private void ImportSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new OpenFileDialog
                {
                    Title = "Import settings from JSON",
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    CheckFileExists = true,
                    Multiselect = false
                };
                if (dlg.ShowDialog() != true) return;

                var json = File.ReadAllText(dlg.FileName);
                var data = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);
                if (data == null)
                {
                    //MessageBox.Show("Invalid JSON.", "Import Settings",
                    //    MessageBoxButton.OK, MessageBoxImage.Warning);
                    Application.ReportStatus("Invalid JSON.", StatusMessageType.Error, null);
                    return;
                }

                var s = AESCConstruct25.Properties.Settings.Default;

                foreach (var kv in data)
                {
                    try
                    {
                        // Try to get the current value to determine the target type.
                        var current = s[kv.Key];               // throws if setting does not exist
                        var targetType = current?.GetType() ?? typeof(string);

                        // Convert JSON -> target type, then assign.
                        var coerced = CoerceJsonToType(kv.Value, targetType);
                        s[kv.Key] = coerced;
                    }
                    catch
                    {
                        // Unknown setting name or assignment failed: skip silently (by request).
                    }
                }

                s.Save();

                ComboBox_Language.ItemsSource = new[] { "EN", "NL", "FR", "DE", "IT", "ES" };
                ComboBox_Language.SelectedItem = Settings.Default.Construct_Language;

                Label_BOMLengthUnitComboBox.ItemsSource = new[] { "mm", "cm", "m", "inch" };
                Label_BOMLengthUnitComboBox.SelectedItem = Settings.Default.LengthUnit;

                Label_BOMAnchorComboBox.ItemsSource = new[] { "TopLeft", "TopRight", "BottomLeft", "BottomRight" };
                Label_BOMAnchorComboBox.SelectedItem = Settings.Default.DocumentAnchor;

                CornerComboBox.ItemsSource = new[] { "TopLeftCorner", "TopRightCorner", "BottomLeftCorner", "BottomRightCorner" };
                CornerComboBox.SelectedItem = Settings.Default.TableLocationPoint;

                // Text fields
                PartNameTextBox.Text = Settings.Default.NameString ?? string.Empty;
                SerialNumberTextBox.Text = Settings.Default.SerialNumber ?? string.Empty;

                // Checkboxes
                Checkbox_ExcelMaterial.IsChecked = Settings.Default.MatInExcel;
                Checkbox_BOMMaterial.IsChecked = Settings.Default.MatInBOM;
                Checkbox_STEPMaterial.IsChecked = Settings.Default.MatInSTEP;

                // Anchor coordinates
                AnchorXTextBox.Text = Settings.Default.TableAnchorX.ToString();
                AnchorYTextBox.Text = Settings.Default.TableAnchorY.ToString();

                // Rebuild the type/name grid from TypeString
                _pairs.Clear();
                var raw = Settings.Default.TypeString ?? "";
                foreach (var entry in raw.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = entry.Split(new[] { '@' }, 2);
                    _pairs.Add(parts.Length == 2
                        ? new TypeNamePair { Type = parts[0], Name = parts[1] }
                        : new TypeNamePair { Type = entry, Name = "" });
                }
                // If you set ItemsSource elsewhere already, no need to reassign; otherwise:
                // TypesDataGrid.ItemsSource = _pairs;

                // Rebuild CSV path rows from settings
                PopulateCsvFilePathFields();

                // If language changed via import, re-localize UI and refresh command texts
                Localization.Language.LocalizeFrameworkElement(this);
                UIMain.UIManager.UpdateCommandTexts();
                Construct25.UpdateCommandTexts();

                //MessageBox.Show("Settings imported.", "Import Settings",
                //    MessageBoxButton.OK, MessageBoxImage.Information);
                Application.ReportStatus("Settings imported.", StatusMessageType.Information, null);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Import failed:\n" + ex.Message, "Import Settings",
                //    MessageBoxButton.OK, MessageBoxImage.Error);
                Application.ReportStatus("Import failed:\n" + ex.Message, StatusMessageType.Error, null);
            }
        }

        private static object CoerceJsonToType(JsonElement je, Type targetType)
        {
            try
            {
                if (targetType == typeof(string))
                {
                    return je.ValueKind == JsonValueKind.Object || je.ValueKind == JsonValueKind.Array
                        ? JsonSerializer.Serialize(je)
                        : (je.ValueKind == JsonValueKind.Null ? string.Empty : je.ToString());
                }
                else if (targetType == typeof(bool))
                {
                    if (je.ValueKind == JsonValueKind.True || je.ValueKind == JsonValueKind.False)
                        return je.GetBoolean();
                    var s = je.ToString();
                    if (bool.TryParse(s, out var b)) return b;
                    if (double.TryParse(s, System.Globalization.NumberStyles.Float,
                                        System.Globalization.CultureInfo.InvariantCulture, out var n))
                        return Math.Abs(n) > double.Epsilon;
                    return false;
                }
                else if (targetType == typeof(int))
                {
                    if (je.ValueKind == JsonValueKind.Number && je.TryGetInt32(out var i)) return i;
                    return int.Parse(je.ToString(), System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (targetType == typeof(long))
                {
                    if (je.ValueKind == JsonValueKind.Number && je.TryGetInt64(out var l)) return l;
                    return long.Parse(je.ToString(), System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (targetType == typeof(double))
                {
                    if (je.ValueKind == JsonValueKind.Number && je.TryGetDouble(out var d)) return d;
                    return double.Parse(je.ToString(), System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (targetType == typeof(float))
                {
                    if (je.ValueKind == JsonValueKind.Number && je.TryGetSingle(out var f)) return f;
                    return float.Parse(je.ToString(), System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (targetType == typeof(decimal))
                {
                    if (je.ValueKind == JsonValueKind.Number && je.TryGetDecimal(out var m)) return m;
                    return decimal.Parse(je.ToString(), System.Globalization.CultureInfo.InvariantCulture);
                }
                else if (targetType.IsEnum)
                {
                    var s = je.ValueKind == JsonValueKind.String ? je.GetString() : je.ToString();
                    return Enum.Parse(targetType, s ?? string.Empty, ignoreCase: true);
                }
                else if (targetType == typeof(StringCollection))
                {
                    var sc = new StringCollection();
                    if (je.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var el in je.EnumerateArray())
                            sc.Add(el.ValueKind == JsonValueKind.String ? el.GetString() : el.ToString());
                    }
                    else if (je.ValueKind != JsonValueKind.Null && je.ValueKind != JsonValueKind.Undefined)
                    {
                        sc.Add(je.ToString());
                    }
                    return sc;
                }

                // Fallback: generic change type via string
                return Convert.ChangeType(je.ToString(), targetType, System.Globalization.CultureInfo.InvariantCulture);
            }
            catch
            {
                // If conversion fails, fall back to string or default
                if (targetType == typeof(StringCollection))
                {
                    var sc = new StringCollection();
                    if (je.ValueKind != JsonValueKind.Null && je.ValueKind != JsonValueKind.Undefined)
                        sc.Add(je.ToString());
                    return sc;
                }
                if (targetType == typeof(string)) return je.ToString();
                return null;
            }
        }

    }
}