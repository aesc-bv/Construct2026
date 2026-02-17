/*
 SettingsControl is the WPF settings panel for Construct2026.
 It loads and saves user preferences (units, BOM placement, CSV paths, language),
 integrates license activation and status display, and supports JSON export/import of all settings.
*/

using AESCConstruct2026.Licensing;
using AESCConstruct2026.Properties;
using Microsoft.Win32;
using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using Application = SpaceClaim.Api.V242.Application;
using Path = System.IO.Path;


namespace AESCConstruct2026.FrameGenerator.UI
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
            public string Template { get; set; }
        }

        // backing collection
        private readonly ObservableCollection<TypeNamePair> _pairs = new ObservableCollection<TypeNamePair>();

        public ObservableCollection<string> CornerOptions { get; }
          = new ObservableCollection<string>(
              Enum.GetNames(typeof(LocationPoint))
          );

        // Initializes the settings control, hooks Loaded, and sets the DataContext.
        public SettingsControl()
        {
            InitializeComponent();
            this.Loaded += SettingsControl_Loaded;
            DataContext = this;
        }

        // Handles first-time control load: initializes comboboxes, localization and CSV path fields.
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
            Construct2026.UpdateCommandTexts();

            PopulateCsvFilePathFields();
        }

        // Handles subsequent load: populates all fields from Settings and refreshes bindings and license info.
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
                var parts = entry.Split('@');
                _pairs.Add(new TypeNamePair
                {
                    Type = parts.Length >= 1 ? parts[0] : entry,
                    Name = parts.Length >= 2 ? parts[1] : "",
                    Template = parts.Length >= 3 ? parts[2] : ""
                });
            }

            // bind it
            TypesDataGrid.ItemsSource = _pairs;

            // Decimals
            DecimalsTextBox.Text = Settings.Default.NameDecimals.ToString();

            // Frame color
            LoadFrameColorUI();

            // anchor coordinates
            AnchorXTextBox.Text = Settings.Default.TableAnchorX.ToString();
            AnchorYTextBox.Text = Settings.Default.TableAnchorY.ToString();

            var lic = ConstructLicenseSpot.CurrentLicense;
            LicenseTypeValue.Text = lic == null
                ? string.Empty
                : (lic.IsNetwork ? "Network" : "Local");

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
            Construct2026.UpdateCommandTexts();
        }

        // Saves all settings from the UI into the Settings store and updates license and BOM configuration.
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Save Part naming
            Settings.Default.NameString = PartNameTextBox.Text.Trim();

            // Reserialize the list back into TypeString (Type@Name@Template)
            var serialized = string.Join("|",
                _pairs
                    .Where(p => !string.IsNullOrWhiteSpace(p.Type) || !string.IsNullOrWhiteSpace(p.Name))
                    .Select(p => $"{p.Type}@{p.Name}@{p.Template}")
            );
            Settings.Default.TypeString = serialized;

            // Save decimals
            if (int.TryParse(DecimalsTextBox.Text.Trim(), out var dec) && dec >= 0)
                Settings.Default.NameDecimals = dec;

            // Save frame color
            Settings.Default.FrameColor = FrameColorCheckBox.IsChecked == true
                ? FrameColorTextBox.Text.Trim()
                : string.Empty;

            Settings.Default.MatInExcel = Checkbox_ExcelMaterial.IsChecked == true;
            Settings.Default.MatInBOM = Checkbox_BOMMaterial.IsChecked == true;
            Settings.Default.MatInSTEP = Checkbox_STEPMaterial.IsChecked == true;


            // Refresh license state from disk (no modal)
            ConstructLicenseSpot.CheckLicense();

            var lic = ConstructLicenseSpot.CurrentLicense;
            LicenseTypeValue.Text = lic == null
                ? string.Empty
                : (lic.IsNetwork ? "Network" : "Local");

            LicenseStatusValue.Text = ConstructLicenseSpot.Status;
            var expiration = ConstructLicenseSpot.CurrentLicense?.GetTimeLimit()?.EndDate;
            LicenseDetailsValue.Text = expiration.HasValue ? expiration.Value.ToShortDateString() : string.Empty;

            Construct2026.RefreshLicenseUI();

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
                Application.ReportStatus("Settings saved.", StatusMessageType.Information, null);
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Error saving settings:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Tries to activate the license using the entered serial and updates UI/Settings on success.
        private void ActivateLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            var serial = SerialNumberTextBox.Text.Trim();
            if (string.IsNullOrEmpty(serial))
            {
                Application.ReportStatus("Please enter a serial number first.", StatusMessageType.Warning, null);
                return;
            }

            bool ok = ConstructLicenseSpot.TryActivate(serial, out var msg);
            LicenseStatusValue.Text = msg;

            var expiration = ConstructLicenseSpot.CurrentLicense?.GetTimeLimit()?.EndDate;
            LicenseDetailsValue.Text = expiration.HasValue ? expiration.Value.ToShortDateString() : string.Empty;

            var lic = ConstructLicenseSpot.CurrentLicense;
            LicenseTypeValue.Text = lic == null
                ? string.Empty
                : (lic.IsNetwork ? "Network" : "Local");

            if (SerialNumberTextBox.Text is string sn)
                Settings.Default.SerialNumber = sn;

            Construct2026.RefreshLicenseUI();

            if (ok)
                Application.ReportStatus("License activated.", StatusMessageType.Information, null);
        }

        // Updates language setting on selection change and re-localizes the panel and command texts.
        private void ComboBox_Language_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComboBox_Language.SelectedItem is string langCode)
            {
                Settings.Default.Construct_Language = langCode;
                Settings.Default.Save();

                // re-translate this panel
                Localization.Language.LocalizeFrameworkElement(this);

                UIMain.UIManager.UpdateCommandTexts();
                Construct2026.UpdateCommandTexts();
            }
        }

        // Builds the CSV path rows dynamically from Settings, including browse and open-folder buttons.
        private void PopulateCsvFilePathFields()
        {
            CsvPathsPanel.Children.Clear();

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
                string labelText = kvp.Key;
                string settingKey = kvp.Value;

                string relPath = (string)Settings.Default[settingKey];
                string fullPath = relPath;

                var grid = new Grid
                {
                    Margin = new Thickness(0, 5, 0, 0)
                };
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var lbl = new Label
                {
                    Content = labelText,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 0, 2)
                };
                Grid.SetColumn(lbl, 0);
                Grid.SetRow(lbl, 0);
                grid.Children.Add(lbl);

                var txt = new TextBox
                {
                    Text = fullPath,
                    Tag = settingKey
                };
                Grid.SetColumn(txt, 0);
                Grid.SetRow(txt, 1);
                grid.Children.Add(txt);

                var btn = new Button
                {
                    Width = 16,
                    Height = 16,
                    Margin = new Thickness(2, 0, 2, 0),
                    Tag = settingKey + "btn",
                    Background = System.Windows.Media.Brushes.Transparent,
                    BorderBrush = System.Windows.Media.Brushes.Transparent,
                    ToolTip = "Select new file",
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(0),
                    VerticalAlignment = System.Windows.VerticalAlignment.Center
                };
                var icon = new System.Windows.Controls.Image
                {
                    Source = new System.Windows.Media.Imaging.BitmapImage(
                        new Uri("pack://application:,,,/AESCConstruct2026;component/FrameGenerator/UI/Images/Icon_Upload.png")
                    )
                };
                btn.Content = icon;
                btn.Click += BrowseCsvButton_Click;
                Grid.SetColumn(btn, 1);
                Grid.SetRow(btn, 1);
                grid.Children.Add(btn);

                var folderBtn = new Button
                {
                    Width = 16,
                    Height = 16,
                    Margin = new Thickness(2, 0, 2, 0),
                    Tag = settingKey + "folderbtn",
                    Background = System.Windows.Media.Brushes.Transparent,
                    BorderBrush = System.Windows.Media.Brushes.Transparent,
                    ToolTip = "Open file location",
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(0),
                    VerticalAlignment = System.Windows.VerticalAlignment.Center
                };
                var folderIcon = new System.Windows.Controls.Image
                {
                    Source = new System.Windows.Media.Imaging.BitmapImage(
                        new Uri("pack://application:,,,/AESCConstruct2026;component/FrameGenerator/UI/Images/openFolder.png")
                    )
                };
                folderBtn.Content = folderIcon;
                //folderBtn.Click += OpenFolder(fullPath);
                folderBtn.Click += (s, e) =>
                {
                    var path = txt.Text.Trim(); // e.g., the sibling TextBox
                    var dir = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
                    if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "explorer.exe",
                            Arguments = $"\"{dir}\"",
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        Application.ReportStatus("Folder not found.", StatusMessageType.Warning, null);
                    }
                };
                Grid.SetColumn(folderBtn, 2);
                Grid.SetRow(folderBtn, 1);
                grid.Children.Add(folderBtn);

                // Add the grid (containing label, textbox, button) to the panel
                CsvPathsPanel.Children.Add(grid);
            }
        }

        // Handles the CSV browse button, opens file dialog, and writes the chosen path to the matching TextBox.
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

        // Exports all settings (excluding license specific ones) to a JSON file chosen by the user.
        private void ExportSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var map = new Dictionary<string, object>();
                var settings = AESCConstruct2026.Properties.Settings.Default;

                foreach (SettingsProperty prop in settings.Properties)
                {
                    var name = prop.Name;

                    if (string.Equals(name, "SerialNumber", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (string.Equals(name, "LicenseValid", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (string.Equals(name, "NetworkLogUser_Enabled", StringComparison.OrdinalIgnoreCase))
                        continue;

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
                    FileName = "Construct2026.Settings.json",
                    AddExtension = true,
                    DefaultExt = ".json",
                    OverwritePrompt = true
                };

                if (dlg.ShowDialog() == true)
                {
                    File.WriteAllText(dlg.FileName, json);
                    Application.ReportStatus("Settings exported:\n" + dlg.FileName, StatusMessageType.Information, null);
                }
            }
            catch (System.Exception ex)
            {
                Application.ReportStatus("Export failed:\n" + ex.Message, StatusMessageType.Error, null);
            }
        }

        // Imports settings from a JSON file, updates the Settings store and rebuilds the UI state.
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
                    Application.ReportStatus("Invalid JSON.", StatusMessageType.Error, null);
                    return;
                }

                var s = Settings.Default;

                foreach (var kv in data)
                {
                    if (string.Equals(kv.Key, "SerialNumber", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (string.Equals(kv.Key, "LicenseValid", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (string.Equals(kv.Key, "NetworkLogUser_Enabled", StringComparison.OrdinalIgnoreCase))
                        continue;

                    try
                    {
                        // Try to get the current value to determine the target type.
                        var current = s[kv.Key];
                        var targetType = current?.GetType() ?? typeof(string);

                        // Convert JSON -> target type, then assign.
                        var coerced = CoerceJsonToType(kv.Value, targetType);
                        s[kv.Key] = coerced;
                    }
                    catch
                    {
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
                SerialNumberTextBox.Text = Settings.Default.SerialNumber ?? "";

                var lic = ConstructLicenseSpot.CurrentLicense;
                LicenseTypeValue.Text = lic == null
                    ? string.Empty
                    : (lic.IsNetwork ? "Network" : "Local");

                // Checkboxes
                Checkbox_ExcelMaterial.IsChecked = Settings.Default.MatInExcel;
                Checkbox_BOMMaterial.IsChecked = Settings.Default.MatInBOM;
                Checkbox_STEPMaterial.IsChecked = Settings.Default.MatInSTEP;

                // Anchor coordinates
                AnchorXTextBox.Text = Settings.Default.TableAnchorX.ToString();
                AnchorYTextBox.Text = Settings.Default.TableAnchorY.ToString();

                // Rebuild the type/name/template grid from TypeString
                _pairs.Clear();
                var raw = Settings.Default.TypeString ?? "";
                foreach (var entry in raw.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = entry.Split('@');
                    _pairs.Add(new TypeNamePair
                    {
                        Type = parts.Length >= 1 ? parts[0] : entry,
                        Name = parts.Length >= 2 ? parts[1] : "",
                        Template = parts.Length >= 3 ? parts[2] : ""
                    });
                }

                // Decimals
                DecimalsTextBox.Text = Settings.Default.NameDecimals.ToString();

                // Frame color
                LoadFrameColorUI();

                // Rebuild CSV path rows from settings
                PopulateCsvFilePathFields();

                // If language changed via import, re-localize UI and refresh command texts
                Localization.Language.LocalizeFrameworkElement(this);
                UIMain.UIManager.UpdateCommandTexts();
                Construct2026.UpdateCommandTexts();

                Application.ReportStatus("Settings imported.", StatusMessageType.Information, null);
            }
            catch (Exception ex)
            {
                Application.ReportStatus("Import failed:\n" + ex.Message, StatusMessageType.Error, null);
            }
        }

        // Loads the FrameColor setting into the checkbox, textbox, and preview rectangle.
        private void LoadFrameColorUI()
        {
            string fc = Settings.Default.FrameColor ?? "";
            if (!string.IsNullOrWhiteSpace(fc))
            {
                FrameColorCheckBox.IsChecked = true;
                FrameColorTextBox.Text = fc;
            }
            else
            {
                FrameColorCheckBox.IsChecked = false;
                FrameColorTextBox.Text = "#006d8b";
            }
            UpdateFrameColorPreview();
        }

        // Updates the preview rectangle to reflect the current hex value in the textbox.
        private void UpdateFrameColorPreview()
        {
            try
            {
                if (FrameColorCheckBox.IsChecked == true && !string.IsNullOrWhiteSpace(FrameColorTextBox.Text))
                {
                    var c = System.Drawing.ColorTranslator.FromHtml(FrameColorTextBox.Text.Trim());
                    FrameColorPreview.Fill = new SolidColorBrush(
                        System.Windows.Media.Color.FromArgb(c.A, c.R, c.G, c.B));
                }
                else
                {
                    FrameColorPreview.Fill = Brushes.Transparent;
                }
            }
            catch
            {
                FrameColorPreview.Fill = Brushes.Transparent;
            }
        }

        private void FrameColorCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            UpdateFrameColorPreview();
        }

        private void FrameColorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdateFrameColorPreview();
        }

        // Converts a JsonElement into a strongly typed value suitable for assigning to a Settings property.
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
