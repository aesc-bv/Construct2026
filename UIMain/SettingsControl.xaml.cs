using AESCConstruct25.Properties;
using Microsoft.Win32;
using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace AESCConstruct25.FrameGenerator.UI
{
    public partial class SettingsControl : UserControl
    {
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

            // 2) translate every tagged element in _this_ control
            Localization.Language.LocalizeFrameworkElement(this);

            UIMain.UIManager.UpdateCommandTexts();
            Construct25.UpdateCommandTexts();

            PopulateCsvFilePathFields();
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
            var corner = Settings.Default.TableLocationPoint;
            if (CornerComboBox.Items.Contains(corner))
                CornerComboBox.SelectedItem = corner;
            else
                CornerComboBox.SelectedIndex = 0;

            var SerialNumber = Settings.Default.SerialNumber;
            SerialNumberTextBox.Text = SerialNumber;

            LicenseStatusValue.Text = LicenseSpot.LicenseSpot.State.Status;
            // Logger.Log($"exp: {LicenseSpot.LicenseSpot.State.Status}");
            var expiration = LicenseSpot.LicenseSpot.License?.GetTimeLimit()?.EndDate;
            // Logger.Log($"exp: {expiration}");
            LicenseDetailsValue.Text = expiration.HasValue
                ? expiration.Value.ToShortDateString()
                : string.Empty;

            ComboBox_Language.ItemsSource = new[] { "EN", "NL", "FR", "DE", "IT", "ES" };
            ComboBox_Language.SelectedItem = Settings.Default.Construct_Language ?? "EN";

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

            // Initialize license state
            LicenseSpot.LicenseSpot.Initialize();
            LicenseStatusValue.Text = LicenseSpot.LicenseSpot.State.Status;

            // Show expiration date if available
            var expiration = LicenseSpot.LicenseSpot.License?.GetTimeLimit()?.EndDate;
            LicenseDetailsValue.Text = expiration.HasValue
                ? expiration.Value.ToShortDateString()
                : string.Empty;


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

            Settings.Default.Save();
            // persist all settings
            try
            {
                Settings.Default.Save();
                MessageBox.Show("Settings saved.", "Settings", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ActivateLicenseButton_Click(object sender, RoutedEventArgs e)
        {
            var serial = SerialNumberTextBox.Text.Trim();
            if (string.IsNullOrEmpty(serial))
            {
                MessageBox.Show("Please enter a serial number first.", "Activate",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            // Delegate into our static helper
            LicenseSpot.LicenseSpot.Activate(serial);
            // Update UI
            LicenseStatusValue.Text = LicenseSpot.LicenseSpot.State.Status;
            // Logger.Log($"exp: {LicenseSpot.LicenseSpot.State.Status}");
            var expiration = LicenseSpot.LicenseSpot.License?.GetTimeLimit()?.EndDate;
            // Logger.Log($"exp: {expiration}");
            LicenseDetailsValue.Text = expiration.HasValue
                ? expiration.Value.ToShortDateString()
                : string.Empty;
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

            string baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "AESCConstruct"
            );
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
                string fullPath = Path.IsPathRooted(relPath) ? relPath : Path.Combine(baseDir, relPath);

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
                    Tag = settingKey
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

            // Logger.Log("=== Done populating CSV path fields ===");
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
    }
}
