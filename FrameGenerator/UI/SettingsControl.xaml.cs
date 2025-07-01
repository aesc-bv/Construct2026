using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using AESCConstruct25.Properties;

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

        public SettingsControl()
        {
            InitializeComponent();
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            // 1) Load the Part naming textbox
            PartNameTextBox.Text = Settings.Default.NameString ?? "";

            // 2) Load and parse the TypeString into the ObservableCollection
            _pairs.Clear();
            var raw = Settings.Default.TypeString ?? "";
            var entries = raw.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

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

            // persist
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
    }
}
