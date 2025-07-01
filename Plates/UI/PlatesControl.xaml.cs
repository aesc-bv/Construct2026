using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace AESCConstruct25.UI
{
    public partial class PlatesControl : UserControl
    {
        public PlatesControl()
        {
            InitializeComponent();
            DataContext = this;

            // dummy defaults
            ProfileTypes = new ObservableCollection<string> { "HEA foot", "IPE", "HSS" };
            ProfileSizes = new ObservableCollection<string> { "HE100A", "HEA200", "IPE100" };
            SelectedProfileType = ProfileTypes[0];
            SelectedProfileSize = ProfileSizes[0];
            Angle = "0";

            L1 = "220"; L2 = "160"; Count1 = "2";
            B1 = "220"; B2 = "160"; Count2 = "2";
            HoleDiameter = "14"; Thickness = "10"; FilletRadius = "0";
        }

        // Top dropdowns
        public ObservableCollection<string> ProfileTypes { get; }
        public ObservableCollection<string> ProfileSizes { get; }
        public string SelectedProfileType { get; set; }
        public string SelectedProfileSize { get; set; }
        public string Angle { get; set; }

        // Dimensions
        public string L1 { get; set; }
        public string L2 { get; set; }
        public string Count1 { get; set; }
        public string B1 { get; set; }
        public string B2 { get; set; }
        public string Count2 { get; set; }

        public string HoleDiameter { get; set; }
        public string Thickness { get; set; }
        public string FilletRadius { get; set; }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            // TODO: pull all bound properties and call your SpaceClaim API
            MessageBox.Show(
              $"Creating plate: {SelectedProfileType} {SelectedProfileSize}, angle {Angle}°\n" +
              $"L1={L1}, L2={L2}, #1={Count1}\n" +
              $"B1={B1}, B2={B2}, #2={Count2}\n" +
              $"D={HoleDiameter}, T={Thickness}, R={FilletRadius}",
              "Info", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
