using AESCConstruct25.Plates.Modules;
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
            InsertInMiddle = true;  // checked by default
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
        public bool InsertInMiddle { get; set; } = true;

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            // parse all the bound string properties into numbers
            // (you may want to add TryParse + a validation step)
            double angle = double.Parse(Angle, System.Globalization.CultureInfo.InvariantCulture);
            double l1 = double.Parse(L1, System.Globalization.CultureInfo.InvariantCulture) * 0.001;
            double l2 = double.Parse(L2, System.Globalization.CultureInfo.InvariantCulture) * 0.001;
            int c1 = int.Parse(Count1, System.Globalization.CultureInfo.InvariantCulture);
            double b1 = double.Parse(B1, System.Globalization.CultureInfo.InvariantCulture) * 0.001;
            double b2 = double.Parse(B2, System.Globalization.CultureInfo.InvariantCulture) * 0.001;
            int c2 = int.Parse(Count2, System.Globalization.CultureInfo.InvariantCulture);
            double t = double.Parse(Thickness, System.Globalization.CultureInfo.InvariantCulture) * 0.001;
            double r = double.Parse(FilletRadius, System.Globalization.CultureInfo.InvariantCulture) * 0.001;
            double d = double.Parse(HoleDiameter, System.Globalization.CultureInfo.InvariantCulture) * 0.001;

            //MessageBox.Show(
            //  $"Creating plate: {SelectedProfileType} {SelectedProfileSize}, angle {Angle}°\n" +
            //  $"L1={L1}, L2={L2}, #1={Count1}\n" +
            //  $"B1={B1}, B2={B2}, #2={Count2}\n" +
            //  $"D={HoleDiameter}, T={Thickness}, R={FilletRadius}",
            //  "Info", MessageBoxButton.OK, MessageBoxImage.Information);

            // call into your module
            PlatesModule.CreatePlateFromUI(
                SelectedProfileType,
                SelectedProfileSize,
                angle,
                l1, l2, c1,
                b1, b2, c2,
                t, r, d,
                insertPlateMid: InsertInMiddle    // or true, if you want a “mid” placement
            );

        }
    }
}
