using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace AESCConstruct25.UI
{
    public partial class FastenersControl : UserControl
    {
        public FastenersControl()
        {
            InitializeComponent();
            DataContext = this;

            // initialize defaults
            SelectedBoltType = BoltTypeOptions.FirstOrDefault();
            SelectedBoltSize = BoltSizeOptions.ElementAtOrDefault(1);
            BoltLength = "35";

            IncludeWasherTop = true;
            SelectedWasherTopType = WasherTypeOptions.FirstOrDefault();
            SelectedWasherTopSize = WasherSizeOptions.ElementAtOrDefault(1);

            IncludeWasherBottom = true;
            SelectedWasherBottomType = WasherTypeOptions.FirstOrDefault();
            SelectedWasherBottomSize = WasherSizeOptions.ElementAtOrDefault(1);

            IncludeNut = true;
            SelectedNutType = NutTypeOptions.FirstOrDefault();
            SelectedNutSize = NutSizeOptions.ElementAtOrDefault(1);

            UseCustomPart = false;
            // populate CustomFileOptions here if needed
        }

        // --- Bolt ---
        public ObservableCollection<string> BoltTypeOptions { get; } = new ObservableCollection<string>
        {
            "ISO4014 / DIN931", "ISO4017 / DIN933"
        };
        public ObservableCollection<string> BoltSizeOptions { get; } = new ObservableCollection<string>
        {
            "M6", "M8", "M10"
        };

        public string SelectedBoltType { get; set; }
        public string SelectedBoltSize { get; set; }
        public string BoltLength { get; set; }

        // --- Washers ---
        public ObservableCollection<string> WasherTypeOptions { get; } = new ObservableCollection<string>
        {
            "DIN125A - ISO 7089"
        };
        public ObservableCollection<string> WasherSizeOptions { get; } = new ObservableCollection<string>
        {
            "M6", "M8"
        };

        public bool IncludeWasherTop { get; set; }
        public string SelectedWasherTopType { get; set; }
        public string SelectedWasherTopSize { get; set; }

        public bool IncludeWasherBottom { get; set; }
        public string SelectedWasherBottomType { get; set; }
        public string SelectedWasherBottomSize { get; set; }

        // --- Nut ---
        public bool IncludeNut { get; set; }
        public ObservableCollection<string> NutTypeOptions { get; } = new ObservableCollection<string>
        {
            "DIN 985"
        };
        public ObservableCollection<string> NutSizeOptions { get; } = new ObservableCollection<string>
        {
            "M6", "M8"
        };

        public string SelectedNutType { get; set; }
        public string SelectedNutSize { get; set; }

        // --- Custom Part ---
        public bool UseCustomPart { get; set; }
        public ObservableCollection<string> CustomFileOptions { get; } = new ObservableCollection<string>();
        public string SelectedCustomFile { get; set; }

        // --- Distance & Insert ---
        public int DistanceMm { get; set; } = 20;
        public bool LockDistance { get; set; }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            var parms = new FastenerParameters
            {
                BoltType = SelectedBoltType,
                BoltSize = SelectedBoltSize,
                BoltLength = BoltLength,
                IncludeWasherTop = IncludeWasherTop,
                WasherTopType = SelectedWasherTopType,
                WasherTopSize = SelectedWasherTopSize,
                IncludeWasherBottom = IncludeWasherBottom,
                WasherBottomType = SelectedWasherBottomType,
                WasherBottomSize = SelectedWasherBottomSize,
                IncludeNut = IncludeNut,
                NutType = SelectedNutType,
                NutSize = SelectedNutSize,
                UseCustomPart = UseCustomPart,
                CustomFilePath = SelectedCustomFile,
                DistanceMm = DistanceMm,
                LockDistance = LockDistance
            };

            // TODO: call your SpaceClaim API with `parms`
        }

        private void GetSizesButton_Click(object sender, RoutedEventArgs e)
        {
            // TO DO
        }
    }

    public class FastenerParameters
    {
        public string BoltType { get; set; }
        public string BoltSize { get; set; }
        public string BoltLength { get; set; }
        public bool IncludeWasherTop { get; set; }
        public string WasherTopType { get; set; }
        public string WasherTopSize { get; set; }
        public bool IncludeWasherBottom { get; set; }
        public string WasherBottomType { get; set; }
        public string WasherBottomSize { get; set; }
        public bool IncludeNut { get; set; }
        public string NutType { get; set; }
        public string NutSize { get; set; }
        public bool UseCustomPart { get; set; }
        public string CustomFilePath { get; set; }
        public int DistanceMm { get; set; }
        public bool LockDistance { get; set; }
    }
}
