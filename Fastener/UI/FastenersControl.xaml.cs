using AESCConstruct25.Fastener.Module;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace AESCConstruct25.UI
{
    public partial class FastenersControl : UserControl, INotifyPropertyChanged
    {
        private readonly FastenerModule _module;

        public FastenersControl()
        {
            InitializeComponent();
            DataContext = this;

            // 1) load data once
            _module = new FastenerModule();

            // 2) populate type lists
            BoltTypeOptions = new ObservableCollection<string>(_module.BoltTypes);
            WasherTypeOptions = new ObservableCollection<string>(_module.WasherTypes);
            NutTypeOptions = new ObservableCollection<string>(_module.NutTypes);

            // 3) pick initial selections
            SelectedBoltType = BoltTypeOptions.FirstOrDefault();
            SelectedWasherTopType = WasherTypeOptions.FirstOrDefault();
            SelectedWasherBottomType = WasherTypeOptions.FirstOrDefault();
            SelectedNutType = NutTypeOptions.FirstOrDefault();

            IncludeWasherTop = false;
            IncludeWasherBottom = false;
            IncludeNut = false;

            // 4) fill dependent size lists
            RefreshBoltSizes();
            RefreshWasherSizes();
            RefreshNutSizes();
        }

        // INotifyPropertyChanged...
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string prop)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));

        // --- Bolt ---
        public ObservableCollection<string> BoltTypeOptions { get; }
        public ObservableCollection<string> BoltSizeOptions { get; private set; }

        private string _selectedBoltType;
        public string SelectedBoltType
        {
            get => _selectedBoltType;
            set
            {
                if (_selectedBoltType == value) return;
                _selectedBoltType = value;
                OnPropertyChanged(nameof(SelectedBoltType));
                RefreshBoltSizes();
            }
        }

        public string SelectedBoltSize { get; set; }
        public string BoltLength { get; set; } = "35";

        // --- Washers ---
        public ObservableCollection<string> WasherTypeOptions { get; }
        public ObservableCollection<string> WasherSizeOptions { get; private set; }

        public bool IncludeWasherTop { get; set; }
        public string SelectedWasherTopType { get; set; }
        public string SelectedWasherTopSize { get; set; }

        public bool IncludeWasherBottom { get; set; }
        public string SelectedWasherBottomType { get; set; }
        public string SelectedWasherBottomSize { get; set; }

        // --- Nuts ---
        public ObservableCollection<string> NutTypeOptions { get; }
        public ObservableCollection<string> NutSizeOptions { get; private set; }

        public bool IncludeNut { get; set; }
        public string SelectedNutType { get; set; }
        public string SelectedNutSize { get; set; }

        // --- Custom & insert ---
        public bool UseCustomPart { get; set; }
        public ObservableCollection<string> CustomFileOptions { get; } = new ObservableCollection<string>();
        public string SelectedCustomFile { get; set; }

        public int DistanceMm { get; set; } = 20;
        public bool LockDistance { get; set; }

        // Called when Bolt-Type changes
        private void RefreshBoltSizes()
        {
            var sizes = _module.BoltSizesFor(SelectedBoltType).ToList();
            BoltSizeOptions = new ObservableCollection<string>(sizes);
            SelectedBoltSize = sizes.FirstOrDefault();
            OnPropertyChanged(nameof(BoltSizeOptions));
            OnPropertyChanged(nameof(SelectedBoltSize));
        }

        // Called when any Washer-Type changes (we use top type for both)
        private void RefreshWasherSizes()
        {
            var sizes = _module.WasherSizesFor(SelectedWasherTopType).ToList();
            WasherSizeOptions = new ObservableCollection<string>(sizes);
            SelectedWasherTopSize = sizes.FirstOrDefault();
            SelectedWasherBottomSize = sizes.FirstOrDefault();
            OnPropertyChanged(nameof(WasherSizeOptions));
            OnPropertyChanged(nameof(SelectedWasherTopSize));
            OnPropertyChanged(nameof(SelectedWasherBottomSize));
        }

        // Called when Nut-Type changes
        private void RefreshNutSizes()
        {
            var sizes = _module.NutSizesFor(SelectedNutType).ToList();
            NutSizeOptions = new ObservableCollection<string>(sizes);
            SelectedNutSize = sizes.FirstOrDefault();
            OnPropertyChanged(nameof(NutSizeOptions));
            OnPropertyChanged(nameof(SelectedNutSize));
        }


        private void GetSizesButton_Click(object sender, RoutedEventArgs e)
        {
            // you can call your CheckSize logic here
            FastenerModule.CheckSize();
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            // push UI state into module
            _module.SetBoltType(SelectedBoltType);
            _module.SetBoltSize(SelectedBoltSize);
            _module.SetBoltLength(BoltLength);
            _module.SetIncludeWasherTop(IncludeWasherTop);
            _module.SetWasherTopType(SelectedWasherTopType);
            _module.SetWasherTopSize(SelectedWasherTopSize);
            _module.SetIncludeWasherBottom(IncludeWasherBottom);
            _module.SetWasherBottomType(SelectedWasherBottomType);
            _module.SetWasherBottomSize(SelectedWasherBottomSize);
            _module.SetIncludeNut(IncludeNut);
            _module.SetNutType(SelectedNutType);
            _module.SetNutSize(SelectedNutSize);
            _module.SetUseCustomPart(UseCustomPart);
            _module.SetCustomFile(SelectedCustomFile);
            _module.SetDistance(DistanceMm);
            _module.SetLockDistance(LockDistance);

            // then invoke creation
            _module.CreateFasteners();
        }
    }
}
