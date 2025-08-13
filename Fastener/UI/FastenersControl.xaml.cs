using AESCConstruct25.Fastener.Module;
using SpaceClaim.Api.V242;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using UserControl = System.Windows.Controls.UserControl;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct25.UI
{
    public partial class FastenersControl : UserControl, INotifyPropertyChanged
    {
        private readonly FastenerModule _module;

        private bool _overwriteDistance;
        public bool OverwriteDistance
        {
            get => _overwriteDistance;
            set
            {
                if (_overwriteDistance == value) return;
                _overwriteDistance = value;
                OnPropertyChanged(nameof(OverwriteDistance));
            }
        }
        private string CustomFolderPath =>
           Path.Combine(
               Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
               "AESCConstruct",
               "Fasteners",
               "Custom"
           );

        public FastenersControl()
        {
            InitializeComponent();
            DataContext = this;
            Localization.Language.LocalizeFrameworkElement(this);

            // 1) load data once
            _module = new FastenerModule();

            // 2) populate type lists
            //BoltTypeOptions = new ObservableCollection<string>(_module.BoltTypes);
            //WasherTypeOptions = new ObservableCollection<string>(_module.WasherTypes);
            //NutTypeOptions = new ObservableCollection<string>(_module.NutTypes);
            BoltTypeOptions = new ObservableCollection<string>(_module.BoltNames);
            WasherTopTypeOptions = new ObservableCollection<string>(_module.WasherNames);
            WasherBottomTypeOptions = new ObservableCollection<string>(_module.WasherNames);
            NutTypeOptions = new ObservableCollection<string>(_module.NutNames);

            // 3) pick initial selections
            SelectedBoltType = BoltTypeOptions.FirstOrDefault();
            SelectedWasherTopType = WasherTopTypeOptions.FirstOrDefault();
            SelectedWasherBottomType = WasherBottomTypeOptions.FirstOrDefault();
            SelectedNutType = NutTypeOptions.FirstOrDefault();

            IncludeWasherTop = false;
            IncludeWasherBottom = false;
            IncludeNut = false;

            // 4) fill dependent size lists
            RefreshBoltSizes();
            //RefreshWasherSizes();
            RefreshWasherTopSizes();
            RefreshWasherBottomSizes();
            RefreshNutSizes();

            LoadCustomFileList();
        }

        private void LoadCustomFileList()
        {
            Directory.CreateDirectory(CustomFolderPath);

            var files = Directory
                .EnumerateFiles(CustomFolderPath, "*.scdoc")
                .Concat(Directory.EnumerateFiles(CustomFolderPath, "*.scdocx"))
                .Concat(Directory.EnumerateFiles(CustomFolderPath, "*.stp"))
                .Select(Path.GetFileName)
                .OrderBy(n => n);

            CustomFileOptions.Clear();
            foreach (var name in files)
                CustomFileOptions.Add(name);

            // notify the list itself changed
            OnPropertyChanged(nameof(CustomFileOptions));

            // pick and notify the initial selection
            SelectedCustomFile = CustomFileOptions.FirstOrDefault();
        }

        // Handles the Browse… button click
        private void BrowseCustomButton_Click(object sender, RoutedEventArgs e)
        {
            string selectedPath;
            using (var dlg = new OpenFileDialog
            {
                Title = "Select a Fastener Part to Import",
                Filter = "SpaceClaim Docs (*.scdoc)|*.scdoc|STEP Files (*.stp)|*.stp|SpaceClaim Docs (*.scdocx)|*.scdocx"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                selectedPath = dlg.FileName;
            }

            // Logger.Log($"AESCConstruct25: Selected custom fastener file – {selectedPath}");

            // Copy into the Custom folder
            var customDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                "AESCConstruct",
                "Fasteners",
                "Custom"
            );
            Directory.CreateDirectory(customDir);

            var destPath = Path.Combine(customDir, Path.GetFileName(selectedPath));
            try
            {
                File.Copy(selectedPath, destPath, overwrite: true);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Failed to import custom fastener:\n{ex.Message}",
                    "Import Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
                return;
            }

            // Refresh the dropdown list and select the new file
            LoadCustomFileList();                        // your existing method
            SelectedCustomFile = Path.GetFileName(selectedPath);
            OnPropertyChanged(nameof(SelectedCustomFile));

            System.Windows.MessageBox.Show(
                "Custom fastener imported successfully.",
                "Import Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );
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
        public ObservableCollection<string> WasherTopTypeOptions { get; }
        public ObservableCollection<string> WasherTopSizeOptions { get; private set; }
        public ObservableCollection<string> WasherBottomTypeOptions { get; }
        public ObservableCollection<string> WasherBottomSizeOptions { get; private set; }

        public bool IncludeWasherTop { get; set; }

        private string _selectedWasherTopType;
        public string SelectedWasherTopType
        {
            get => _selectedWasherTopType;
            set
            {
                if (_selectedWasherTopType == value) return;
                _selectedWasherTopType = value;
                OnPropertyChanged(nameof(SelectedWasherTopType));
                RefreshWasherTopSizes();
            }
        }

        public string SelectedWasherTopSize { get; set; }

        public bool IncludeWasherBottom { get; set; }

        private string _selectedWasherBottomType;
        public string SelectedWasherBottomType
        {
            get => _selectedWasherBottomType;
            set
            {
                if (_selectedWasherBottomType == value) return;
                _selectedWasherBottomType = value;
                OnPropertyChanged(nameof(SelectedWasherBottomType));
                RefreshWasherBottomSizes();
            }
        }

        public string SelectedWasherBottomSize { get; set; }

        // --- Nuts ---
        public ObservableCollection<string> NutTypeOptions { get; }
        public ObservableCollection<string> NutSizeOptions { get; private set; }

        public bool IncludeNut { get; set; }

        private string _selectedNutType;
        public string SelectedNutType
        {
            get => _selectedNutType;
            set
            {
                if (_selectedNutType == value) return;
                _selectedNutType = value;
                OnPropertyChanged(nameof(SelectedNutType));
                RefreshNutSizes();
            }
        }
        public string SelectedNutSize { get; set; }

        // --- Custom & insert ---
        private bool _useCustomPart;
        public bool UseCustomPart
        {
            get => _useCustomPart;
            set
            {
                if (_useCustomPart == value) return;
                _useCustomPart = value;
                OnPropertyChanged(nameof(UseCustomPart));

                // When toggled on, refresh the list
                if (_useCustomPart)
                    LoadCustomFileList();
            }
        }
        public ObservableCollection<string> CustomFileOptions { get; } = new ObservableCollection<string>();
        private string _selectedCustomFile;
        public string SelectedCustomFile
        {
            get => _selectedCustomFile;
            set
            {
                if (_selectedCustomFile == value) return;
                _selectedCustomFile = value;
                OnPropertyChanged(nameof(SelectedCustomFile));
            }
        }

        public int DistanceMm { get; set; } = 20;
        public bool LockDistance { get; set; }

        // Called when Bolt-Type changes
        private void RefreshBoltSizes()
        {
            var sizes = _module.BoltSizesFor(SelectedBoltType).ToList();
            BoltSizeOptions = new ObservableCollection<string>(sizes);
            SelectedBoltSize = BoltSizeOptions.FirstOrDefault();
            OnPropertyChanged(nameof(BoltSizeOptions));
            OnPropertyChanged(nameof(SelectedBoltSize));
        }

        // Called when any Washer-Type changes (we use top type for both)
        //private void RefreshWasherSizes()
        //{
        //    var sizes = _module.WasherSizesFor(SelectedWasherTopType).ToList();
        //    WasherSizeOptions = new ObservableCollection<string>(sizes);
        //    SelectedWasherTopSize = sizes.FirstOrDefault();
        //    SelectedWasherBottomSize = sizes.FirstOrDefault();
        //    OnPropertyChanged(nameof(WasherSizeOptions));
        //    OnPropertyChanged(nameof(SelectedWasherTopSize));
        //    OnPropertyChanged(nameof(SelectedWasherBottomSize));
        //}

        /// <summary>
        /// Called when SelectedWasherTopType changes.
        /// Populates only the Top washer‐size list.
        /// </summary>
        private void RefreshWasherTopSizes()
        {
            var sizes = _module.WasherSizesFor(SelectedWasherTopType).ToList();
            WasherTopSizeOptions = new ObservableCollection<string>(sizes);
            SelectedWasherTopSize = WasherTopSizeOptions.FirstOrDefault();
            OnPropertyChanged(nameof(WasherTopSizeOptions));
            OnPropertyChanged(nameof(SelectedWasherTopSize));
        }

        /// <summary>
        /// Called when SelectedWasherBottomType changes.
        /// Populates only the Bottom washer‐size list.
        /// </summary>
        private void RefreshWasherBottomSizes()
        {
            var sizes = _module.WasherSizesFor(SelectedWasherBottomType).ToList();
            WasherBottomSizeOptions = new ObservableCollection<string>(sizes);
            SelectedWasherBottomSize = WasherBottomSizeOptions.FirstOrDefault();
            OnPropertyChanged(nameof(WasherBottomSizeOptions));
            OnPropertyChanged(nameof(SelectedWasherBottomSize));
        }

        // Called when Nut-Type changes
        private void RefreshNutSizes()
        {
            var sizes = _module.NutSizesFor(SelectedNutType).ToList();
            NutSizeOptions = new ObservableCollection<string>(sizes);
            SelectedNutSize = NutSizeOptions.FirstOrDefault();
            OnPropertyChanged(nameof(NutSizeOptions));
            OnPropertyChanged(nameof(SelectedNutSize));
        }


        //private void GetSizesButton_Click(object sender, RoutedEventArgs e)
        //{
        //    // you can call your CheckSize logic here
        //    FastenerModule.CheckSize();
        //}
        private void GetSizesButton_Click(object sender, RoutedEventArgs e)
        {
            Window window = Window.ActiveWindow;
            Document doc = window?.Document;
            Part rootPart = doc?.MainPart;

            if (window == null || doc == null || rootPart == null)
                return;

            if (!FastenerModule.CheckSelectedCircle(window, true))
                return;

            double radiusMM = FastenerModule.GetSizeCircle(window, out double depthMM);
            // Logger.Log($"radiusMM - {radiusMM}");

            if (radiusMM == 0)
                return;

            void SelectClosestSize(System.Windows.Controls.ComboBox comboBox)
            {
                if (comboBox == null || comboBox.Items.Count == 0)
                    return;

                double maxSize = 0.0;
                int selectedIndex = -1;

                for (int i = 0; i < comboBox.Items.Count; i++)
                {
                    string item = comboBox.Items[i].ToString().Trim();

                    if (item.StartsWith("M", StringComparison.OrdinalIgnoreCase) &&
                        double.TryParse(item.Substring(1), NumberStyles.Any, CultureInfo.InvariantCulture, out double diameter))
                    {
                        // Logger.Log($"====");
                        // Logger.Log($"diameter / 2.0 = {diameter / 2.0}");
                        // Logger.Log($"radiusMM = {radiusMM}");
                        // Logger.Log($"diameter = {diameter}");
                        // Logger.Log($"maxSize = {maxSize}");
                        // Logger.Log($"(diameter / 2.0) - radiusMM < 1e-6 = {(diameter / 2.0) - radiusMM < 1e-6}");
                        // Logger.Log($"diameter > maxSize = {diameter > maxSize}");
                        if ((diameter / 2.0) - radiusMM < 1e-6 && diameter > maxSize)
                        {
                            maxSize = diameter;
                            selectedIndex = i;
                        }
                    }
                }

                if (selectedIndex != -1)
                    comboBox.SelectedIndex = selectedIndex;
            }

            // Apply to all relevant ComboBoxes:
            SelectClosestSize(BoltSizeCombo);
            SelectClosestSize(NutSizeCombo);
            SelectClosestSize(WasherTopSizeCombo);
            SelectClosestSize(WasherBottomSizeCombo);
        }

        private void InsertButton_Click(object sender, RoutedEventArgs e)
        {
            // push UI state into module
            //_module.SetBoltType(SelectedBoltType);
            // Convert the UI‐picked Name into the actual bolt.type
            var realBoltType = _module.GetBoltTypeByName(SelectedBoltType);
            _module.SetBoltType(realBoltType);
            _module.SetBoltSize(SelectedBoltSize);
            _module.SetBoltLength(BoltLength);
            _module.SetIncludeWasherTop(IncludeWasherTop);
            //_module.SetWasherTopType(SelectedWasherTopType);
            var realWasherTopType = _module.GetWasherTopTypeByName(SelectedWasherTopType);
            _module.SetWasherTopType(realWasherTopType);
            _module.SetWasherTopSize(SelectedWasherTopSize);
            _module.SetIncludeWasherBottom(IncludeWasherBottom);
            //_module.SetWasherBottomType(SelectedWasherBottomType);
            var realWasherBottomType = _module.GetWasherTopTypeByName(SelectedWasherBottomType);
            _module.SetWasherBottomType(realWasherBottomType);
            _module.SetWasherBottomSize(SelectedWasherBottomSize);
            _module.SetIncludeNut(IncludeNut);
            //_module.SetNutType(SelectedNutType);
            var realNutType = _module.GetNutTypeByName(SelectedNutType);
            _module.SetNutType(realNutType);
            _module.SetNutSize(SelectedNutSize);
            _module.SetUseCustomPart(UseCustomPart);
            _module.SetCustomFile(SelectedCustomFile);
            _module.SetDistance(DistanceMm);
            _module.SetLockDistance(LockDistance);
            _module.SetOverwriteDistance(OverwriteDistance);

            // then invoke creation
            _module.CreateFasteners();
        }
    }
}
