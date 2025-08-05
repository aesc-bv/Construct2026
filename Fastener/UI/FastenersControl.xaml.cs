using AESCConstruct25.Fastener.Module;
using AESCConstruct25.FrameGenerator.Utilities;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using UserControl = System.Windows.Controls.UserControl;

namespace AESCConstruct25.UI
{
    public partial class FastenersControl : UserControl, INotifyPropertyChanged
    {
        private readonly FastenerModule _module;
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

            Logger.Log($"AESCConstruct25: Selected custom fastener file – {selectedPath}");

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
