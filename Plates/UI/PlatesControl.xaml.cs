using AESCConstruct25.FrameGenerator.Utilities;
using AESCConstruct25.Plates.Modules;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Settings = AESCConstruct25.Properties.Settings;

namespace AESCConstruct25.UI
{
    public partial class PlatesControl : UserControl, INotifyPropertyChanged
    {
        // Represents one row in PlatesProperties.csv
        private class PlateRecord
        {
            public string Type { get; set; }
            public string Name { get; set; }
            public string L1 { get; set; }
            public string L2 { get; set; }
            public string Count1 { get; set; }
            public string B1 { get; set; }
            public string B2 { get; set; }
            public string Count2 { get; set; }
            public string HoleDiameter { get; set; }
            public string Thickness { get; set; }
            public string FilletRadius { get; set; }
            public bool InsertInMiddle { get; set; }
        }

        private readonly List<PlateRecord> _records = new List<PlateRecord>();

        public PlatesControl()
        {
            InitializeComponent();
            DataContext = this;
            Localization.Language.LocalizeFrameworkElement(this);

            try
            {
                // 1) Load CSV
                var folder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    "AESCConstruct", "Plates");
                //var csvPath = Path.Combine(folder, "PlatesProperties.csv");
                var csvPath = Settings.Default.PlatesProperties;

                if (!File.Exists(csvPath))
                    throw new FileNotFoundException("PlatesProperties.csv not found", csvPath);

                foreach (var line in File.ReadAllLines(csvPath).Skip(1))
                {
                    var cols = line.Split(';');
                    if (cols.Length < 11)
                        continue;

                    _records.Add(new PlateRecord
                    {
                        Type = cols[0].Trim(),
                        Name = cols[1].Trim(),
                        L1 = cols[2].Trim(),
                        L2 = cols[5].Trim(),
                        Count1 = cols[4].Trim(),
                        B1 = cols[5].Trim(),
                        B2 = cols[6].Trim(),
                        Count2 = cols[7].Trim(),
                        HoleDiameter = cols[8].Trim(),
                        Thickness = cols[9].Trim(),
                        FilletRadius = cols[10].Trim(),
                        InsertInMiddle = cols.Length > 11 && bool.TryParse(cols[11].Trim(), out var m) && m
                    });
                }

                // 2) Populate ProfileTypes
                foreach (var t in _records.Select(r => r.Type).Distinct().OrderBy(x => x))
                    ProfileTypes.Add(t);

                SelectedProfileType = ProfileTypes.FirstOrDefault();

                // 3) Logging for debug
                Logger.Log($"[PlatesControl] Loaded {_records.Count} records from CSV.");
                Logger.Log($"[PlatesControl] ProfileTypes: {string.Join(", ", ProfileTypes)}");
                Logger.Log($"[PlatesControl] Initial SelectedProfileType: {SelectedProfileType}");
            }
            catch (Exception ex)
            {
                // Log and alert user—no silent fallback defaults
                Logger.Log($"[PlatesControl] ERROR loading PlatesProperties.csv: {ex}");
                MessageBox.Show(
                    $"Failed to load plate properties:\n{ex.Message}",
                    "Plates Data Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }

            // 4) Even if _records is empty, wire up Size list (may be empty)
            if (SelectedProfileType != null)
                SelectedProfileType = SelectedProfileType; // re-invoke setter to populate sizes

            // 5) Populate lower fields if possible
            PopulateBottomFields();
        }


        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string prop) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        #endregion

        // Top dropdowns
        public ObservableCollection<string> ProfileTypes { get; } = new ObservableCollection<string>();
        public ObservableCollection<string> ProfileSizes { get; } = new ObservableCollection<string>();

        private string _selectedProfileType;
        public string SelectedProfileType
        {
            get => _selectedProfileType;
            set
            {
                if (_selectedProfileType == value) return;
                _selectedProfileType = value;
                OnPropertyChanged(nameof(SelectedProfileType));

                UpdatePlateImage();

                // repopulate ProfileSizes
                ProfileSizes.Clear();
                foreach (var name in _records
                    .Where(r => r.Type == value)
                    .Select(r => r.Name)
                    .Distinct()
                    .OrderBy(n => n))
                {
                    ProfileSizes.Add(name);
                }
                SelectedProfileSize = ProfileSizes.FirstOrDefault();
                OnPropertyChanged(nameof(ProfileSizes));
            }
        }

        private string _selectedImageSource;
        public string SelectedImageSource
        {
            get => _selectedImageSource;
            set { _selectedImageSource = value; OnPropertyChanged(nameof(SelectedImageSource)); }
        }

        private string _selectedProfileSize;
        public string SelectedProfileSize
        {
            get => _selectedProfileSize;
            set
            {
                if (_selectedProfileSize == value) return;
                _selectedProfileSize = value;
                OnPropertyChanged(nameof(SelectedProfileSize));
                PopulateBottomFields();
            }
        }

        private void PopulateBottomFields()
        {
            var rec = _records.FirstOrDefault(r =>
                r.Type == SelectedProfileType && r.Name == SelectedProfileSize);
            if (rec == null) return;

            L1 = rec.L1;
            L2 = rec.L2;
            Count1 = rec.Count1;
            B1 = rec.B1;
            B2 = rec.B2;
            Count2 = rec.Count2;
            HoleDiameter = rec.HoleDiameter;
            Thickness = rec.Thickness;
            FilletRadius = rec.FilletRadius;
            InsertInMiddle = rec.InsertInMiddle;
        }

        // Angle (°)
        private string _angle = "0";
        public string Angle
        {
            get => _angle;
            set { _angle = value; OnPropertyChanged(nameof(Angle)); }
        }

        // Bottom‐field properties with notifications
        private string _l1;
        public string L1
        {
            get => _l1;
            set { _l1 = value; OnPropertyChanged(nameof(L1)); }
        }

        private string _l2;
        public string L2
        {
            get => _l2;
            set { _l2 = value; OnPropertyChanged(nameof(L2)); }
        }

        private string _count1;
        public string Count1
        {
            get => _count1;
            set { _count1 = value; OnPropertyChanged(nameof(Count1)); }
        }

        private string _b1;
        public string B1
        {
            get => _b1;
            set { _b1 = value; OnPropertyChanged(nameof(B1)); }
        }

        private string _b2;
        public string B2
        {
            get => _b2;
            set { _b2 = value; OnPropertyChanged(nameof(B2)); }
        }

        private string _count2;
        public string Count2
        {
            get => _count2;
            set { _count2 = value; OnPropertyChanged(nameof(Count2)); }
        }

        private string _holeDiameter;
        public string HoleDiameter
        {
            get => _holeDiameter;
            set { _holeDiameter = value; OnPropertyChanged(nameof(HoleDiameter)); }
        }

        private string _thickness;
        public string Thickness
        {
            get => _thickness;
            set { _thickness = value; OnPropertyChanged(nameof(Thickness)); }
        }

        private string _filletRadius;
        public string FilletRadius
        {
            get => _filletRadius;
            set { _filletRadius = value; OnPropertyChanged(nameof(FilletRadius)); }
        }

        private bool _insertInMiddle;
        public bool InsertInMiddle
        {
            get => _insertInMiddle;
            set { _insertInMiddle = value; OnPropertyChanged(nameof(InsertInMiddle)); }
        }

        private void UpdatePlateImage()
        {
            string imageSuffix = SelectedProfileType switch
            {
                "HEA base" or "HEA cap" or "IPE base" or "IPE cap" or "HEA support" or "HEB support" => "2",
                "UNP" => "3",
                "Flat flange" or "Blind flange" => "1",
                _ => "1"
            };

            SelectedImageSource = $"/AESCConstruct25;component/Plates/UI/Images/Img_Measures_Plate_{imageSuffix}.png";
        }

        // Create button handler
        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            // parse + convert units (mm → m)
            double angleDeg = double.Parse(Angle, CultureInfo.InvariantCulture);
            double l1_m = double.Parse(L1, CultureInfo.InvariantCulture) * 0.001;
            double l2_m = double.Parse(L2, CultureInfo.InvariantCulture) * 0.001;
            int c1 = int.Parse(Count1, CultureInfo.InvariantCulture);
            double b1_m = double.Parse(B1, CultureInfo.InvariantCulture) * 0.001;
            double b2_m = double.Parse(B2, CultureInfo.InvariantCulture) * 0.001;
            int c2 = int.Parse(Count2, CultureInfo.InvariantCulture);
            double t_m = double.Parse(Thickness, CultureInfo.InvariantCulture) * 0.001;
            double r_m = double.Parse(FilletRadius, CultureInfo.InvariantCulture) * 0.001;
            double d_m = double.Parse(HoleDiameter, CultureInfo.InvariantCulture) * 0.001;

            // call into your module
            PlatesModule.CreatePlateFromUI(
                SelectedProfileType,
                SelectedProfileSize,
                angleDeg,
                l1_m, l2_m, c1,
                b1_m, b2_m, c2,
                t_m, r_m, d_m,
                insertPlateMid: InsertInMiddle
            );
        }
    }
}
