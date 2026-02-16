/*
 PlatesControl is the WPF UI for inserting standard plates based on a CSV definition.
 It loads plate definitions, exposes selectable types and sizes, maps CSV fields to UI,
 and passes validated parameters to PlatesModule.CreatePlateFromUI to build geometry in SpaceClaim.
*/

using AESCConstruct2026.Plates.Modules;
using SpaceClaim.Api.V242;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Application = SpaceClaim.Api.V242.Application;
using Settings = AESCConstruct2026.Properties.Settings;

namespace AESCConstruct2026.UI
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

        // Initializes the plate control, loads CSV records, populates type list, and sets up localization.
        public PlatesControl()
        {
            InitializeComponent();
            DataContext = this;
            Localization.Language.LocalizeFrameworkElement(this);

            try
            {
                var csvPath = Settings.Default.PlatesProperties;

                if (!File.Exists(csvPath))
                    throw new FileNotFoundException($"{Settings.Default.PlatesProperties} not found", csvPath);

                int lineNum = 1;
                foreach (var line in File.ReadAllLines(csvPath).Skip(1))
                {
                    lineNum++;
                    try
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
                            Count1 = cols[6].Trim(),
                            B1 = cols[3].Trim(),
                            B2 = cols[7].Trim(),
                            Count2 = cols[8].Trim(),
                            HoleDiameter = cols[9].Trim(),
                            Thickness = cols[4].Trim(),
                            FilletRadius = cols[10].Trim(),
                            InsertInMiddle = cols.Length > 11 && bool.TryParse(cols[11].Trim(), out var m) && m
                        });
                    }
                    catch (Exception ex)
                    {
                        Application.ReportStatus($"Error parsing PlatesProperties.csv at line {lineNum}:\n{ex.Message}", StatusMessageType.Error, null);
                    }
                }

                foreach (var t in _records.Select(r => r.Type).Distinct().OrderBy(x => x))
                    ProfileTypes.Add(t);

                SelectedProfileType = ProfileTypes.FirstOrDefault();
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to load plate properties:\n{ex.Message}", StatusMessageType.Error, null);
            }

            if (SelectedProfileType != null)
                SelectedProfileType = SelectedProfileType;

            PopulateBottomFields();
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        // Raises the PropertyChanged event for WPF data binding.
        private void OnPropertyChanged(string prop) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        #endregion

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

                ProfileSizes.Clear();

                var names = _records
                    .Where(r => r.Type == value)
                    .Select(r => r.Name)
                    .Distinct()
                    .ToList();

                IEnumerable<string> ordered =
                    names.All(IsNumericString)
                        ? names.OrderBy(n => ParseDoubleSafe(n))
                        : names.OrderBy(n => NaturalSortKey(n), StringComparer.OrdinalIgnoreCase);

                foreach (var name in ordered)
                    ProfileSizes.Add(name);

                SelectedProfileSize = ProfileSizes.FirstOrDefault();
                OnPropertyChanged(nameof(ProfileSizes));
                UpdateFieldLabelsAndVisibility();
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

        // Copies the currently selected plate record values into the bottom input fields.
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
            //InsertInMiddle = rec.InsertInMiddle;
        }

        private string _angle = "0";
        public string Angle
        {
            get => _angle;
            set { _angle = value; OnPropertyChanged(nameof(Angle)); }
        }

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

        private bool _insertInMiddle = true;
        public bool InsertInMiddle
        {
            get => _insertInMiddle;
            set { _insertInMiddle = value; OnPropertyChanged(nameof(InsertInMiddle)); }
        }

        // Updates the measure image path according to the selected plate type.
        private void UpdatePlateImage()
        {
            string imageSuffix = SelectedProfileType switch
            {
                "HEA base" or "HEA cap" or "IPE base" or "IPE cap" or "HEA support" or "HEB support" => "2",
                "UNP" => "3",
                "Blind flange" => "4",
                "Flat flange" => "1",
                _ => "1"
            };

            SelectedImageSource = $"/AESCConstruct2026;component/Plates/UI/Images/Img_Measures_Plate_{imageSuffix}.png";
        }

        // Parses a string to double using invariant culture, returning 0 on failure.
        private double ParseDoubleSafe(string text)
        {
            double val;
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out val) ? val : 0;
        }

        // Checks whether the string is numeric (optionally with decimal separator).
        private static bool IsNumericString(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            // Accept "12", "12.5", "12,5"
            return System.Text.RegularExpressions.Regex.IsMatch(s.Trim(), @"^\d+(?:[.,]\d+)?$");
        }

        // Builds a key for natural (human-friendly) sorting of strings with embedded numbers.
        private static string NaturalSortKey(string input)
        {
            return System.Text.RegularExpressions.Regex.Replace(input ?? "", @"\d+", m => m.Value.PadLeft(10, '0'));
        }

        // Returns the numeric value if the associated field is visible, otherwise 0.
        private double GetValueOrZero(string value, Visibility vis)
        {
            return vis == Visibility.Visible ? ParseDoubleSafe(value) : 0;
        }

        // Handles the Create button; gathers and converts all inputs and calls PlatesModule.CreatePlateFromUI.
        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double angleDeg = ParseDoubleSafe(Angle);

                // Always parse as double, pass 0 if hidden
                double l1_m = GetValueOrZero(L1, L1Visible) * 0.001;
                double l2_m = GetValueOrZero(L2, L2Visible) * 0.001;
                double c1 = GetValueOrZero(Count1, Count1Visible);
                double b1_m = GetValueOrZero(B1, B1Visible) * 0.001;
                double b2_m = GetValueOrZero(B2, B2Visible) * 0.001;
                double c2 = GetValueOrZero(Count2, Count2Visible);
                double t_m = GetValueOrZero(Thickness, TVisible) * 0.001;
                double r_m = GetValueOrZero(FilletRadius, RVisible) * 0.001;
                double d_m = GetValueOrZero(HoleDiameter, DVisible) * 0.001;

                PlatesModule.CreatePlateFromUI(
                    SelectedProfileType,
                    SelectedProfileSize,
                    angleDeg,
                    l1_m, l2_m, (int)c1,
                    b1_m, b2_m, (int)c2,
                    t_m, r_m, d_m,
                    insertPlateMid: InsertInMiddle
                );
            }
            catch (FormatException ex)
            {
                Application.ReportStatus($"Invalid input format: {ex.Message}\nPlease check all numeric fields.", StatusMessageType.Error, null);
            }
            catch (OverflowException ex)
            {
                Application.ReportStatus($"Input value is too large or too small: {ex.Message}", StatusMessageType.Error, null);
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Unexpected error: {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Adjusts labels and visibility of measurement fields depending on the selected plate type.
        private void UpdateFieldLabelsAndVisibility()
        {
            // Default: show all, standard labels
            L1Label = "l1"; L1Visible = Visibility.Visible;
            L2Label = "l2"; L2Visible = Visibility.Visible;
            Count1Label = "#"; Count1Visible = Visibility.Visible;
            B1Label = "b1"; B1Visible = Visibility.Visible;
            B2Label = "b2"; B2Visible = Visibility.Visible;
            Count2Label = "#"; Count2Visible = Visibility.Visible;
            DLabel = "d"; DVisible = Visibility.Visible;
            TLabel = "T"; TVisible = Visibility.Visible;
            RLabel = "r"; RVisible = Visibility.Visible;

            if (SelectedProfileType == "Blind flange")
            {
                L1Label = "D"; L1Visible = Visibility.Visible;
                L2Label = "l1"; L2Visible = Visibility.Visible;
                DLabel = "d2"; DVisible = Visibility.Visible;
                TLabel = "T"; TVisible = Visibility.Visible;
                Count1Label = "#"; Count1Visible = Visibility.Visible;
                // Hide all others
                B1Visible = B2Visible = Count2Visible = RVisible = Visibility.Collapsed;
            }
            else if (SelectedProfileType == "Flat flange")
            {
                L1Label = "D"; L1Visible = Visibility.Visible;
                L2Label = "l1"; L2Visible = Visibility.Visible;
                DLabel = "d2"; DVisible = Visibility.Visible;
                B1Label = "d3"; B1Visible = Visibility.Visible;
                TLabel = "T"; TVisible = Visibility.Visible;
                Count1Label = "#"; Count1Visible = Visibility.Visible;
                // Hide all others
                B2Visible = Count2Visible = RVisible = Visibility.Collapsed;
            }
            else if (SelectedProfileType == "UNP")
            {
                L1Label = "l1"; L1Visible = Visibility.Visible;
                B1Label = "b1"; B1Visible = Visibility.Visible;
                TLabel = "T"; TVisible = Visibility.Visible;
                RLabel = "r"; RVisible = Visibility.Visible;
                // Hide all others
                L2Visible = Count1Visible = B2Visible = Count2Visible = DVisible = Visibility.Collapsed;
            }

            // Notify UI
            OnPropertyChanged(nameof(L1Label)); OnPropertyChanged(nameof(L1Visible));
            OnPropertyChanged(nameof(L2Label)); OnPropertyChanged(nameof(L2Visible));
            OnPropertyChanged(nameof(Count1Label)); OnPropertyChanged(nameof(Count1Visible));
            OnPropertyChanged(nameof(B1Label)); OnPropertyChanged(nameof(B1Visible));
            OnPropertyChanged(nameof(B2Label)); OnPropertyChanged(nameof(B2Visible));
            OnPropertyChanged(nameof(Count2Label)); OnPropertyChanged(nameof(Count2Visible));
            OnPropertyChanged(nameof(DLabel)); OnPropertyChanged(nameof(DVisible));
            OnPropertyChanged(nameof(TLabel)); OnPropertyChanged(nameof(TVisible));
            OnPropertyChanged(nameof(RLabel)); OnPropertyChanged(nameof(RVisible));
        }

        public string L1Label { get; set; } = "L1";
        public Visibility L1Visible { get; set; } = Visibility.Visible;
        public string L2Label { get; set; } = "L2";
        public Visibility L2Visible { get; set; } = Visibility.Visible;
        public string Count1Label { get; set; } = "#";
        public Visibility Count1Visible { get; set; } = Visibility.Visible;
        public string B1Label { get; set; } = "B1";
        public Visibility B1Visible { get; set; } = Visibility.Visible;
        public string B2Label { get; set; } = "B2";
        public Visibility B2Visible { get; set; } = Visibility.Visible;
        public string Count2Label { get; set; } = "#";
        public Visibility Count2Visible { get; set; } = Visibility.Visible;
        public string DLabel { get; set; } = "D";
        public Visibility DVisible { get; set; } = Visibility.Visible;
        public string TLabel { get; set; } = "T";
        public Visibility TVisible { get; set; } = Visibility.Visible;
        public string RLabel { get; set; } = "R";
        public Visibility RVisible { get; set; } = Visibility.Visible;
    }
}
