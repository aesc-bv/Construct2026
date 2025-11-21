/*
 EngravingControl is a WPF user control that lets the user configure text, font and placement
 for engravings or cut-outs, and delegates the actual operations to EngravingService.
*/

using AESCConstruct25.UIMain;
using SpaceClaim.Api.V242;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Application = SpaceClaim.Api.V242.Application;

namespace AESCConstruct25.UI
{
    public partial class EngravingControl : UserControl
    {
        // Initializes the engraving UI, sets default flags and font options, and binds the data context.
        public EngravingControl()
        {
            InitializeComponent();
            DataContext = this;
            Localization.Language.LocalizeFrameworkElement(this);

            UseComponentName = true;
            UseBodyName = false;
            UseCustomText = false;
            CustomText = "";
            Engraving = true;
            CutOut = false;

            FontOptions = new ObservableCollection<string>(
                System.Drawing.FontFamily.Families.Select(f => f.Name));
            SelectedFont = FontOptions.FirstOrDefault() ?? "Arial";
            Size = "5";
            Center = true;
        }

        public bool UseComponentName { get; set; }
        public bool UseBodyName { get; set; }
        public bool UseCustomText { get; set; }
        public string CustomText { get; set; }
        public bool Engraving { get; set; }
        public bool CutOut { get; set; }
        public ObservableCollection<string> FontOptions { get; }
        public string SelectedFont { get; set; }
        public string Size { get; set; }
        public bool Center { get; set; }

        // Handles “Add note” by parsing size and calling EngravingService.AddNote with current UI options.
        private void AddNoteButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double sz = double.Parse(
                    Size, System.Globalization.CultureInfo.InvariantCulture);

                EngravingService.AddNote(
                    sz,
                    Engraving,
                    UseCustomText,
                    CustomText,
                    CutOut,
                    SelectedFont,
                    UseBodyName,
                    Center);

                SpaceClaim.Api.V242.Application.ReportStatus(
                    "Note placed.",
                    StatusMessageType.Information,
                    null);
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Engraving Error {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Calls EngravingService to convert imprinted entities into engraving and export the result.
        private void ImprintToEngravingButton_Click(object sender, RoutedEventArgs e)
        {
            try { EngravingService.ImprintToEngravingAndExport(); }
            catch (Exception ex)
            {
                Application.ReportStatus($"Engraving Error {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Calls EngravingService to imprint the current engraving geometry onto a selected body.
        private void ImprintBodyButton_Click(object sender, RoutedEventArgs e)
        {
            try { EngravingService.ImprintBody(); }
            catch (Exception ex)
            {
                Application.ReportStatus($"Engraving Error {ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Uses EngravingService to imprint selected lines as engraving and export (same path as imprint-to-engraving).
        private void ImprintLinesButton_Click(object sender, RoutedEventArgs e)
        {
            try { EngravingService.ImprintToEngravingAndExport(); }
            catch (Exception ex)
            {
                Application.ReportStatus($"Engraving Error {ex.Message}", StatusMessageType.Error, null);
            }
        }
    }
}
