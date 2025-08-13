using AESCConstruct25.UIMain;
using SpaceClaim.Api.V242;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace AESCConstruct25.UI
{
    public partial class EngravingControl : UserControl
    {
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
                MessageBox.Show(
                    ex.Message,
                    "Engraving Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void ImprintToEngravingButton_Click(object sender, RoutedEventArgs e)
        {
            try { EngravingService.ImprintToEngravingAndExport(); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Engraving Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ImprintBodyButton_Click(object sender, RoutedEventArgs e)
        {
            try { EngravingService.ImprintBody(); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Engraving Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // If you still have an “Imprint Lines” button, add this:
        private void ImprintLinesButton_Click(object sender, RoutedEventArgs e)
        {
            // reuse ImprintToEngravingAndExport or implement separate logic
            try { EngravingService.ImprintToEngravingAndExport(); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Engraving Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
