using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SpaceClaim.Api.V251;
using Application = SpaceClaim.Api.V251.Application;

namespace AESCConstruct25.UI
{
    public partial class EngravingControl : UserControl
    {
        public EngravingControl()
        {
            InitializeComponent();
            DataContext = this;

            // default radio + text
            UseComponentName = true;
            UseBodyName = false;
            UseCustomText = false;
            CustomText = "";

            // engraving vs cutout
            Engraving = true;
            CutOut = false;

            // fonts
            FontOptions = new ObservableCollection<string>
            {
                "1CamBam_Stick_1",
                "Arial",
                "Helvetica"
            };
            SelectedFont = FontOptions.FirstOrDefault();
            Size = "5";

            Center = false;
        }

        // Radio button backing properties
        public bool UseComponentName { get; set; }
        public bool UseBodyName { get; set; }
        public bool UseCustomText { get; set; }
        public string CustomText { get; set; }

        // Mode
        public bool Engraving { get; set; }
        public bool CutOut { get; set; }

        // Font + size
        public ObservableCollection<string> FontOptions { get; }
        public string SelectedFont { get; set; }
        public string Size { get; set; }

        // Center flag
        public bool Center { get; set; }

        private void AddNoteButton_Click(object sender, RoutedEventArgs e)
        {
            // e.g. ask user to pick a location
            Application.ReportStatus(
                "Click on the model to place the note.",
                StatusMessageType.Information,
                null);
        }

        private void ImprintBodyButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                "Imprinting text onto the body...",
                StatusMessageType.Information,
                null);
        }

        private void ImprintLinesButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                "Imprinting guide lines...",
                StatusMessageType.Information,
                null);
        }

        private void ImprintToEngravingButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                "Finalizing engraving outline...",
                StatusMessageType.Information,
                null);
        }
    }
}
