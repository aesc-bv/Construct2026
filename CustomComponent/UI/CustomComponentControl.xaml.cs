using SpaceClaim.Api.V242;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Application = SpaceClaim.Api.V242.Application;

namespace AESCConstruct25.UI
{
    public partial class CustomComponentControl : UserControl
    {
        public CustomComponentControl()
        {
            InitializeComponent();
            DataContext = this;

            // Example template list; replace with your actual CSV names
            TemplateOptions = new ObservableCollection<string>
            {
                "CompProperties.csv",
                "OtherTemplate.csv"
            };
            SelectedTemplate = TemplateOptions.FirstOrDefault();
        }

        // Templates dropdown
        public ObservableCollection<string> TemplateOptions { get; }
        public string SelectedTemplate { get; set; }

        // --- Button handlers ---
        private void AddDocPropsButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                $"Adding document properties from {SelectedTemplate}",
                StatusMessageType.Information, null);
            // TODO: load CSV and add to document
        }

        private void DeleteDocPropsButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                $"Deleting document properties defined in {SelectedTemplate}",
                StatusMessageType.Information, null);
            // TODO: delete those properties
        }

        private void AddToSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                $"Adding component properties to current selection from {SelectedTemplate}",
                StatusMessageType.Information, null);
            // TODO: apply to selected components
        }

        private void DeleteFromSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                $"Deleting component properties from current selection as per {SelectedTemplate}",
                StatusMessageType.Information, null);
            // TODO: remove from selected components
        }

        private void AddToAllButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                $"Adding component properties to all components using {SelectedTemplate}",
                StatusMessageType.Information, null);
            // TODO: apply to every component
        }

        private void DeleteFromAllButton_Click(object sender, RoutedEventArgs e)
        {
            Application.ReportStatus(
                $"Deleting component properties from all components as per {SelectedTemplate}",
                StatusMessageType.Information, null);
            // TODO: remove from every component
        }
    }
}
