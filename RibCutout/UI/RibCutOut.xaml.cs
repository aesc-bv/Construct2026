using AESCConstruct25.FrameGenerator.Utilities;     // RibCutOutSelectionHelper
using SpaceClaim.Api.V242;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct25.UI
{
    public partial class RibCutOutControl : UserControl
    {
        public RibCutOutControl()
        {
            InitializeComponent();
            DataContext = this;
            Localization.Language.LocalizeFrameworkElement(this);
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.ActiveWindow;
            if (window == null) return;

            var doc = window.Document;

            // Get all design bodies in main part
            var designBodies = doc.MainPart.Bodies;
            var bodies = designBodies.Select(db => db.Shape).ToList();

            var pairs = RibCutOutSelectionHelper.GetOverlappingPairs(bodies);
            if (pairs.Count == 0)
            {
                MessageBox.Show(
                    "Please select two overlapping bodies.",
                    "Rib Cut-Out",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                return;
            }

            // UI options
            bool perpendicular = PerpendicularCut.IsChecked == true;
            bool applyMiddleTolerance = MiddleTolerance.IsChecked == true;

            // New options
            bool reverseDirection = ReverseDirection != null && ReverseDirection.IsChecked == true;
            bool addWeldRound = AddWeldRound != null && AddWeldRound.IsChecked == true;

            // Parse tolerance (mm)
            if (!double.TryParse(
                    ToleranceInput.Text,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out double toleranceMM))
            {
                toleranceMM = 0.0;
            }

            // Parse weld radius (mm)
            double weldRadiusMM = 0.0;
            if (addWeldRound && RadiusInput != null)
            {
                double.TryParse(
                    RadiusInput.Text,
                    NumberStyles.Any,
                    CultureInfo.InvariantCulture,
                    out weldRadiusMM);
                if (weldRadiusMM < 0) weldRadiusMM = 0;
            }

            WriteBlock.ExecuteTask("Rib Cut-Out", () =>
            {
                AESCConstruct25.RibCutout.Modules.RibCutOutModule.ProcessPairs(
                    doc,
                    pairs,
                    perpendicular,
                    toleranceMM,
                    applyMiddleTolerance,
                    reverseDirection,
                    addWeldRound,
                    weldRadiusMM
                );
            });
        }
    }
}
