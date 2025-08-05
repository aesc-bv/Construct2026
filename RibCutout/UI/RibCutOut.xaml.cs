using AESCConstruct25.FrameGenerator.Utilities;     // RibCutOutSelectionHelper
using SpaceClaim.Api.V242;
using System.Linq;                                  // ← add this
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
            Logger.Log("RibCutOut: control initialized.");
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            Logger.Log("RibCutOut: CreateButton_Click start.");

            var window = Window.ActiveWindow;
            if (window == null) return;

            var doc = window.Document;

            var designBodies = window.Document.MainPart.Bodies;
            var bodies = designBodies
                                  .Select(db => db.Shape)
                                  .ToList();

            var pairs = RibCutOutSelectionHelper.GetOverlappingPairs(bodies);
            Logger.Log($"RibCutOut: found {pairs.Count} overlapping pairs.");

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

            bool perpendicular = PerpendicularCut.IsChecked == true;
            bool applyMiddleTolerance = MiddleTolerance.IsChecked == true;

            Logger.Log($"RibCutOut: perpendicularCut={perpendicular}.");

            if (!double.TryParse(
                    ToleranceInput.Text,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out double toleranceMM))
            {
                Logger.Log($"RibCutOut: Invalid tolerance input '{ToleranceInput.Text}'. Defaulting to 0.0.");
                toleranceMM = 0.0;
            }

            WriteBlock.ExecuteTask("Rib Cut-Out", () =>
            {
                Logger.Log("RibCutOut: WriteBlock start.");
                RibCutout.Modules.RibCutOutModule.ProcessPairs(doc, pairs, perpendicular, toleranceMM, applyMiddleTolerance);
                Logger.Log("RibCutOut: WriteBlock end.");
            });
        }
    }
}
