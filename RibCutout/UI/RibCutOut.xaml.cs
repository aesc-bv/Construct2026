/*
 RibCutOutControl is the WPF UI for creating rib cut-out (half-lap) joints.
 It reads selection and user options from the panel, then calls the RibCutOut module
 to create joints between overlapping bodies in the active SpaceClaim document.
*/

using AESCConstruct25.FrameGenerator.Utilities;     // RibCutOutSelectionHelper
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Modeler;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Application = SpaceClaim.Api.V242.Application;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct25.UI
{
    public partial class RibCutOutControl : UserControl
    {
        // Initializes the rib cut-out panel, sets the DataContext, and applies localization.
        public RibCutOutControl()
        {
            InitializeComponent();
            DataContext = this;
            Localization.Language.LocalizeFrameworkElement(this);
        }

        // Handles the Create button click: validates selection, reads options, and invokes the rib cut-out creation.
        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            var window = Window.ActiveWindow;
            if (window == null) return;

            var doc = window.Document;

            // Get all design bodies in main part
            //var designBodies = doc.MainPart.Bodies;

            var sel = window?.ActiveContext?.Selection;
            List<Body> bodies;
            if (sel != null && sel.Count > 0)
            {
                // Filter selected bodies
                bodies = sel
                    .Where(o => o is DesignBody)
                    .Select(o => (o as DesignBody).Shape)
                    .ToList();

                var pairs = RibCutOutSelectionHelper.GetOverlappingPairs(bodies);
                if (pairs.Count == 0)
                {
                    Application.ReportStatus("Please select two overlapping bodies.", StatusMessageType.Warning, null);
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
                    // Use dummy values for HalfLapParams for now
                    var halfLapParams = new RibCutout.Modules.RibCutOutModule.HalfLapParams
                    {
                        Width = 0,
                        Length = 0,
                        Thickness = 0,
                        Depth = 0,
                        Shoulder = 0,
                        Check = 0
                    };

                    RibCutout.Modules.RibCutOutModule.CreateHalfLapJoints(
                        doc,
                        pairs,
                        halfLapParams,
                        perpendicular,
                        reverseDirection,
                        toleranceMM,
                        applyMiddleTolerance,
                        addWeldRound,
                        weldRadiusMM
                    );
                });
            }
        }
    }
}
