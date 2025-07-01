using System.Windows;
using System.Windows.Controls;
using SpaceClaim.Api.V251;
using Application = SpaceClaim.Api.V251.Application;

namespace AESCConstruct25.UI
{
    public partial class RibCutOut : UserControl
    {
        public RibCutOut()
        {
            InitializeComponent();
            DataContext = this;

            // defaults
            Tolerance = "0.5";
            ReverseDirection = false;
            MiddleTolerance = false;
            PerpendicularCut = true;

            AddWeldRound = false;
            WeldRoundRadius = "10";
            WeldRoundReverse = false;
        }

        // --- Cutout properties ---
        public string Tolerance { get; set; }
        public bool ReverseDirection { get; set; }
        public bool MiddleTolerance { get; set; }
        public bool PerpendicularCut { get; set; }

        // --- Weld-round properties ---
        public bool AddWeldRound { get; set; }
        public string WeldRoundRadius { get; set; }
        public bool WeldRoundReverse { get; set; }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            // TODO: call your SpaceClaim API with these parameters
            Application.ReportStatus(
                $"Cutout: tol={Tolerance}, rev={ReverseDirection}, mid={MiddleTolerance}, perp={PerpendicularCut}\n" +
                $"Weld round: add={AddWeldRound}, r={WeldRoundRadius}, rev={WeldRoundReverse}",
                StatusMessageType.Information,
                null);
        }
    }
}
