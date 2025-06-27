using System;
using System.Windows;
using System.Windows.Controls;
using AESCConstruct25.Commands;
using AESCConstruct25.Utilities;
using SpaceClaim.Api.V251;           // for Window.ActiveWindow
using SpaceClaim.Api.V251.Extensibility;
using Application = SpaceClaim.Api.V251.Application;
using Window = SpaceClaim.Api.V251.Window; // for WriteBlock

namespace AESCConstruct25.UI
{
    public partial class JointSelectionControl : UserControl
    {
        public JointSelectionControl()
        {
            InitializeComponent();
        }

        private void GenerateJoint_Click(object sender, RoutedEventArgs e)
        {
            var oldOri = Application.UserOptions.WorldOrientation;
            Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
            try
            {
                var window = Window.ActiveWindow;
                if (window == null)
                    return;

                // parse gap (mm → m)
                double spacing = 0;
                if (double.TryParse(Gap.Text, out var mm))
                    spacing = mm / 1000.0;

                // pick joint type
                string jointType =
                    MiterJoint.IsChecked == true ? "Miter" :
                    StraightJoint.IsChecked == true ? "Straight" :
                    StraightJoint2.IsChecked == true ? "Straight2" :
                    TJoint.IsChecked == true ? "T" :
                    CutOut.IsChecked == true ? "CutOut" :
                    Trim.IsChecked == true ? "Trim" :
                    "None"
                ;

                // do it inside a write‐block
                WriteBlock.ExecuteTask("Execute Joint", () => {
                    ExecuteJointCommand.ExecuteJoint(window, spacing, jointType);
                });

                // close the sidebar
                //JointSidebar.CloseSidebar();
            }
            finally
            {
                Application.UserOptions.WorldOrientation = oldOri;
            }
        }

        private void RestoreGeometry_Click(object sender, RoutedEventArgs e)
        {
            var oldOri = Application.UserOptions.WorldOrientation;
            Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
            try
            {
                var window = Window.ActiveWindow;
                if (window == null)
                    return;

                WriteBlock.ExecuteTask("Restore Geometry", () => {
                    ExecuteJointCommand.RestoreGeometry(window);
                });
            }
            finally
            {
                Application.UserOptions.WorldOrientation = oldOri;
            }
        }

        private void RestoreJoint_Click(object sender, RoutedEventArgs e)
        {
            var oldOri = Application.UserOptions.WorldOrientation;
            Application.UserOptions.WorldOrientation = WorldOrientation.UpIsY;
            try
            {
                var window = Window.ActiveWindow;
                if (window == null)
                    return;

                WriteBlock.ExecuteTask("Restore Joint", () => {
                    ExecuteJointCommand.RestoreJoint(window);
                });
            }
            finally
            {
                Application.UserOptions.WorldOrientation = oldOri;
            }
        }
    }
}
