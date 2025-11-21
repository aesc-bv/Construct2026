/*
 KruisRibCmd defines the SpaceClaim command capsule for the Rib Cut-Out tool.
 It hosts the RibCutOutControl in a dockable panel and exposes it as a ribbon/button command
 that is enabled when a SpaceClaim document window is active.
*/

using AESCConstruct25.UI;                         // for RibCutOutControl
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Extensibility;          // for CommandCapsule, ExecutionContext
using System.Drawing;                             // for Rectangle
using System.Windows.Forms;                       // for DockStyle
using System.Windows.Forms.Integration;           // for ElementHost
using Application = SpaceClaim.Api.V242.Application;
using Panel = SpaceClaim.Api.V242.Panel;

namespace AESCConstruct25.RibCutout.Commands
{
    class KruisRibCmd : CommandCapsule
    {
        public const string CommandName = "AESCConstruct25.RibCutOut.KruisRib";

        private RibCutOutControl _ribCutControl;
        private ElementHost _ribCutHost;

        // Initializes the command capsule with id, display text and tooltip for the Rib Cut-Out panel.
        public KruisRibCmd()
            : base(
                CommandName,
                "Rib Cut-Out",
                null,                     // no icon
                "Create cut-out ribs"
              )
        { }

        // Updates command state; enables the button only when there is an active SpaceClaim window.
        protected override void OnUpdate(Command command)
        {
            // enable only when a document is open
            command.IsEnabled = Window.ActiveWindow != null;
        }

        // Executes the command by creating (once) and showing the RibCutOutControl in a dockable panel.
        protected override void OnExecute(Command command, ExecutionContext context, Rectangle buttonRect)
        {
            var window = Window.ActiveWindow;
            if (window == null)
                return;

            if (_ribCutControl == null)
            {
                _ribCutControl = new RibCutOutControl();
                _ribCutHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _ribCutControl
                };
            }

            Application.AddPanelContent(
                command,
                _ribCutHost,
                Panel.Options
            );
        }
    }
}
