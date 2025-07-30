using System;
using System.Drawing;                             // for Rectangle
using System.Windows.Forms;                       // for DockStyle
using System.Windows.Forms.Integration;           // for ElementHost
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Extensibility;          // for CommandCapsule, ExecutionContext
using Application = SpaceClaim.Api.V242.Application;
using Panel = SpaceClaim.Api.V242.Panel;
using AESCConstruct25.UI;                         // for RibCutOutControl

namespace AESCConstruct25.RibCutout.Commands
{
    class KruisRibCmd : CommandCapsule
    {
        public const string CommandName = "AESCConstruct25.RibCutOut.KruisRib";

        private RibCutOutControl _ribCutControl;
        private ElementHost _ribCutHost;

        public KruisRibCmd()
            : base(
                CommandName,
                "Rib Cut-Out",
                null,                     // no icon
                "Create cut-out ribs"
              )
        { }

        protected override void OnUpdate(Command command)
        {
            // enable only when a document is open
            command.IsEnabled = Window.ActiveWindow != null;
        }

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
