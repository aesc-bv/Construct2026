/*
 KruisRibCmd defines the SpaceClaim command capsule for the Rib Cut-Out tool.
 It routes the panel display through UIManager so it appears in the dedicated
 Construct panel instead of creating its own AddPanelContent call.
*/

using AESCConstruct2026.UIMain;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Extensibility;          // for CommandCapsule, ExecutionContext
using System.Drawing;                             // for Rectangle

namespace AESCConstruct2026.RibCutout.Commands
{
    class KruisRibCmd : CommandCapsule
    {
        public const string CommandName = "AESCConstruct2026.RibCutOut.KruisRib";

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

        // Executes the command by showing the RibCutOut panel through UIManager.
        protected override void OnExecute(Command command, ExecutionContext context, Rectangle buttonRect)
        {
            if (Window.ActiveWindow == null) return;
            UIManager.ShowRibCutOut();
        }
    }
}
