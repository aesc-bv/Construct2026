/*
 UIManager centralizes registration and localization of Construct2026 sidebar commands
 and manages their hosting as docked SpaceClaim panels.
*/

using AESCConstruct2026.FrameGenerator.UI;
using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Licensing;
using AESCConstruct2026.UI;
using SpaceClaim.Api.V242;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.Integration;       // for ElementHost
using Application = SpaceClaim.Api.V242.Application;
using Image = System.Drawing.Image;



namespace AESCConstruct2026.UIMain
{
    public static class UIManager
    {
        // Construct panel
        private const string ConstructPanelCommand = "AESCConstruct2026.ConstructPanel";
        private static Command _constructPanelCmd;
        private static PanelTab _constructPanelTab;
        private static ElementHost _constructHost;
        private static string _activeDockedKey;   // tracks which module is currently shown

        // Docked controls
        private static ProfileSelectionControl _profileControl;
        private static SettingsControl _settingsControl;
        private static PlatesControl _plateControl;
        private static FastenersControl _fastenerControl;
        private static RibCutOutControl _ribCutOutControl;
        private static CustomComponentControl _customPropertiesControl;
        private static EngravingControl _engravingControl;
        private static ConnectorControl _connectorControl;

        // Command names
        public const string ProfileCommand = "AESCConstruct2026.ProfileSidebar";
        public const string SettingsCommand = "AESCConstruct2026.SettingsSidebar";
        public const string PlateCommand = "AESCConstruct2026.Plate";
        public const string FastenerCommand = "AESCConstruct2026.Fastener";
        public const string RibCutOutCommand = "AESCConstruct2026.RibCutOut";
        public const string CustomPropertiesCommand = "AESCConstruct2026.CustomProperties";
        public const string EngravingCommand = "AESCConstruct2026.EngravingControl";
        public const string ConnectorCommand = "AESCConstruct2026.ConnectorSidebar";

        // Registers all sidebar commands with SpaceClaim and wires them to UI handlers.
        public static void RegisterAll()
        {
            bool valid = ConstructLicenseSpot.IsValid;

            Register(
                ProfileCommand,
                Localization.Language.Translate("Ribbon.Button.FrameGenerator"),
                "Open the profile selection sidebar",
                Resources.FrameGen,
                () => Show(ProfileCommand),
                valid
            );

            Register(
                SettingsCommand,
                Localization.Language.Translate("Ribbon.Button.Settings"),
                "Open the settings sidebar",
                Resources.settings,
                () => Show(SettingsCommand),
                true
            );

            Register(
                PlateCommand,
                Localization.Language.Translate("Ribbon.Button.Plate"),
                "Open the plate-creation pane",
                Resources.InsertPlate,
                () => Show(PlateCommand),
                valid
            );

            Register(
                FastenerCommand,
                Localization.Language.Translate("Ribbon.Button.Fastener"),
                "Open the fastener insertion pane",
                Resources.Fasteners,
                () => Show(FastenerCommand),
                valid
            );

            Register(
                RibCutOutCommand,
                Localization.Language.Translate("Ribbon.Button.RibCutOut"),
                "Open the rib cut-out pane",
                Resources.ribCutout,
                () => Show(RibCutOutCommand),
                valid
            );

            Register(
                CustomPropertiesCommand,
                Localization.Language.Translate("Ribbon.Button.CustomProperties"),
                "Open the custom-properties pane",
                Resources.Custom_Properties,
                () => Show(CustomPropertiesCommand),
                valid
            );

            Register(
                EngravingCommand,
                Localization.Language.Translate("Ribbon.Button.Engraving"),
                "Open the engraving pane",
                Resources.Engraving,
                () => Show(EngravingCommand),
                valid
            );

            Register(
                ConnectorCommand,
                Localization.Language.Translate("Ribbon.Button.Connector"),
                "Open the connector pane",
                Resources.Menu_Connector,
                () => Show(ConnectorCommand),
                valid
            );

            RegisterConstructPanel();
        }

        // Refreshes sidebar command enabled state based on license validity and updates texts.
        public static void RefreshLicenseUI()
        {
            bool valid = ConstructLicenseSpot.IsValid;

            SetEnabled(ProfileCommand, valid);
            SetEnabled(PlateCommand, valid);
            SetEnabled(FastenerCommand, valid);
            SetEnabled(RibCutOutCommand, valid);
            SetEnabled(CustomPropertiesCommand, valid);
            SetEnabled(EngravingCommand, valid);
            SetEnabled(ConnectorCommand, valid);

            UpdateCommandTexts();
        }

        // Sets the IsEnabled property of a SpaceClaim command by id.
        private static void SetEnabled(string commandId, bool enabled)
        {
            var c = Command.GetCommand(commandId);
            if (c != null) c.IsEnabled = enabled;
        }

        // Creates and configures a SpaceClaim command including icon, text, and execute delegate.
        private static void Register(string name, string text, string hint, byte[] icon, Action execute, bool enabled)
        {
            var cmd = Command.Create(name);
            cmd.Text = text;
            cmd.Hint = hint;
            cmd.Image = LoadImage(icon);
            cmd.IsEnabled = enabled;
            cmd.Executing += (s, e) => execute();
            cmd.KeepAlive(true);
        }

        // Shows the requested panel in the docked Construct panel.
        private static void Show(string panelKey)
        {
            ShowDocked(panelKey);
        }

        // Creates the dedicated Construct panel tab docked on the right side.
        private static void RegisterConstructPanel()
        {
            _constructPanelCmd = Command.Create(ConstructPanelCommand);
            _constructPanelCmd.Text = "AESC Construct";
            _constructPanelCmd.Hint = "AESC Construct tools panel";
            _constructPanelCmd.Image = LoadImage(Resources.FrameGen);
            _constructPanelCmd.IsEnabled = true;
            _constructPanelCmd.IsVisible = false;   // hidden until first use
            _constructPanelCmd.KeepAlive(true);

            _constructHost = new ElementHost { Dock = DockStyle.Fill };

            _constructPanelTab = PanelTab.Create(_constructPanelCmd, _constructHost, DockLocation.Right, 300, false);
        }

        // Re-creates the Construct panel if the user closed it.
        private static void EnsureConstructPanel()
        {
            if (_constructPanelCmd == null) { RegisterConstructPanel(); return; }

            if (_constructPanelTab == null || _constructPanelTab.IsDeleted)
            {
                if (_constructHost != null)
                    _constructHost.Child = null;      // detach WPF control before disposing
                _constructHost?.Dispose();
                _constructHost = new ElementHost { Dock = DockStyle.Fill };
                _constructPanelTab = PanelTab.Create(_constructPanelCmd, _constructHost, DockLocation.Right, 300, false);
                _activeDockedKey = null;
            }

            _constructPanelCmd.IsVisible = true;
        }

        // Returns the WPF control for a given command key, creating it on demand.
        private static System.Windows.Controls.UserControl GetDockedControl(string key)
        {
            switch (key)
            {
                case ProfileCommand:          EnsureProfile();           return _profileControl;
                case SettingsCommand:         EnsureSettings();          return _settingsControl;
                case PlateCommand:            EnsurePlate();             return _plateControl;
                case FastenerCommand:         EnsureFastener();          return _fastenerControl;
                case RibCutOutCommand:        EnsureRibCutOut();         return _ribCutOutControl;
                case CustomPropertiesCommand: EnsureCustomProperties();  return _customPropertiesControl;
                case EngravingCommand:        EnsureEngraving();         return _engravingControl;
                case ConnectorCommand:        EnsureConnector();         return _connectorControl;
                default: return null;
            }
        }

        // Displays the requested module in the dedicated Construct panel.
        private static void ShowDocked(string key)
        {
            Command.Execute("AESC.Construct.SetMode3D");
            EnsureConstructPanel();

            var control = GetDockedControl(key);
            if (control == null) return;

            if (_activeDockedKey != key)
            {
                _constructHost.Child = null;          // clear old child first
                _constructHost.Child = control;
                _activeDockedKey = key;
            }

            // Bring the panel tab to the front
            if (_constructPanelTab != null && !_constructPanelTab.IsDeleted)
                _constructPanelTab.Activate();
        }

        // Public entry point for KruisRibCmd to show the RibCutOut panel.
        public static void ShowRibCutOut()
        {
            Show(RibCutOutCommand);
        }

        // Updates localized texts for all sidebar commands.
        public static void UpdateCommandTexts()
        {
            // Sidebar buttons
            var c = Command.GetCommand(ProfileCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.FrameGenerator");

            c = Command.GetCommand(SettingsCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Settings");

            c = Command.GetCommand(PlateCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Plate");

            c = Command.GetCommand(FastenerCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Fastener");

            c = Command.GetCommand(RibCutOutCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.RibCutOut");

            c = Command.GetCommand(CustomPropertiesCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.CustomProperties");

            c = Command.GetCommand(EngravingCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Engraving");

            c = Command.GetCommand(ConnectorCommand);
            if (c != null) c.Text = Localization.Language.Translate("Ribbon.Button.Connector");
        }

        // Lazily creates the Profile sidebar control.
        private static void EnsureProfile() { if (_profileControl == null) _profileControl = new ProfileSelectionControl(); }

        // Lazily creates the Settings sidebar control.
        private static void EnsureSettings() { if (_settingsControl == null) _settingsControl = new SettingsControl(); }

        // Lazily creates the Plate sidebar control.
        private static void EnsurePlate() { if (_plateControl == null) _plateControl = new PlatesControl(); }

        // Lazily creates the Fastener sidebar control.
        private static void EnsureFastener() { if (_fastenerControl == null) _fastenerControl = new FastenersControl(); }

        // Lazily creates the RibCutOut sidebar control.
        private static void EnsureRibCutOut() { if (_ribCutOutControl == null) _ribCutOutControl = new RibCutOutControl(); }

        // Lazily creates the CustomProperties sidebar control.
        private static void EnsureCustomProperties() { if (_customPropertiesControl == null) _customPropertiesControl = new CustomComponentControl(); }

        // Lazily creates the Engraving sidebar control.
        private static void EnsureEngraving() { if (_engravingControl == null) _engravingControl = new EngravingControl(); }

        // Lazily creates the Connector sidebar control.
        private static void EnsureConnector() { if (_connectorControl == null) _connectorControl = new ConnectorControl(); }

        // Converts raw icon bytes to a System.Drawing.Image used as SpaceClaim command icon.
        private static Image LoadImage(byte[] bytes)
        {
            try { return new Bitmap(new MemoryStream(bytes)); }
            catch (Exception ex) { Logger.Log("[UIManager] LoadImage failed: " + ex.Message); return null; }
        }
    }
}
