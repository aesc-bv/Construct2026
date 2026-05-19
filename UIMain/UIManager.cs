/*
 UIManager centralizes registration and localization of Construct2026 sidebar commands
 and manages their hosting as docked SpaceClaim panels.
*/

using AESCConstruct2026.Connector2.UI;
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
using WpfButton = System.Windows.Controls.Button;
using WpfDockPanel = System.Windows.Controls.DockPanel;
using WpfDock = System.Windows.Controls.Dock;



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
        private static Connector2Control _connectorControl;

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
                Localization.Language.Translate("UIManager_Hint_FrameGenerator"),
                Resources.FrameGen,
                () => Show(ProfileCommand),
                valid
            );

            Register(
                SettingsCommand,
                Localization.Language.Translate("Ribbon.Button.Settings"),
                Localization.Language.Translate("UIManager_Hint_Settings"),
                Resources.settings,
                () => Show(SettingsCommand),
                true
            );

            Register(
                PlateCommand,
                Localization.Language.Translate("Ribbon.Button.Plate"),
                Localization.Language.Translate("UIManager_Hint_Plate"),
                Resources.InsertPlate,
                () => Show(PlateCommand),
                valid
            );

            Register(
                FastenerCommand,
                Localization.Language.Translate("Ribbon.Button.Fastener"),
                Localization.Language.Translate("UIManager_Hint_Fastener"),
                Resources.Fasteners,
                () => Show(FastenerCommand),
                valid
            );

            Register(
                RibCutOutCommand,
                Localization.Language.Translate("Ribbon.Button.RibCutOut"),
                Localization.Language.Translate("UIManager_Hint_RibCutOut"),
                Resources.ribCutout,
                () => Show(RibCutOutCommand),
                valid
            );

            Register(
                CustomPropertiesCommand,
                Localization.Language.Translate("Ribbon.Button.CustomProperties"),
                Localization.Language.Translate("UIManager_Hint_CustomProperties"),
                Resources.Custom_Properties,
                () => Show(CustomPropertiesCommand),
                valid
            );

            Register(
                EngravingCommand,
                Localization.Language.Translate("Ribbon.Button.Engraving"),
                Localization.Language.Translate("UIManager_Hint_Engraving"),
                Resources.Engraving,
                () => Show(EngravingCommand),
                valid
            );

            Register(
                ConnectorCommand,
                Localization.Language.Translate("Ribbon.Button.Connector"),
                Localization.Language.Translate("UIManager_Hint_Connector"),
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
            try
            {
                var preExisting = Command.GetCommand(ConstructPanelCommand);
                Logger.Log($"[UIManager] RegisterConstructPanel start. _constructPanelCmd={(_constructPanelCmd == null ? "null" : "set")}, preExistingCommand={(preExisting == null ? "null" : "FOUND")}");

                _constructPanelCmd = Command.Create(ConstructPanelCommand);
                _constructPanelCmd.Text = Localization.Language.Translate("UIManager_Panel_Title");
                _constructPanelCmd.Hint = Localization.Language.Translate("UIManager_Panel_Hint");
                _constructPanelCmd.Image = LoadImage(Resources.FrameGen);
                _constructPanelCmd.IsEnabled = true;
                _constructPanelCmd.IsVisible = false;   // hidden until first use
                _constructPanelCmd.KeepAlive(true);

                _constructHost = new ElementHost { Dock = DockStyle.Fill };

                _constructPanelTab = PanelTab.Create(_constructPanelCmd, _constructHost, DockLocation.Right, 300, false);
                Logger.Log($"[UIManager] RegisterConstructPanel OK. PanelTab={(_constructPanelTab == null ? "null" : "created")}");
            }
            catch (Exception ex)
            {
                Logger.Log("[UIManager] RegisterConstructPanel FAILED: " + ex.ToString());
            }
        }

        // Ensures the Construct panel exists. Previously this recreated the PanelTab if
        // IsDeleted reported true, but that turned out to stack up duplicate tabs because
        // the SpaceClaim remoting proxy reports IsDeleted=true spuriously (the visual tab
        // is still alive in the UI, so creating another PanelTab adds a second orphan).
        // We now only (re)register if we've never created one in this AppDomain.
        private static void EnsureConstructPanel()
        {
            if (_constructPanelCmd == null)
            {
                Logger.Log("[UIManager] EnsureConstructPanel: _constructPanelCmd is null → RegisterConstructPanel");
                RegisterConstructPanel();
                return;
            }

            // Keep the panel command visible on the side strip.
            try { _constructPanelCmd.IsVisible = true; } catch { }

            // Diagnostic only — do NOT recreate based on this. IsDeleted is unreliable
            // (remoting proxy can claim deletion while the real tab persists).
            if (_constructPanelTab != null)
            {
                try
                {
                    if (_constructPanelTab.IsDeleted)
                        Logger.Log("[UIManager] EnsureConstructPanel: IsDeleted=true reported (ignored — not recreating)");
                }
                catch (Exception ex)
                {
                    Logger.Log("[UIManager] EnsureConstructPanel: IsDeleted threw " + ex.GetType().Name + " (ignored)");
                }
            }

            // If the host is gone for some reason (e.g. SpaceClaim disposed it), recreate
            // the host only — re-attach into the existing PanelTab slot rather than
            // creating a new PanelTab.
            if (_constructHost == null)
            {
                Logger.Log("[UIManager] EnsureConstructPanel: host is null → creating new ElementHost (keeping existing PanelTab)");
                _constructHost = new ElementHost { Dock = DockStyle.Fill };
                _activeDockedKey = null;
                ClearCachedControls();
            }
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
            try
            {
                Logger.Log($"[UIManager] ShowDocked('{key}') start");

                try { Command.Execute("AESC.Construct.SetMode3D"); }
                catch (Exception ex) { Logger.Log("[UIManager] SetMode3D non-fatal: " + ex.Message); }

                EnsureConstructPanel();
                Logger.Log($"[UIManager] After EnsureConstructPanel: panelCmd={(_constructPanelCmd == null ? "null" : "set")}, panelTab={(_constructPanelTab == null ? "null" : "set")}, host={(_constructHost == null ? "null" : "set")}");

                var control = GetDockedControl(key);
                if (control == null) { Logger.Log($"[UIManager] GetDockedControl returned null for '{key}'"); return; }

                if (_activeDockedKey != key)
                {
                    _constructHost.Child = null;          // clear old child first
                    _constructHost.Child = WrapWithCloseButton(control);
                    _activeDockedKey = key;
                }

                // Bring the panel tab to the front
                try
                {
                    if (_constructPanelTab != null && !_constructPanelTab.IsDeleted)
                        _constructPanelTab.Activate();
                }
                catch { /* proxy disconnected — panel was just recreated, will activate on next call */ }

                Logger.Log($"[UIManager] ShowDocked('{key}') complete");
            }
            catch (Exception ex)
            {
                Logger.Log("[UIManager] ShowDocked failed for '" + key + "': " + ex.ToString());
            }
        }

        // Wraps a control in a DockPanel with a close button at the top-right.
        private static WpfDockPanel WrapWithCloseButton(System.Windows.Controls.UserControl control)
        {
            // Detach from any previous wrapper parent
            if (System.Windows.LogicalTreeHelper.GetParent(control) is WpfDockPanel oldParent)
                oldParent.Children.Remove(control);

            var closeBtn = new WpfButton
            {
                Content = "\u2715",
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                VerticalAlignment = System.Windows.VerticalAlignment.Top,
                Width = 24,
                Height = 24,
                FontSize = 14,
                Padding = new System.Windows.Thickness(0),
                Background = System.Windows.Media.Brushes.Transparent,
                BorderBrush = System.Windows.Media.Brushes.Transparent,
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = Localization.Language.Translate("UIManager_Tooltip_ClosePanel")
            };
            closeBtn.Click += (s, e) => ClosePanel();

            WpfDockPanel.SetDock(closeBtn, WpfDock.Top);

            var panel = new WpfDockPanel();
            panel.Children.Add(closeBtn);
            panel.Children.Add(control);

            return panel;
        }

        // Closes the Construct panel and detaches the active control.
        public static void ClosePanel()
        {
            if (_constructHost != null)
                _constructHost.Child = null;
            _activeDockedKey = null;

            try
            {
                if (_constructPanelTab != null && !_constructPanelTab.IsDeleted)
                    _constructPanelTab.Close();
            }
            catch { /* proxy already disconnected */ }
        }

        // Called from Construct2026.Disconnect() when SpaceClaim is unloading our AppDomain.
        // Must close the PanelTab so the sidebar tab visual is removed from SpaceClaim's UI
        // before our AppDomain is disposed — otherwise the next AppDomain load creates a
        // second "AESC Construct" tab alongside the orphaned one.
        public static void ClosePanelAndDispose()
        {
            Logger.Log("[UIManager] ClosePanelAndDispose() called");

            try
            {
                if (_constructHost != null)
                    _constructHost.Child = null;
            }
            catch (Exception ex) { Logger.Log("[UIManager] Clear host child failed: " + ex.Message); }

            try
            {
                if (_constructPanelTab != null && !_constructPanelTab.IsDeleted)
                    _constructPanelTab.Close();
            }
            catch (Exception ex) { Logger.Log("[UIManager] PanelTab.Close failed: " + ex.Message); }

            try
            {
                _constructHost?.Dispose();
            }
            catch (Exception ex) { Logger.Log("[UIManager] Host dispose failed: " + ex.Message); }

            try
            {
                if (_constructPanelCmd != null)
                    _constructPanelCmd.IsVisible = false;
            }
            catch (Exception ex) { Logger.Log("[UIManager] Hide panel command failed: " + ex.Message); }

            _constructPanelTab = null;
            _constructHost = null;
            _activeDockedKey = null;

            Logger.Log("[UIManager] ClosePanelAndDispose() complete");
        }

        // Public entry point for KruisRibCmd to show the RibCutOut panel.
        public static void ShowRibCutOut()
        {
            Show(RibCutOutCommand);
        }

        // Re-applies the current language to every already-constructed module control.
        // Called after the user picks a new language in Settings so cached panels refresh
        // without requiring a SpaceClaim restart.
        public static void RelocalizeAll()
        {
            var controls = new System.Windows.Controls.UserControl[]
            {
                _profileControl, _settingsControl, _plateControl, _fastenerControl,
                _ribCutOutControl, _customPropertiesControl, _engravingControl, _connectorControl
            };

            foreach (var ctl in controls)
            {
                if (ctl == null) continue;
                try
                {
                    Localization.Language.LocalizeFrameworkElement(ctl);
                    // Refresh non-Tag strings (tooltips, watermarks, grid headers, code-built labels).
                    if (ctl is Localization.ILocalizable loc) loc.LocalizeUI();
                }
                catch (Exception ex) { Logger.Log("[UIManager] RelocalizeAll: " + ctl.GetType().Name + " failed: " + ex.Message); }
            }
        }

        // Sets a sidebar command's localized Text and Hint by id.
        private static void SetTextHint(string id, string textKey, string hintKey)
        {
            var c = Command.GetCommand(id);
            if (c == null) return;
            if (textKey != null) c.Text = Localization.Language.Translate(textKey);
            if (hintKey != null) c.Hint = Localization.Language.Translate(hintKey);
        }

        // Updates localized texts and hints for all sidebar commands and the Construct panel.
        public static void UpdateCommandTexts()
        {
            SetTextHint(ProfileCommand, "Ribbon.Button.FrameGenerator", "UIManager_Hint_FrameGenerator");
            SetTextHint(SettingsCommand, "Ribbon.Button.Settings", "UIManager_Hint_Settings");
            SetTextHint(PlateCommand, "Ribbon.Button.Plate", "UIManager_Hint_Plate");
            SetTextHint(FastenerCommand, "Ribbon.Button.Fastener", "UIManager_Hint_Fastener");
            SetTextHint(RibCutOutCommand, "Ribbon.Button.RibCutOut", "UIManager_Hint_RibCutOut");
            SetTextHint(CustomPropertiesCommand, "Ribbon.Button.CustomProperties", "UIManager_Hint_CustomProperties");
            SetTextHint(EngravingCommand, "Ribbon.Button.Engraving", "UIManager_Hint_Engraving");
            SetTextHint(ConnectorCommand, "Ribbon.Button.Connector", "UIManager_Hint_Connector");
            SetTextHint(ConstructPanelCommand, "UIManager_Panel_Title", "UIManager_Panel_Hint");
        }

        // Nulls all cached controls so they are freshly constructed for a new ElementHost.
        private static void ClearCachedControls()
        {
            _profileControl = null;
            _settingsControl = null;
            _plateControl = null;
            _fastenerControl = null;
            _ribCutOutControl = null;
            _customPropertiesControl = null;
            _engravingControl = null;
            _connectorControl = null;
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
        private static void EnsureConnector() { if (_connectorControl == null) _connectorControl = new Connector2Control(); }

        // Converts raw icon bytes to a System.Drawing.Image used as SpaceClaim command icon.
        private static Image LoadImage(byte[] bytes)
        {
            try { return new Bitmap(new MemoryStream(bytes)); }
            catch (Exception ex) { Logger.Log("[UIManager] LoadImage failed: " + ex.Message); return null; }
        }
    }
}
