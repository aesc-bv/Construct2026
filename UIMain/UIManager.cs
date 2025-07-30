using AESCConstruct25.FrameGenerator.UI;
using AESCConstruct25.FrameGenerator.Utilities;
using AESCConstruct25.UI;
using SpaceClaim.Api.V242;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.Integration;       // for ElementHost
using Image = System.Drawing.Image;
using WinForms = System.Windows.Forms;        // alias WinForms namespace

namespace AESCConstruct25.UIMain
{
    public static class UIManager
    {

        static PanelTab myPanelTab;
        // Command names
        public const string ProfileCommand = "AESCConstruct25.ProfileSidebar";
        public const string SettingsCommand = "AESCConstruct25.SettingsSidebar";
        public const string PlateCommand = "AESCConstruct25.Plate";
        public const string FastenerCommand = "AESCConstruct25.Fastener";
        public const string RibCutOutCommand = "AESCConstruct25.RibCutOut";
        public const string CustomPropertiesCommand = "AESCConstruct25.CustomProperties";
        public const string EngravingCommand = "AESCConstruct25.EngravingControl";

        // Shared ElementHosts for WPF controls
        private static ElementHost _profileHost, _settingsHost,
                                  _plateHost, _fastenerHost, _ribCutOutHost, _customPropertiesHost, _engravingHost;

        // WPF controls (instances of UserControl)
        private static ProfileSelectionControl _profileControl;
        private static SettingsControl _settingsControl;
        private static PlatesControl _plateControl;
        private static FastenersControl _fastenerControl;
        private static RibCutOutControl _ribCutOutControl;
        private static CustomComponentControl _customPropertiesControl;
        private static EngravingControl _engravingControl;

        public static void RegisterAll()
        {
            bool valid = DateTime.Now < new DateTime(2025, 8, 31, 23, 59, 59);
            // Profile
            var profileCmd = Command.Create(ProfileCommand);
            profileCmd.Text = "Profile Selector";
            profileCmd.Hint = "Open the profile selection sidebar";
            profileCmd.Image = LoadImage(Resources.FrameGen);
            profileCmd.IsEnabled = valid;// LicenseSpot.LicenseSpot.State.Valid;
            profileCmd.Executing += (s, e) => ShowProfile();
            profileCmd.KeepAlive(true);

            // Settings
            var settingsCmd = Command.Create(SettingsCommand);
            settingsCmd.Text = "Settings";
            settingsCmd.Hint = "Open the settings sidebar";
            settingsCmd.Image = LoadImage(Resources.settings);
            settingsCmd.IsEnabled = valid;// LicenseSpot.LicenseSpot.State.Valid;
            settingsCmd.Executing += (s, e) => ShowSettings();
            settingsCmd.KeepAlive(true);

            // Plate
            var plateCmd = Command.Create(PlateCommand);
            plateCmd.Text = "Plate";
            plateCmd.Hint = "Open the plate‐creation pane";
            plateCmd.Image = LoadImage(Resources.InsertPlate);
            plateCmd.IsEnabled = valid;// LicenseSpot.LicenseSpot.State.Valid;
            plateCmd.Executing += (s, e) => ShowPlate();
            plateCmd.KeepAlive(true);

            // Fastener
            var fastenerCmd = Command.Create(FastenerCommand);
            fastenerCmd.Text = "Fastener";
            fastenerCmd.Hint = "Open the fastener insertion pane";
            fastenerCmd.Image = LoadImage(Resources.Fasteners);
            fastenerCmd.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
            fastenerCmd.Executing += (s, e) => ShowFastener();
            fastenerCmd.KeepAlive(true);

            // Rib CutOut
            var ribCmd = Command.Create(RibCutOutCommand);
            ribCmd.Text = "Rib Cutout";
            ribCmd.Hint = "Open the rib cutout pane";
            ribCmd.Image = LoadImage(Resources.ribCutout);
            ribCmd.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
            ribCmd.Executing += (s, e) => ShowRibCutOut();
            ribCmd.KeepAlive(true);

            var customPropCmd = Command.Create(CustomPropertiesCommand);
            customPropCmd.Text = "Custom Properties";
            customPropCmd.Hint = "Open the custom‐properties pane";
            customPropCmd.Image = LoadImage(Resources.Custom_Properties);
            customPropCmd.IsEnabled = valid;  // or LicenseSpot.Valid
            customPropCmd.Executing += (s, e) => ShowCustomProperties();
            customPropCmd.KeepAlive(true);

            var engravingCmd = Command.Create(EngravingCommand);
            engravingCmd.Text = "Engraving";
            engravingCmd.Hint = "Open the engraving pane";
            engravingCmd.Image = LoadImage(Resources.Engraving); // if you have an icon
            engravingCmd.IsEnabled = valid;
            engravingCmd.Executing += (s, e) => ShowEngraving();
            engravingCmd.KeepAlive(true);
        }

        private static void ShowEngraving()
        {
            Logger.Log("showengraving");
            if (_engravingControl == null)
            {
                Logger.Log("showengraving null");
                try
                {
                    _engravingControl = new EngravingControl();
                    Logger.Log("EngravingControl ctor succeeded");
                }
                catch (Exception ex)
                {
                    Logger.Log($"EngravingControl ctor failed: {ex}");
                    WinForms.MessageBox.Show(
                        "Could not initialize the Engraving panel:\n" + ex.Message,
                        "Engraving", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error
                    );
                    return;
                }

                _engravingHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _engravingControl
                };
                Logger.Log("_engravingHost created");
            }
            Logger.Log("showengraving2");

            var cmd = Command.GetCommand(EngravingCommand);
            Logger.Log("GetCommand");
            SpaceClaim.Api.V242.Application.AddPanelContent(
                cmd,
                _engravingHost,
                SpaceClaim.Api.V242.Panel.Options
            );
            Logger.Log("AddPanelContent");
        }

        private static void ShowCustomProperties()
        {
            if (_customPropertiesControl == null)
            {
                // your CustomPropertiesControl is a WPF UserControl
                _customPropertiesControl = new CustomComponentControl();
                _customPropertiesHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = _customPropertiesControl
                };
            }

            var cmd = Command.GetCommand(CustomPropertiesCommand);
            SpaceClaim.Api.V242.Application.AddPanelContent(
                cmd,
                _customPropertiesHost,
                SpaceClaim.Api.V242.Panel.Options
            );
        }

        private static void ShowProfile()
        {
            if (_profileControl == null)
            {
                _profileControl = new ProfileSelectionControl();
                // No TopLevel on WPF UserControl :contentReference[oaicite:5]{index=5}
                _profileHost = new ElementHost { Dock = WinForms.DockStyle.Fill, Child = _profileControl };  // host WPF :contentReference[oaicite:6]{index=6}
            }
            var cmd = Command.GetCommand(ProfileCommand);
            SpaceClaim.Api.V242.Application.AddPanelContent(   // fully qualified Application :contentReference[oaicite:7]{index=7}
                cmd,
                _profileHost,
                SpaceClaim.Api.V242.Panel.Options             // fully qualified Panel enum :contentReference[oaicite:8]{index=8}
            );
        }

        private static void ShowSettings()
        {
            if (_settingsControl == null)
            {
                _settingsControl = new SettingsControl();
                _settingsHost = new ElementHost { Dock = WinForms.DockStyle.Fill, Child = _settingsControl };
            }
            var cmd = Command.GetCommand(SettingsCommand);
            SpaceClaim.Api.V242.Application.AddPanelContent(cmd, _settingsHost, SpaceClaim.Api.V242.Panel.Options);
        }

        private static void ShowPlate()
        {
            if (_plateControl == null)
            {
                _plateControl = new PlatesControl();
                _plateHost = new ElementHost { Dock = WinForms.DockStyle.Fill, Child = _plateControl };
            }
            var cmd = Command.GetCommand(PlateCommand);
            SpaceClaim.Api.V242.Application.AddPanelContent(cmd, _plateHost, SpaceClaim.Api.V242.Panel.Options);
        }

        private static void ShowFastener()
        {
            if (_fastenerControl == null)
            {
                try
                {
                    _fastenerControl = new FastenersControl();
                }
                catch (Exception ex)
                {
                    Logger.Log($"FastenersControl ctor failed: {ex}");
                    WinForms.MessageBox.Show(
                        "Could not initialize the Fastener panel:\n" + ex.Message,
                        "Fastener", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error
                    );
                    return;
                }
                _fastenerHost = new ElementHost { Dock = WinForms.DockStyle.Fill, Child = _fastenerControl };
            }
            var cmd = Command.GetCommand(FastenerCommand);
            SpaceClaim.Api.V242.Application.AddPanelContent(cmd, _fastenerHost, SpaceClaim.Api.V242.Panel.Options);
        }

        private static void ShowRibCutOut()
        {
            if (_ribCutOutControl == null)
            {
                _ribCutOutControl = new RibCutOutControl();
                _ribCutOutHost = new ElementHost { Dock = WinForms.DockStyle.Fill, Child = _ribCutOutControl };
            }
            var cmd = Command.GetCommand(RibCutOutCommand);
            SpaceClaim.Api.V242.Application.AddPanelContent(cmd, _ribCutOutHost, SpaceClaim.Api.V242.Panel.Options);
        }

        private static Image LoadImage(byte[] bytes)
        {
            try { return new Bitmap(new MemoryStream(bytes)); }
            catch { return null; }
        }
    }
}
