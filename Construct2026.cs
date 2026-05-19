/*
 Construct2026 SpaceClaim add-in entry point.
 Initializes the SpaceClaim API and licensing, registers all Construct2026 ribbon commands,
 and provides helpers for DXF import/export, BOM/STEP/Excel export, and license UI state.
*/

using AESCConstruct2026.ClashDetection;
using AESCConstruct2026.FrameGenerator.Commands;
using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Licensing;
using AESCConstruct2026.UIMain;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Extensibility;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Application = SpaceClaim.Api.V242.Application;
using Clipboard = System.Windows.Forms.Clipboard;
using Image = System.Drawing.Image;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct2026
{
    [Serializable]
    public class Construct2026 : MarshalByRefObject, IExtensibility, ICommandExtensibility, IRibbonExtensibility
    {
        private static bool isCommandRegistered = false;

        // Early entry point: initializes the SpaceClaim API and returns true so the add-in is marked Active.
        // Session-dependent work (commands, licensing) is deferred to Initialize().
        public bool Connect()
        {
            try
            {
                Api.Initialize();
                var apiAsm = typeof(Application).Assembly;
                Logger.Log($"API 1.1 Initialized — {apiAsm.GetName().Version} from {apiAsm.Location}");

                // Fallback cleanup: SpaceClaim does not always call Disconnect() when it
                // reloads our addin in a new AppDomain during a running session. Hooking
                // DomainUnload gives us a second chance to close the PanelTab so the bar
                // entry is removed from barlayout2.xml, preventing orphan sidebar tabs.
                AppDomain.CurrentDomain.DomainUnload += OnDomainUnload;
                AppDomain.CurrentDomain.ProcessExit += OnDomainUnload;

                return true;
            }
            catch (Exception ex)
            {
                Logger.Log("Connect() failed: " + ex.Message);
                return false;
            }
        }

        private static void OnDomainUnload(object sender, EventArgs e)
        {
            try
            {
                Logger.Log("[Construct2026] DomainUnload fired — running panel cleanup");
                UIManager.ClosePanelAndDispose();
            }
            catch (Exception ex)
            {
                try { Logger.Log("[Construct2026] DomainUnload cleanup failed: " + ex.Message); } catch { }
            }
        }

        // Called by SpaceClaim after sessions are available. Registers all commands and sets up licensing.
        public void Initialize()
        {
            try
            {
                var session = Session.GetSessions().FirstOrDefault();
                if (session == null)
                {
                    Logger.Log("No active SpaceClaim session found.");
                    return;
                }
                Api.AttachToSession(session);

                // Check license, but don't let it prevent command registration
                bool valid = false;
                try
                {
                    ConstructLicenseSpot.CheckLicense();
                    ConstructLicenseSpot.EnsureNetworkDeactivatedOnStartup();
                    valid = ConstructLicenseSpot.IsValid;
                }
                catch (Exception ex)
                {
                    LogFull("License check failed", ex);
                }

                if (!isCommandRegistered)
                {
                    //AESC.Construct.SetMode3D
                    var set3DConstruct = Command.Create("AESC.Construct.SetMode3D");
                    set3DConstruct.Hint = Localization.L.T("Construct_Hint_SetMode3D");
                    set3DConstruct.Image = Command.GetCommand("SetMode3D").Image;
                    set3DConstruct.Executing += (s, e) => setMode3D();
                    set3DConstruct.KeepAlive(true);

                    // Bind Alt+D to this command (only if not already taken)
                    const Keys altD = Keys.Alt | Keys.D;
                    if (Command.GetCommand(altD) == null)
                        set3DConstruct.Shortcuts = new[] { altD };

                    // Export to Excel
                    var exportExcel = Command.Create("AESCConstruct2026.ExportExcel");
                    exportExcel.Text = Localization.Language.Translate("Ribbon.Button.ExportExcel");
                    exportExcel.Hint = Localization.L.T("Construct_Hint_ExportExcel");
                    exportExcel.Image = loadImg(Resources.ExcelLogo);
                    exportExcel.IsEnabled = valid;
                    exportExcel.Executing += (s, e) => ExportCommands.ExportExcel(Window.ActiveWindow);
                    exportExcel.KeepAlive(true);

                    // Export BOM
                    var exportBOM = Command.Create("AESCConstruct2026.ExportBOM");
                    exportBOM.Text = Localization.Language.Translate("Ribbon.Button.GenerateBOM");
                    exportBOM.Hint = Localization.L.T("Construct_Hint_ExportBOM");
                    exportBOM.Image = loadImg(Resources.BOMLogo);
                    exportBOM.IsEnabled = valid;
                    exportBOM.Executing += (s, e) => ExportCommands.ExportBOM(Window.ActiveWindow, false);
                    exportBOM.KeepAlive(true);


                    var updateBOM = Command.Create("AESCConstruct2026.UpdateBOM");
                    updateBOM.Text = Localization.Language.Translate("Ribbon.Button.UpdateBOM");
                    updateBOM.Hint = Localization.L.T("Construct_Hint_UpdateBOM");
                    updateBOM.Image = loadImg(Resources.Icon_Update);
                    updateBOM.IsEnabled = valid;
                    updateBOM.Executing += (s, e) => ExportCommands.ExportBOM(Window.ActiveWindow, update: true);
                    updateBOM.KeepAlive(true);

                    // Export STEP
                    var exportSTEP = Command.Create("AESCConstruct2026.ExportSTEP");
                    exportSTEP.Text = Localization.Language.Translate("Ribbon.Button.ExportSTEP");
                    exportSTEP.Hint = Localization.L.T("Construct_Hint_ExportSTEP");
                    exportSTEP.Image = loadImg(Resources.STEPLogo);
                    exportSTEP.IsEnabled = valid;
                    exportSTEP.Executing += (s, e) => ExportCommands.ExportSTEP(Window.ActiveWindow);
                    exportSTEP.KeepAlive(true);
                    //
                    // ─── NEW DXF COMMANDS ────────────────────────────────────────────────────────
                    //

                    // 1) Import DXF Contours
                    var importDxfContours = Command.Create("AESCConstruct2026.ImportDXFContours");
                    importDxfContours.Text = Localization.L.T("Ribbon.Button.ImportDXFContours");
                    importDxfContours.Hint = Localization.L.T("Construct_Hint_ImportDXFContours");
                    importDxfContours.Executing += ImportDXFContours_Execute;
                    importDxfContours.KeepAlive(true);

                    // 2) Convert (open) DXF → Profile
                    var dxfToProfile = Command.Create("AESCConstruct2026.DXFToProfile");
                    dxfToProfile.Text = Localization.L.T("Ribbon.Button.DXFToProfile");
                    dxfToProfile.Hint = Localization.L.T("Construct_Hint_DXFToProfile");
                    dxfToProfile.Executing += DXFtoProfile_Execute;
                    dxfToProfile.KeepAlive(true);

                    // 3) Save DXFProfile list to CSV
                    var saveDxfCsv = Command.Create("AESCConstruct2026.SaveDXFProfileCsv");
                    saveDxfCsv.Text = Localization.L.T("Ribbon.Button.SaveDXFProfileCsv");
                    saveDxfCsv.Hint = Localization.L.T("Construct_Hint_SaveDXFProfileCsv");
                    saveDxfCsv.Executing += SaveDXFProfiles_Execute;
                    saveDxfCsv.KeepAlive(true);

                    // 4) Load DXFProfile list from CSV
                    var loadDxfCsv = Command.Create("AESCConstruct2026.LoadDXFProfileCsv");
                    loadDxfCsv.Text = Localization.L.T("Ribbon.Button.LoadDXFProfileCsv");
                    loadDxfCsv.Hint = Localization.L.T("Construct_Hint_LoadDXFProfileCsv");
                    loadDxfCsv.Executing += LoadDXFProfiles_Execute;
                    loadDxfCsv.KeepAlive(true);

                    // 5)Compare bodies in document
                    var CompareCmd = Command.Create("AESCConstruct2026.CompareBodies");
                    CompareCmd.Text = Localization.L.T("Ribbon.Button.CompareBodies");
                    CompareCmd.Hint = Localization.L.T("Construct_Hint_Compare");
                    CompareCmd.Image = loadImg(Resources.compare);
                    CompareCmd.IsEnabled = valid;
                    CompareCmd.KeepAlive(true);
                    CompareCmd.Executing += (s, e) =>
                        CompareCommand.compareLegacy();

                    // 6) Detect Clashes
                    var clashCmd = Command.Create("AESCConstruct2026.DetectClashes");
                    clashCmd.Text = Localization.L.T("Ribbon.Button.DetectClashes");
                    clashCmd.Hint = Localization.L.T("Construct_Hint_DetectClashes");
                    clashCmd.Image = loadImg(Resources.compare);
                    clashCmd.IsEnabled = valid;
                    clashCmd.Executing += (s, e) => ClashDetectionCommand.DetectClashes(Window.ActiveWindow);
                    clashCmd.KeepAlive(true);

                    //
                    // ─── LEGACY / OTHER COMMANDS ─────────────────────────────────────────────────
                    //

                    // Legacy Joint
                    var jointCmd = Command.Create(ExecuteJointCommand.CommandName);
                    jointCmd.Text = Localization.L.T("Ribbon.Button.ExecuteJointLegacy");
                    jointCmd.Hint = Localization.L.T("Construct_Hint_ExecuteJointLegacy");
                    jointCmd.KeepAlive(true);
                    jointCmd.Executing += (s, e) =>
                        ExecuteJointCommand.ExecuteJoint(Window.ActiveWindow, 0.0, "Miter", false);

                    // Network license toggle — create the command BEFORE the ribbon needs it
                    var cmdNet = Command.Create("AESCConstruct2026.ActivateNetwork");
                    cmdNet.IsEnabled = true;                 // let UpdateNetworkButtonUI refine this later
                    cmdNet.Text = Localization.Language.Translate("Ribbon.Button.ActivateNetwork");
                    cmdNet.KeepAlive(true);                  // IMPORTANT: prevent GC
                    cmdNet.Executing += (s, e) =>            // Use Executing (not Executed)
                    {
                        try
                        {
                            // Make sure we have a license handle/state
                            ConstructLicenseSpot.CheckLicense();

                            var lic = ConstructLicenseSpot.CurrentLicense;
                            if (lic == null)
                            {
                                Localization.L.Status("Construct_Net_Msg_NoHandle", StatusMessageType.Error);

                                return;
                            }

                            if (!lic.IsNetwork)
                            {
                                Localization.L.Status("Construct_Net_Msg_NotNetwork", StatusMessageType.Warning);
                                return;
                            }

                            // Toggle
                            if (lic.IsValidConnection())
                                ConstructLicenseSpot.LicenseCheckIn();
                            else
                                ConstructLicenseSpot.LicenseCheckOut();

                            // Refresh UI after the operation
                            RefreshLicenseUI();
                            ConstructLicenseSpot.UpdateNetworkButtonUI();
                        }
                        catch (Exception ex)
                        {
                            Localization.L.Status("Construct_Net_Err_ToggleFailed", StatusMessageType.Warning, ex.Message);
                        }
                    };

                    // Initial button state (enabled/disabled) based on license type
                    try { ConstructLicenseSpot.UpdateNetworkButtonUI(); } catch (Exception ex) { Logger.Log("UpdateNetworkButtonUI failed: " + ex.ToString()); }

                    // Sidebar commands
                    var _ = Localization.Language.Translate("Settings");
                    UIManager.RegisterAll();
                    UpdateCommandTexts();
                    UIManager.UpdateCommandTexts();

                    isCommandRegistered = true;
                }
                Logger.Log("Commands registered");
            }
            catch (Exception ex)
            {
                LogFull("Initialize() failed", ex);
            }
        }

        // Logs a message with the full exception chain (including InnerException).
        private static void LogFull(string context, Exception ex)
        {
            Logger.Log(context + ": " + ex.ToString());
            for (var inner = ex.InnerException; inner != null; inner = inner.InnerException)
                Logger.Log("  Caused by: " + inner.ToString());
        }

        // Switches SpaceClaim to solid interaction mode while temporarily hiding all visible curves.
        public static void setMode3D()
        {

            try
            {
                // Find all smaller bends that are in the same part as the selected bend.
                var curveList = new List<IDesignCurve>();
                foreach (var curve in Window.ActiveWindow.ActiveContext.Root.GetDescendants<IDesignCurve>())
                {
                    if (curve.IsVisible(null))
                    {
                        curveList.Add(curve);
                        curve.SetVisibility(null, false);
                    }
                }

                Window.ActiveWindow.InteractionMode = InteractionMode.Solid;
                Command.Execute("Select");

                foreach (var curve in curveList)
                {
                    curve.SetVisibility(null, true);
                }
            }

            catch (Exception ex)
            {
                Logger.Log("[Construct2026] setMode3D failed: " + ex.ToString());
                Localization.L.Status("Construct_Err_SetMode3D", StatusMessageType.Error, ex.Message);
            }
        }

        // Updates the enabled state of license dependent commands and refreshes all related UI text.
        public static void RefreshLicenseUI()
        {
            bool valid = ConstructLicenseSpot.IsValid;

            SetEnabled("AESCConstruct2026.ExportExcel", valid);
            SetEnabled("AESCConstruct2026.ExportBOM", valid);
            SetEnabled("AESCConstruct2026.UpdateBOM", valid);
            SetEnabled("AESCConstruct2026.ExportSTEP", valid);
            SetEnabled("AESCConstruct2026.CompareBodies", valid);
            SetEnabled("AESCConstruct2026.DetectClashes", valid);

            UIManager.RefreshLicenseUI();
            UpdateCommandTexts();
        }

        // Refreshes localized button texts and hints for all registered ribbon commands.
        // Note: SpaceClaim may cache Command.Hint after the ribbon is first built, so a live
        // hint refresh is best-effort; the label (Text) refresh is reliable.
        private static void SetTextHint(string id, string textKey, string hintKey)
        {
            var cmd = Command.GetCommand(id);
            if (cmd == null) return;
            if (textKey != null) cmd.Text = Localization.L.T(textKey);
            if (hintKey != null) cmd.Hint = Localization.L.T(hintKey);
        }

        public static void UpdateCommandTexts()
        {
            SetTextHint("AESC.Construct.SetMode3D", null, "Construct_Hint_SetMode3D");
            SetTextHint("AESCConstruct2026.ExportExcel", "Ribbon.Button.ExportExcel", "Construct_Hint_ExportExcel");
            SetTextHint("AESCConstruct2026.ExportBOM", "Ribbon.Button.GenerateBOM", "Construct_Hint_ExportBOM");
            SetTextHint("AESCConstruct2026.UpdateBOM", "Ribbon.Button.UpdateBOM", "Construct_Hint_UpdateBOM");
            SetTextHint("AESCConstruct2026.ExportSTEP", "Ribbon.Button.ExportSTEP", "Construct_Hint_ExportSTEP");
            SetTextHint("AESCConstruct2026.ImportDXFContours", "Ribbon.Button.ImportDXFContours", "Construct_Hint_ImportDXFContours");
            SetTextHint("AESCConstruct2026.DXFToProfile", "Ribbon.Button.DXFToProfile", "Construct_Hint_DXFToProfile");
            SetTextHint("AESCConstruct2026.SaveDXFProfileCsv", "Ribbon.Button.SaveDXFProfileCsv", "Construct_Hint_SaveDXFProfileCsv");
            SetTextHint("AESCConstruct2026.LoadDXFProfileCsv", "Ribbon.Button.LoadDXFProfileCsv", "Construct_Hint_LoadDXFProfileCsv");
            SetTextHint("AESCConstruct2026.CompareBodies", "Ribbon.Button.CompareBodies", "Construct_Hint_Compare");
            SetTextHint("AESCConstruct2026.DetectClashes", "Ribbon.Button.DetectClashes", "Construct_Hint_DetectClashes");
            SetTextHint(ExecuteJointCommand.CommandName, "Ribbon.Button.ExecuteJointLegacy", "Construct_Hint_ExecuteJointLegacy");
            SetTextHint("AESCConstruct2026.ActivateNetwork", "Ribbon.Button.ActivateNetworkBtn", null);
        }

        // Creates a bitmap image from embedded byte resources for use as command icons.
        private Image loadImg(byte[] bytes)
        {
            try
            {
                return new Bitmap(new MemoryStream(bytes));
            }
            catch (Exception ex)
            {
                Logger.Log("loadImg failed: " + ex.ToString());
                return null;
            }
        }

        // Called by SpaceClaim when the addin AppDomain is being unloaded (session end,
        // addin refresh, or reload after build). Must close the PanelTab here so the
        // docked tab visual is removed from SpaceClaim's sidebar before our AppDomain
        // goes away — otherwise a subsequent AppDomain creates a second PanelTab while
        // the orphaned first one lingers in the UI.
        public void Disconnect()
        {
            try
            {
                Logger.Log("[Construct2026] Disconnect() called — cleaning up panel");
                UIManager.ClosePanelAndDispose();
            }
            catch (Exception ex)
            {
                Logger.Log("[Construct2026] Disconnect cleanup failed: " + ex.ToString());
            }
        }

        // Helper to enable or disable a ribbon command by id.
        private static void SetEnabled(string commandId, bool enabled)
        {
            var cmd = Command.GetCommand(commandId);
            if (cmd != null) cmd.IsEnabled = enabled;
        }

        // Loads the custom ribbon XML definition that SpaceClaim uses to build the UI for this add-in.
        public string GetCustomUI()
        {
            try
            {
                string resourceName = "AESCConstruct2026.UIMain.Ribbon.xml";

                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                        throw new InvalidOperationException($"Resource {resourceName} not found. Ensure it is set as 'Embedded Resource'.");

                    using (StreamReader reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log("[Construct2026] GetCustomUI failed: " + ex.Message);
                return "";
            }
        }

        // Command handler: prompts for a DXF file and imports its contours into the active document.
        private void ImportDXFContours_Execute(object sender, EventArgs e)
        {
            // Prompt user to pick a .dxf file
            string filePath;
            using (var dlg = new OpenFileDialog
            {
                Title = Localization.L.T("DXF_FileDialog_ImportTitle"),
                Filter = Localization.L.T("DXF_FileFilter_Dxf")
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                filePath = dlg.FileName;
            }

            // Call the helper
            if (DXFImportHelper.ImportDXFContours(filePath, out var contours))
            {
                Localization.L.Status("DXF_Msg_Imported", StatusMessageType.Information, contours.Count, filePath);
            }
            else
            {
                Localization.L.Status("DXF_Err_ImportFailed", StatusMessageType.Error, filePath);
            }
        }

        // Command handler: converts the active DXF window into a profile, copies it to clipboard and shows a preview image.
        private void DXFtoProfile_Execute(object sender, EventArgs e)
        {
            // Assumes a DXF window is currently active
            var profile = DXFImportHelper.DXFtoProfile();
            if (profile == null)
            {
                Localization.L.Status("DXF_Err_ProfileInvalid", StatusMessageType.Error);
                return;
            }

            // Copy the ProfileString to the clipboard
            Clipboard.SetText(profile.ProfileString);

            Localization.L.Status("DXF_Msg_ProfileSucceeded", StatusMessageType.Information, profile.Name);

            // Decode and display the preview image if available
            if (!string.IsNullOrEmpty(profile.ImgString))
            {
                Bitmap preview;
                byte[] bytes = Convert.FromBase64String(profile.ImgString);
                using (var ms = new MemoryStream(bytes))
                using (var bmp = new Bitmap(ms))
                {
                    preview = new Bitmap(bmp);
                }

                var frm = new Form
                {
                    Text = "DXF Preview: " + profile.Name,
                    ClientSize = new System.Drawing.Size(preview.Width, preview.Height)
                };
                var pb = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    Image = preview,
                    SizeMode = PictureBoxSizeMode.Zoom
                };
                frm.Controls.Add(pb);
                frm.FormClosed += (s2, e2) => preview.Dispose();
                frm.ShowDialog();
            }
        }

        // Command handler: saves all DXF profiles collected in the current session to a CSV file.
        private void SaveDXFProfiles_Execute(object sender, EventArgs e)
        {
            var profiles = DXFImportHelper.SessionProfiles;
            var profileCount = profiles?.Count() ?? 0;
            if (profileCount == 0)
            {
                Localization.L.Status("DXF_Msg_NoProfilesToSave", StatusMessageType.Information);
                return;
            }

            string csvPath;
            using (var dlg = new SaveFileDialog
            {
                Title = Localization.L.T("DXF_FileDialog_SaveTitle"),
                Filter = Localization.L.T("DXF_FileFilter_Csv"),
                FileName = "DXFProfiles.csv"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                csvPath = dlg.FileName;
            }

            try
            {
                DXFImportHelper.DXFProfileCsvHandler.SaveDXFProfiles(csvPath, profiles);

                Localization.L.Status("DXF_Msg_Saved", StatusMessageType.Information, profileCount, csvPath);
            }
            catch (Exception ex)
            {
                Localization.L.Status("DXF_Err_SaveFailed", StatusMessageType.Error, ex.Message);
            }
        }

        // Returns localized labels for ribbon controls based on their control id.
        public string GetRibbonLabel(string controlId)
        {
            switch (controlId)
            {
                // Buttons
                case "AESCConstruct2026.ProfileSidebarBtn":
                    return Localization.Language.Translate("Ribbon.Button.FrameGenerator");
                case "AESCConstruct2026.ExportSTEPBtn":
                    return Localization.Language.Translate("Ribbon.Button.ExportSTEP");
                case "AESCConstruct2026.CompareBodiesBtn":
                    return Localization.Language.Translate("Ribbon.Button.CompareBodies");
                case "AESCConstruct2026.DetectClashesBtn":
                    return Localization.Language.Translate("Ribbon.Button.DetectClashes");
                case "AESCConstruct2026.ExportBOMBtn":
                    return Localization.Language.Translate("Ribbon.Button.GenerateBOM");
                case "AESCConstruct2026.UpdateBOM":
                    return Localization.Language.Translate("Ribbon.Button.UpdateBOM");
                case "AESCConstruct2026.ExportExcelBtn":
                    return Localization.Language.Translate("Ribbon.Button.ExportExcel");
                case "AESCConstruct2026.Plate":
                    return Localization.Language.Translate("Ribbon.Button.Plate");
                case "AESCConstruct2026.Fastener":
                    return Localization.Language.Translate("Ribbon.Button.Fastener");
                case "AESCConstruct2026.RibCutOut":
                    return Localization.Language.Translate("Ribbon.Button.RibCutOut");
                case "AESCConstruct2026.SettingsSidebarBtn":
                    return Localization.Language.Translate("Ribbon.Button.Settings");
                case "AESCConstruct2026.ConnectorSidebarBtn":
                    return Localization.Language.Translate("Ribbon.Button.Connector");
                case "AESCConstruct2026.ActivateNetworkBtn":
                    return Localization.Language.Translate("Ribbon.Button.ActivateNetwork");

                // Groups
                case "AESCConstruct2026.Group":
                    return Localization.Language.Translate("Ribbon.Group.FrameGenerator");
                case "AESCConstruct2026.ToolsGroup":
                    return Localization.Language.Translate("Ribbon.Group.SketchTools");
                case "AESCConstruct2026.ExportGroup":
                    return Localization.Language.Translate("Ribbon.Group.Export");
                case "AESCConstruct2026.PlateGroup":
                    return Localization.Language.Translate("Ribbon.Group.Plate");
                case "AESCConstruct2026.FastenerGroup":
                    return Localization.Language.Translate("Ribbon.Group.Fastener");
                case "AESCConstruct2026.ConnectorGroup":
                    return Localization.Language.Translate("Ribbon.Group.Connector");
                case "AESCConstruct2026.RibCutOutGroup":
                    return Localization.Language.Translate("Ribbon.Group.RibCutOut");
                case "AESCConstruct2026.Engraving":
                    return Localization.Language.Translate("Ribbon.Group.Engraving");
                case "AESCConstruct2026.CustomProperties":
                    return Localization.Language.Translate("Ribbon.Group.CustomProperties");
                case "AESCConstruct2026.ToolGroup":
                    return Localization.Language.Translate("Ribbon.Group.Settings");
                case "AESCConstruct2026.Network":
                    return Localization.Language.Translate("Ribbon.Group.Network");

                default:
                    // If you ever add more ids, they will at least show their id
                    return controlId;
            }
        }

        // Returns the localized label for engraving related ribbon controls.
        public string GetEngravingLabel(string controlId) => Localization.Language.Translate("Ribbon.Button.Engraving");

        // Returns the localized label for network activation related ribbon controls.
        public string GetNetworkLabel(string controlId) => Localization.Language.Translate("Ribbon.Button.ActivateNetwork");

        // Returns the localized label for custom properties related ribbon controls.
        public string GetCustomPropertiesLabel(string controlId) => Localization.Language.Translate("Ribbon.Button.CustomProperties");

        // Command handler: loads DXF profiles from a CSV, stores them in the session, and optionally rebuilds bodies.
        private void LoadDXFProfiles_Execute(object sender, EventArgs e)
        {
            string csvPath;
            using (var dlg = new OpenFileDialog
            {
                Title = Localization.L.T("DXF_FileDialog_LoadTitle"),
                Filter = Localization.L.T("DXF_FileFilter_Csv")
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                csvPath = dlg.FileName;
            }

            try
            {
                var profiles = DXFImportHelper.DXFProfileCsvHandler.LoadDXFProfiles(csvPath);
                DXFImportHelper.SessionProfiles.Clear();
                DXFImportHelper.SessionProfiles.AddRange(profiles);


                Localization.L.Status("DXF_Msg_Loaded", StatusMessageType.Information, profiles.Count, csvPath);
                // (Optional) Immediately reconstruct each profile in the main part:
                foreach (var prof in profiles)
                {
                    var body = DXFImportHelper.BodyFromString(prof.ProfileString);
                    DesignBody.Create(Window.ActiveWindow.Document.MainPart, $"Rebuild_{prof.Name}", body);
                }
            }
            catch (Exception ex)
            {
                Localization.L.Status("DXF_Err_LoadFailed", StatusMessageType.Error, ex.Message);
            }
        }
    }
}
