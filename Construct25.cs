using AESCConstruct25.FrameGenerator.Commands;
using AESCConstruct25.FrameGenerator.Utilities;
using AESCConstruct25.Licensing;
using AESCConstruct25.UIMain;
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

namespace AESCConstruct25
{
    [Serializable]
    public class Construct25 : MarshalByRefObject, IExtensibility, IRibbonExtensibility
    {
        private static bool isCommandRegistered = false;

        public bool Connect()
        {
            try
            {
                Api.Initialize();

                // Attach to first active session
                var session = Session.GetSessions().FirstOrDefault();
                if (session == null)
                {
                    return false;
                }

                Api.AttachToSession(session);

                bool LicenseValid = ConstructLicenseSpot.CheckLicense(); 

                ConstructLicenseSpot.EnsureNetworkDeactivatedOnStartup();

                if (!isCommandRegistered)
                {
                    bool valid = ConstructLicenseSpot.IsValid;

                    //AESC.Construct.SetMode3D
                    var set3DConstruct = Command.Create("AESC.Construct.SetMode3D");
                    set3DConstruct.Hint = "3D mode without converting closed line loops to surfaces";
                    set3DConstruct.Image = Command.GetCommand("SetMode3D").Image;
                    set3DConstruct.Executing += (s, e) => setMode3D();
                    set3DConstruct.KeepAlive(true);

                    // Export to Excel
                    var exportExcel = Command.Create("AESCConstruct25.ExportExcel");
                    exportExcel.Text = Localization.Language.Translate("Ribbon.Button.ExportExcel");
                    exportExcel.Hint = "Export frame data to an Excel file.";
                    exportExcel.Image = loadImg(Resources.ExcelLogo);
                    exportExcel.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
                    exportExcel.Executing += (s, e) => ExportCommands.ExportExcel(Window.ActiveWindow);
                    exportExcel.KeepAlive(true);

                    // Export BOM
                    var exportBOM = Command.Create("AESCConstruct25.ExportBOM");
                    exportBOM.Text = Localization.Language.Translate("Ribbon.Button.GenerateBOM");
                    exportBOM.Hint = "Create a bill-of-materials.";
                    exportBOM.Image = loadImg(Resources.BOMLogo);
                    exportBOM.IsEnabled = valid;// LicenseSpot.LicenseSpot.State.Valid;
                    exportBOM.Executing += (s, e) => ExportCommands.ExportBOM(Window.ActiveWindow, false);
                    exportBOM.KeepAlive(true);


                    var updateBOM = Command.Create("AESCConstruct25.UpdateBOM");
                    updateBOM.Text = Localization.Language.Translate("Ribbon.Button.UpdateBOM");
                    updateBOM.Hint = "Update an existing bill-of-materials.";
                    updateBOM.Image = loadImg(Resources.Icon_Update);
                    updateBOM.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
                    updateBOM.Executing += (s, e) => ExportCommands.ExportBOM(Window.ActiveWindow, update: true);
                    updateBOM.KeepAlive(true);

                    // Export STEP
                    var exportSTEP = Command.Create("AESCConstruct25.ExportSTEP");
                    exportSTEP.Text = Localization.Language.Translate("Ribbon.Button.ExportSTEP");
                    exportSTEP.Hint = "Export frame as a STEP file.";
                    exportSTEP.Image = loadImg(Resources.STEPLogo);
                    exportSTEP.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
                    exportSTEP.Executing += (s, e) => ExportCommands.ExportSTEP(Window.ActiveWindow);
                    exportSTEP.KeepAlive(true);
                    //
                    // ─── NEW DXF COMMANDS ────────────────────────────────────────────────────────
                    //

                    // 1) Import DXF Contours
                    var importDxfContours = Command.Create("AESCConstruct25.ImportDXFContours");
                    importDxfContours.Text = "Import DXF Contours";
                    importDxfContours.Hint = "Load DXF contours into the active document.";
                    importDxfContours.Executing += ImportDXFContours_Execute;
                    importDxfContours.KeepAlive(true);

                    // 2) Convert (open) DXF → Profile
                    var dxfToProfile = Command.Create("AESCConstruct25.DXFToProfile");
                    dxfToProfile.Text = "DXF → Profile";
                    dxfToProfile.Hint = "Convert an open DXF window into a profile string and preview image.";
                    dxfToProfile.Executing += DXFtoProfile_Execute;
                    dxfToProfile.KeepAlive(true);

                    // 3) Save DXFProfile list to CSV
                    var saveDxfCsv = Command.Create("AESCConstruct25.SaveDXFProfileCsv");
                    saveDxfCsv.Text = "Save DXFProfile CSV";
                    saveDxfCsv.Hint = "Save all collected DXFProfile objects to a CSV file.";
                    saveDxfCsv.Executing += SaveDXFProfiles_Execute;
                    saveDxfCsv.KeepAlive(true);

                    // 4) Load DXFProfile list from CSV
                    var loadDxfCsv = Command.Create("AESCConstruct25.LoadDXFProfileCsv");
                    loadDxfCsv.Text = "Load DXFProfile CSV";
                    loadDxfCsv.Hint = "Load DXFProfile objects from a CSV file.";
                    loadDxfCsv.Executing += LoadDXFProfiles_Execute;
                    loadDxfCsv.KeepAlive(true);

                    // 5)Compare bodies in document
                    var CompareCmd = Command.Create("AESCConstruct25.CompareBodies");
                    CompareCmd.Text = "Compare";
                    CompareCmd.Hint = "Compare bodies to look for duplicates";
                    CompareCmd.Image = loadImg(Resources.compare);
                    CompareCmd.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
                    CompareCmd.KeepAlive(true);
                    CompareCmd.Executing += (s, e) =>
                        CompareCommand.CompareSimple();

                    //
                    // ─── LEGACY / OTHER COMMANDS ─────────────────────────────────────────────────
                    //

                    // Legacy Joint
                    var jointCmd = Command.Create(ExecuteJointCommand.CommandName);
                    jointCmd.Text = "Execute Joint (Legacy)";
                    jointCmd.Hint = "Applies a joint between selected components.";
                    jointCmd.KeepAlive(true);
                    jointCmd.Executing += (s, e) =>
                        ExecuteJointCommand.ExecuteJoint(Window.ActiveWindow, 0.0, "Miter", false);

                    // Network license toggle — create the command BEFORE the ribbon needs it
                    var cmdNet = Command.Create("AESCConstruct25.ActivateNetwork");
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
                                Application.ReportStatus("No license handle available (activate or check your license files).", StatusMessageType.Error, null);

                                return;
                            }

                            if (!lic.IsNetwork)
                            {
                                Application.ReportStatus("Current license is not a network license.", StatusMessageType.Warning, null);
                                return;
                            }

                            // Toggle
                            if (lic.IsValidConnection())
                                ConstructLicenseSpot.licenseCheckIn();
                            else
                                ConstructLicenseSpot.licenseCheckOut();

                            // Refresh UI after the operation
                            RefreshLicenseUI();
                            ConstructLicenseSpot.UpdateNetworkButtonUI();
                        }
                        catch (Exception ex)
                        {
                            Application.ReportStatus("Network license toggle failed:\n" + ex.Message, StatusMessageType.Warning, null);
                        }
                    };

                    // Initial button state (enabled/disabled) based on license type
                    ConstructLicenseSpot.UpdateNetworkButtonUI();

                    // Sidebar commands
                    var _ = Localization.Language.Translate("Settings");
                    UIManager.RegisterAll();
                    UpdateCommandTexts();
                    UIManager.UpdateCommandTexts();

                    isCommandRegistered = true;
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
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
                MessageBox.Show(ex.ToString());
            }
        }

        public static void RefreshLicenseUI()
        {
            bool valid = ConstructLicenseSpot.IsValid;

            SetEnabled("AESCConstruct25.ExportExcel", valid);
            SetEnabled("AESCConstruct25.ExportBOM", valid);
            SetEnabled("AESCConstruct25.UpdateBOM", valid);
            SetEnabled("AESCConstruct25.ExportSTEP", valid);
            SetEnabled("AESCConstruct25.CompareBodies", valid);

            UIManager.RefreshLicenseUI();
            UpdateCommandTexts();
        }

        public static void UpdateCommandTexts()
        {
            var cmd = Command.GetCommand("AESCConstruct25.ExportExcel");
            if (cmd != null) cmd.Text = Localization.Language.Translate("Ribbon.Button.ExportExcel");

            cmd = Command.GetCommand("AESCConstruct25.ExportBOM");
            if (cmd != null) cmd.Text = Localization.Language.Translate("Ribbon.Button.GenerateBOM");

            cmd = Command.GetCommand("AESCConstruct25.UpdateBOM");
            if (cmd != null) cmd.Text = Localization.Language.Translate("Ribbon.Button.UpdateBOM");

            cmd = Command.GetCommand("AESCConstruct25.ExportSTEP");
            if (cmd != null) cmd.Text = Localization.Language.Translate("Ribbon.Button.ExportSTEP");

            cmd = Command.GetCommand("AESCConstruct25.ActivateNetwork");
            if (cmd != null) cmd.Text = Localization.Language.Translate("Ribbon.Button.ActivateNetworkBtn");
        }

        private Image loadImg(byte[] bytes)
        {
            try
            {
                return new Bitmap(new MemoryStream(bytes));
            }
            catch
            {
                return null;
            }
        }

        public void Disconnect()
        {

        }

        private static void SetEnabled(string commandId, bool enabled)
        {
            var cmd = Command.GetCommand(commandId);
            if (cmd != null) cmd.IsEnabled = enabled;
        }

        public string GetCustomUI()
        {
            try
            {
                string resourceName = "AESCConstruct25.UIMain.Ribbon.xml";

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
            catch (Exception)
            {
                return "";
            }
        }

        private void ImportDXFContours_Execute(object sender, EventArgs e)
        {
            // Prompt user to pick a .dxf file
            string filePath;
            using (var dlg = new OpenFileDialog
            {
                Title = "Select DXF file to import",
                Filter = "DXF Files (*.dxf)|*.dxf"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;
                filePath = dlg.FileName;
            }

            // Call the helper
            if (DXFImportHelper.ImportDXFContours(filePath, out var contours))
            {
                Application.ReportStatus($"Imported {contours.Count} contour curves from:\n{filePath}", StatusMessageType.Information, null);
            }
            else
            {
                Application.ReportStatus($"Failed to import valid contours from:\n{filePath}", StatusMessageType.Error, null);
            }
        }

        private void DXFtoProfile_Execute(object sender, EventArgs e)
        {
            // Assumes a DXF window is currently active
            var profile = DXFImportHelper.DXFtoProfile();
            if (profile == null)
            {
                Application.ReportStatus("DXF → Profile failed or was invalid.", StatusMessageType.Error, null);
                return;
            }

            // Copy the ProfileString to the clipboard
            Clipboard.SetText(profile.ProfileString);

            Application.ReportStatus($"DXF→Profile succeeded.\n\nName = {profile.Name}\n(Profile string copied to clipboard.)", StatusMessageType.Information, null);

            // Decode and display the preview image if available
            if (!string.IsNullOrEmpty(profile.ImgString))
            {
                byte[] bytes = Convert.FromBase64String(profile.ImgString);
                using (var ms = new MemoryStream(bytes))
                {
                    var bmp = new Bitmap(ms);
                    var frm = new Form
                    {
                        Text = "DXF Preview: " + profile.Name,
                        ClientSize = new System.Drawing.Size(bmp.Width, bmp.Height)
                    };
                    var pb = new PictureBox
                    {
                        Dock = DockStyle.Fill,
                        Image = bmp,
                        SizeMode = PictureBoxSizeMode.Zoom
                    };
                    frm.Controls.Add(pb);
                    frm.ShowDialog();
                }
            }
        }

        private void SaveDXFProfiles_Execute(object sender, EventArgs e)
        {
            var profiles = DXFImportHelper.SessionProfiles;
            if (profiles == null || profiles.Count() == 0)
            {
                Application.ReportStatus("No DXF profiles available to save.", StatusMessageType.Information, null);
                return;
            }

            string csvPath;
            using (var dlg = new SaveFileDialog
            {
                Title = "Save DXFProfile List as CSV",
                Filter = "CSV Files (*.csv)|*.csv",
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

                Application.ReportStatus($"Saved {profiles.Count()} profiles to:\n{csvPath}", StatusMessageType.Information, null);
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to save CSV:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        public string GetRibbonLabel(string controlId)
        {
            // Map only button IDs (you said groups don’t need translation)
            switch (controlId)
            {
                case "AESCConstruct25.ProfileSidebarBtn":
                    return Localization.Language.Translate("Ribbon.Button.FrameGenerator");

                case "AESCConstruct25.ExportSTEPBtn":
                    return Localization.Language.Translate("Ribbon.Button.ExportSTEP");
                case "AESCConstruct25.ExportBOMBtn":
                    return Localization.Language.Translate("Ribbon.Button.GenerateBOM");
                case "AESCConstruct25.UpdateBOM":
                    return Localization.Language.Translate("Ribbon.Button.UpdateBOM");
                case "AESCConstruct25.ExportExcelBtn":
                    return Localization.Language.Translate("Ribbon.Button.ExportExcel");

                case "AESCConstruct25.Plate":
                    return Localization.Language.Translate("Ribbon.Button.Plate");
                case "AESCConstruct25.Fastener":
                    return Localization.Language.Translate("Ribbon.Button.Fastener");
                case "AESCConstruct25.RibCutOut":
                    return Localization.Language.Translate("Ribbon.Button.RibCutOut");
                case "AESCConstruct25.SettingsSidebarBtn":
                    return Localization.Language.Translate("Ribbon.Button.Settings");
                case "AESCConstruct25.ConnectorSidebarBtn":
                    return Localization.Language.Translate("Ribbon.Button.Connector");
                case "AESCConstruct25.ActivateNetworkBtn":
                    return Localization.Language.Translate("Ribbon.Button.ActivateNetwork");

                case "AESCConstruct25.DockBtn":
                    // Reflect current toggle state by reading the command’s Text,
                    // which UIManager already sets to a translated string.
                    var dockCmd = Command.GetCommand("AESCConstruct25.DockCmd");
                    if (dockCmd != null && !string.IsNullOrWhiteSpace(dockCmd.Text))
                        return dockCmd.Text;
                    // Fallback if command isn’t created yet:
                    return Localization.Language.Translate("Ribbon.Button.Float");

                default:
                    // If you ever add more ids, they will at least show their id
                    return controlId;
            }
        }

        public string GetEngravingLabel(string controlId) => Localization.Language.Translate("Ribbon.Button.Engraving");
        public string GetNetworkLabel(string controlId) => Localization.Language.Translate("Ribbon.Button.ActivateNetwork");
        public string GetCustomPropertiesLabel(string controlId) => Localization.Language.Translate("Ribbon.Button.CustomProperties");

        private void LoadDXFProfiles_Execute(object sender, EventArgs e)
        {
            string csvPath;
            using (var dlg = new OpenFileDialog
            {
                Title = "Load DXFProfile CSV",
                Filter = "CSV Files (*.csv)|*.csv"
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


                Application.ReportStatus($"Loaded {profiles.Count} profiles from:\n{csvPath}", StatusMessageType.Information, null);
                // (Optional) Immediately reconstruct each profile in the main part:
                foreach (var prof in profiles)
                {
                    var body = DXFImportHelper.BodyFromString(prof.ProfileString);
                    DesignBody.Create(Window.ActiveWindow.Document.MainPart, $"Rebuild_{prof.Name}", body);
                }
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to load CSV:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }
    }
}
