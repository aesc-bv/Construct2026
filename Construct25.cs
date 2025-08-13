using AESCConstruct25.FrameGenerator.Commands;
using AESCConstruct25.FrameGenerator.Utilities;
using AESCConstruct25.UIMain;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Extensibility;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Image = System.Drawing.Image;

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

                LicenseSpot.LicenseSpot.Initialize();

                if (!isCommandRegistered)
                {
                    //Logger.Log("AESCConstruct25: Registering Commands...");
                    bool valid = LicenseSpot.LicenseSpot.State.Valid;//DateTime.Now < new DateTime(2025, 8, 31, 23, 59, 59);

                    // Export to Excel
                    var exportExcel = Command.Create("AESCConstruct25.ExportExcel");
                    exportExcel.Text = Localization.Language.Translate("Ribbon.Button.ExportExcel");
                    exportExcel.Hint = "Export frame data to an Excel file.";
                    exportExcel.Image = loadImg(Resources.ExcelLogo);
                    exportExcel.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
                    exportExcel.Executing += (s, e) => ExportCommands.ExportExcel(Window.ActiveWindow);
                    exportExcel.KeepAlive(true);
                    //Logger.Log("AESCConstruct25: 2");

                    // Export BOM
                    var exportBOM = Command.Create("AESCConstruct25.ExportBOM");
                    exportBOM.Text = Localization.Language.Translate("Ribbon.Button.GenerateBOM");
                    exportBOM.Hint = "Create a bill-of-materials.";
                    exportBOM.Image = loadImg(Resources.BOMLogo);
                    exportBOM.IsEnabled = valid;// LicenseSpot.LicenseSpot.State.Valid;
                    exportBOM.Executing += (s, e) => ExportCommands.ExportBOM(Window.ActiveWindow, false);
                    exportBOM.KeepAlive(true);
                    //Logger.Log("AESCConstruct25: 3");


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
                    //Logger.Log("AESCConstruct25: 4");

                    //
                    // ─── NEW DXF COMMANDS ────────────────────────────────────────────────────────
                    //

                    // 1) Import DXF Contours
                    var importDxfContours = Command.Create("AESCConstruct25.ImportDXFContours");
                    importDxfContours.Text = "Import DXF Contours";
                    importDxfContours.Hint = "Load DXF contours into the active document.";
                    importDxfContours.Executing += ImportDXFContours_Execute;
                    importDxfContours.KeepAlive(true);
                    //Logger.Log("AESCConstruct25: 5");

                    // 2) Convert (open) DXF → Profile
                    var dxfToProfile = Command.Create("AESCConstruct25.DXFToProfile");
                    dxfToProfile.Text = "DXF → Profile";
                    dxfToProfile.Hint = "Convert an open DXF window into a profile string and preview image.";
                    dxfToProfile.Executing += DXFtoProfile_Execute;
                    dxfToProfile.KeepAlive(true);
                    //Logger.Log("AESCConstruct25: 6");

                    // 3) Save DXFProfile list to CSV
                    var saveDxfCsv = Command.Create("AESCConstruct25.SaveDXFProfileCsv");
                    saveDxfCsv.Text = "Save DXFProfile CSV";
                    saveDxfCsv.Hint = "Save all collected DXFProfile objects to a CSV file.";
                    saveDxfCsv.Executing += SaveDXFProfiles_Execute;
                    saveDxfCsv.KeepAlive(true);
                    //Logger.Log("AESCConstruct25: 7");

                    // 4) Load DXFProfile list from CSV
                    var loadDxfCsv = Command.Create("AESCConstruct25.LoadDXFProfileCsv");
                    loadDxfCsv.Text = "Load DXFProfile CSV";
                    loadDxfCsv.Hint = "Load DXFProfile objects from a CSV file.";
                    loadDxfCsv.Executing += LoadDXFProfiles_Execute;
                    loadDxfCsv.KeepAlive(true);
                    //Logger.Log("AESCConstruct25: 8");

                    // 5)Compare bodies in document
                    var CompareCmd = Command.Create("AESCConstruct25.CompareBodies");
                    CompareCmd.Text = "Compare";
                    CompareCmd.Hint = "Compare bodies to look for duplicates";
                    CompareCmd.Image = loadImg(Resources.compare);
                    CompareCmd.IsEnabled = valid;//LicenseSpot.LicenseSpot.State.Valid;
                    CompareCmd.KeepAlive(true);
                    CompareCmd.Executing += (s, e) =>
                        CompareCommand.CompareSimple();
                    //Logger.Log("AESCConstruct25: 9");

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

                    // Sidebar commands
                    var _ = Localization.Language.Translate("Settings");
                    UIManager.RegisterAll();
                    UpdateCommandTexts();
                    UIManager.UpdateCommandTexts();
                    //UIManager.ShowPanel("RibCutOut");

                    isCommandRegistered = true;
                }

                return true;
            }
            catch (Exception)
            {
                //Logger.Log($"AESCConstruct25: Error during initialization - {ex.Message}");
                return false;
            }
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
            //Logger.Log("AESCConstruct25: Disconnect() called");
        }

        public string GetCustomUI()
        {
            try
            {
                string resourceName = "AESCConstruct25.UIMain.Ribbon.xml";
                // Logger.Log($"AESCConstruct25: Loading Ribbon UI from {resourceName}");

                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                        throw new InvalidOperationException($"Resource {resourceName} not found. Ensure it is set as 'Embedded Resource'.");

                    using (StreamReader reader = new StreamReader(stream))
                    {
                        // Logger.Log("AESCConstruct25: Ribbon UI Loaded Successfully");
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception)
            {
                // Logger.Log($"AESCConstruct25: Error loading Ribbon UI - {ex.Message}");
                return "";
            }
        }

        private void Profile_Executing(object sender, EventArgs e)
        {
            //Logger.Log("AESCConstruct25: Profile Command Executing");

            Window activeWindow = Window.ActiveWindow;
            if (activeWindow == null)
            {
                //Logger.Log("AESCConstruct25: ERROR - No active window detected.");
                MessageBox.Show("No active SpaceClaim window found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<ITrimmedCurve> selectedCurves = ProfileSelectionHelper.GetSelectedCurves(activeWindow);

            bool hasComponentOrBodySelected = activeWindow.ActiveContext.Selection
                .OfType<IDocObject>()
                .Any(obj => obj is Component || obj is DesignBody);

            if (selectedCurves.Count == 0 && !hasComponentOrBodySelected)
            {
                //Logger.Log("AESCConstruct25: ERROR - No valid lines or bodies selected.");
                MessageBox.Show("Please select at least one valid line, edge, body or component before opening the profile selection window.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //ProfileSelectionWindow profileWindow = new ProfileSelectionWindow();
            //profileWindow.ShowDialog();
        }

        //
        // ─── DXF COMMAND HANDLERS ───────────────────────────────────────────────────────
        //

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
                MessageBox.Show(
                    $"Imported {contours.Count} contour curves from:\n{filePath}",
                    "DXF Import",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                // (Optional) create a planar body just to visualize
                //var planar = Body.CreatePlanarBody(Plane.PlaneXY, contours);
                //DesignBody.Create(Window.ActiveWindow.Document.MainPart, "ImportedContours", planar);
            }
            else
            {
                MessageBox.Show(
                    $"Failed to import valid contours from:\n{filePath}",
                    "DXF Import",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        private void DXFtoProfile_Execute(object sender, EventArgs e)
        {
            // Assumes a DXF window is currently active
            var profile = DXFImportHelper.DXFtoProfile();
            if (profile == null)
            {
                MessageBox.Show("DXF → Profile failed or was invalid.", "DXF→Profile", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Copy the ProfileString to the clipboard
            Clipboard.SetText(profile.ProfileString);
            MessageBox.Show(
                $"DXF→Profile succeeded.\n\nName = {profile.Name}\n(Profile string copied to clipboard.)",
                "DXF→Profile",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );

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
                        ClientSize = new Size(bmp.Width, bmp.Height)
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
                MessageBox.Show("No DXF profiles available to save.", "Save DXFProfile CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                MessageBox.Show($"Saved {profiles.Count()} profiles to:\n{csvPath}", "Save DXFProfile CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save CSV:\n{ex.Message}", "Save DXFProfile CSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                MessageBox.Show($"Loaded {profiles.Count} profiles from:\n{csvPath}", "Load DXFProfile CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // (Optional) Immediately reconstruct each profile in the main part:
                foreach (var prof in profiles)
                {
                    var body = DXFImportHelper.BodyFromString(prof.ProfileString);
                    DesignBody.Create(Window.ActiveWindow.Document.MainPart, $"Rebuild_{prof.Name}", body);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load CSV:\n{ex.Message}", "Load DXFProfile CSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
