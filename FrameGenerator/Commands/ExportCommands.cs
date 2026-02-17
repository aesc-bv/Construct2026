/*
 ExportCommands collects all BOM and export related functionality for the frame generator.

 It can:
 - Generate and place a BOM table on a drawing sheet in SpaceClaim (ExportBOM).
 - Export the same BOM content to an Excel workbook (ExportExcel).
 - Export all non-main parts as individual STEP files into a chosen folder (ExportSTEP).
 It also contains helpers for cut-angle calculation and unit-aware length formatting.
*/

using AESCConstruct2026.FrameGenerator.Utilities;
using AESCConstruct2026.Properties;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WK.Libraries.BetterFolderBrowserNS;
using Application = SpaceClaim.Api.V242.Application;
using Body = SpaceClaim.Api.V242.Modeler.Body;
using Clipboard = System.Windows.Forms.Clipboard;
using Component = SpaceClaim.Api.V242.Component;
using MessageBox = System.Windows.Forms.MessageBox;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using Table = SpaceClaim.Api.V242.Table;
using Vector = SpaceClaim.Api.V242.Geometry.Vector;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct2026.FrameGenerator.Commands
{
    public static class ExportCommands
    {
        private static readonly Regex TrailingParenRegex = new Regex(@"\(\d+\)$", RegexOptions.Compiled);

        // Builds a BOM from all frame components and writes it as a table onto a drawing sheet (and clipboard text), honoring material/unit settings.
        public static void ExportBOM(Window window, bool update)
        {
            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            bool matSetting = Settings.Default.MatInBOM;
            var unit = GetSelectedUnit();

            // Ensure a drawing sheet exists (offer to create one if none)
            DrawingSheet sheet = doc.DrawingSheets.FirstOrDefault();
            if (sheet == null)
            {
                var choice = MessageBox.Show(
                    "No drawing sheet present. Create a drawing sheet?",
                    "Export BOM",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question
                );

                if (choice == DialogResult.No)
                    return;

                try
                {
                    Command.Execute("NewDrawingSheet");
                    sheet = doc.DrawingSheets.FirstOrDefault();
                }
                catch { }

                if (sheet == null)
                {
                    Application.ReportStatus("Failed to create a drawing sheet.", StatusMessageType.Error, null);
                    return;
                }
            }
            var existing = Window.AllWindows.FirstOrDefault(w => w.Scene == sheet);
            if (existing != null) Window.ActiveWindow = existing;
            else Window.ActiveWindow = Window.Create(sheet);

            // This preserves the placement even if the user dragged the table manually.
            var existingBoms = sheet.GetDescendants<Table>()
                                    .Where(t => t.TryGetTextAttribute("IsExportedBOM", out var tag) && tag == "true")
                                    .ToList();

            PointUV? reuseLocation = null;
            if (existingBoms.Count > 0)
            {
                try
                {
                    // Always use the visual top-left as the anchor we will reuse
                    reuseLocation = existingBoms[0].GetLocation(LocationPoint.TopLeftCorner);
                }
                catch { reuseLocation = null; }

                foreach (var tbl in existingBoms)
                    tbl.Delete();
            }

            // Normalize duplicate names & ensure Construct_Tubelength / Construct_Length are set
            CompareCommand.CompareSimple();

            var comps = doc.MainPart
                           .GetChildren<Component>()
                           .Where(c => c.Template.Bodies.Any())
                           .ToList();

            if (!comps.Any())
            {
                Application.ReportStatus("No components found to export.", StatusMessageType.Information, null);
                return;
            }

            bool anyMaterial = comps.Any(c => c.Template.Material != null && !string.IsNullOrWhiteSpace(c.Template.Material.Name));
            bool includeMaterial = matSetting && anyMaterial;

            var bomRows = comps
                .GroupBy(c => TrailingParenRegex.Replace(c.Template.Name, ""))
                .Select(g =>
                {
                    var first = g.First();

                    // length saved in meters (string)
                    string lengthMeters = "N/A";
                    if (first.Template.CustomProperties.TryGetValue("Construct_Tubelength", out var p) ||
                        first.Template.CustomProperties.TryGetValue("Construct_Length", out p))
                        lengthMeters = p.Value.ToString();

                    string material = first.Template.Material?.Name ?? "";

                    string cuts;
                    try { cuts = GetCutString(first); } catch { cuts = "N/A"; }

                    return new
                    {
                        Part = g.Key,
                        Qty = g.Count(),
                        TubeLengthMeters = lengthMeters,
                        CutAngles = cuts,
                        Material = material
                    };
                })
                .OrderBy(x => x.Part)
                .ToList();

            int baseCols = 4;
            int cols = includeMaterial ? baseCols + 1 : baseCols;
            var sb = new StringBuilder();

            // Header with selected unit
            sb.Append("Part Name").Append('\t')
              .Append("Qty").Append('\t')
              .Append($"Tube Length ({unit})").Append('\t')
              .Append("Cut Angles");
            if (includeMaterial) sb.Append('\t').Append("Material");
            sb.Append("\r\n");

            foreach (var row in bomRows)
            {
                sb.Append(row.Part).Append('\t')
                  .Append(row.Qty).Append('\t')
                  .Append(FormatLengthFromMetersString(row.TubeLengthMeters)).Append('\t')
                  .Append(row.CutAngles);
                if (includeMaterial)
                    sb.Append('\t').Append(string.IsNullOrWhiteSpace(row.Material) ? "" : row.Material);
                sb.Append("\r\n");
            }

            var payload = sb.ToString();

            try { Clipboard.SetText(payload); }
            catch (Exception ex)
            {
                Application.ReportStatus("Failed to copy BOM to clipboard:\n" + ex.Message, StatusMessageType.Error, null);
                return;
            }

            try
            {
                int rows = bomRows.Count + 1;
                var contents = new string[rows, cols];

                contents[0, 0] = "Part Name";
                contents[0, 1] = "Qty";
                contents[0, 2] = $"Tube Length ({unit})";
                contents[0, 3] = "Cut Angles";
                if (includeMaterial) contents[0, 4] = "Material";

                for (int i = 0; i < bomRows.Count; i++)
                {
                    contents[i + 1, 0] = bomRows[i].Part;
                    contents[i + 1, 1] = bomRows[i].Qty.ToString();
                    contents[i + 1, 2] = FormatLengthFromMetersString(bomRows[i].TubeLengthMeters);
                    contents[i + 1, 3] = bomRows[i].CutAngles;
                    if (includeMaterial)
                        contents[i + 1, 4] = string.IsNullOrWhiteSpace(bomRows[i].Material) ? "" : bomRows[i].Material;
                }

                double rowHeight = 0.008;
                double width = 0.30;
                double columnWidth = width / cols;
                double fontSize = rowHeight * 0.4;

                // Decide placement:
                //  - If we captured a previous table's location, reuse it and anchor by TopLeftCorner
                //  - Otherwise, compute from DocumentAnchor + AnchorX/Y settings
                PointUV anchorUV;
                LocationPoint corner;

                if (reuseLocation.HasValue)
                {
                    anchorUV = reuseLocation.Value;
                    corner = LocationPoint.TopLeftCorner;
                }
                else
                {
                    bool left = (Settings.Default.DocumentAnchor == "TopLeft" || Settings.Default.DocumentAnchor == "BottomLeft");
                    bool top = (Settings.Default.DocumentAnchor == "TopLeft" || Settings.Default.DocumentAnchor == "TopRight");

                    double anchorX = Settings.Default.TableAnchorX / 1000.0; // mm → m
                    if (!left)
                        anchorX = sheet.Width - (Settings.Default.TableAnchorX / 1000.0);

                    double anchorY = Settings.Default.TableAnchorY / 1000.0; // mm → m
                    if (top)
                        anchorY = sheet.Height - (Settings.Default.TableAnchorY / 1000.0);

                    anchorUV = new PointUV(anchorX, anchorY);
                    corner = (LocationPoint)Enum.Parse(typeof(LocationPoint), Settings.Default.TableLocationPoint);
                }

                var table = Table.Create(sheet, anchorUV, corner, rowHeight, columnWidth, fontSize, contents);
                table.SetTextAttribute("IsExportedBOM", "true");

                double[] columnWidths = includeMaterial
                    ? new[] { 0.06, 0.015, 0.04, 0.05, 0.08 }
                    : new[] { 0.08, 0.02, 0.06, 0.06 };

                for (int colIndex = 0; colIndex < columnWidths.Length && colIndex < table.Columns.Count; colIndex++)
                    table.Columns[colIndex].Width = columnWidths[colIndex];

                Application.ReportStatus("BOM table has been written to the active drawing sheet.", StatusMessageType.Information, null);
            }
            catch (Exception ex)
            {
                Application.ReportStatus("Failed to build BOM table:\n" + ex.Message, StatusMessageType.Error, null);
            }
        }

        // Exports the BOM to an .xlsx file on disk using OpenXML, then opens the file in the default spreadsheet application.
        public static void ExportExcel(Window window)
        {
            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                Application.ReportStatus("No active document.", StatusMessageType.Warning, null);
                return;
            }

            bool matExcelSetting = Settings.Default.MatInExcel;
            var unit = GetSelectedUnit();

            CompareCommand.CompareSimple();

            var comps = doc.MainPart
                           .GetChildren<SpaceClaim.Api.V242.Component>()
                           .Where(c => c.Template.Bodies.Any())
                           .ToList();
            if (!comps.Any())
            {
                Application.ReportStatus("No components found to export.", StatusMessageType.Information, null);
                return;
            }

            bool anyMaterial = comps.Any(c => c.Template.Material != null && !string.IsNullOrWhiteSpace(c.Template.Material.Name));
            bool includeMaterial = matExcelSetting && anyMaterial;

            var bomRows = comps
                .GroupBy(c => TrailingParenRegex.Replace(c.Template.Name, ""))
                .Select(g =>
                {
                    var first = g.First();

                    // length saved in meters
                    string lengthMeters = "N/A";
                    if (first.Template.CustomProperties.TryGetValue("Construct_Tubelength", out var p) ||
                        first.Template.CustomProperties.TryGetValue("Construct_Length", out p))
                        lengthMeters = p.Value.ToString();

                    string material = first.Template.Material?.Name ?? "";

                    string cuts;
                    try { cuts = GetCutString(first); } catch { cuts = "N/A"; }

                    return new
                    {
                        Part = g.Key,
                        Qty = g.Count(),
                        TubeLengthMeters = lengthMeters,
                        CutAngles = cuts,
                        Material = material
                    };
                })
                .OrderBy(x => x.Part)
                .ToList();

            // Ask for file name + location (like ExportSettingsButton_Click)
            var dlg = new SaveFileDialog
            {
                Title = "Export BOM to Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FileName = "AESC_Construct2026_BOM.xlsx",
                AddExtension = true,
                DefaultExt = ".xlsx",
                OverwritePrompt = true
            };
            if (dlg.ShowDialog() != true)
                return;

            var excelPath = dlg.FileName;

            try
            {
                using (var docX = SpreadsheetDocument.Create(excelPath, SpreadsheetDocumentType.Workbook))
                {
                    var wbPart = docX.AddWorkbookPart();
                    wbPart.Workbook = new Workbook();

                    var sheetPart = wbPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    // Header row
                    var header = new Row { RowIndex = 1 };
                    header.Append(
                        MakeTextCell("A", 1, "Part Name"),
                        MakeTextCell("B", 1, "Quantity"),
                        MakeTextCell("C", 1, $"Tube Length ({unit})"),
                        MakeTextCell("D", 1, "Cut Angles")
                    );
                    if (includeMaterial)
                        header.Append(MakeTextCell("E", 1, "Material"));
                    sheetData.Append(header);

                    // Data rows
                    for (int i = 0; i < bomRows.Count; i++)
                    {
                        uint rowIndex = (uint)(i + 2);
                        var row = new Row { RowIndex = rowIndex };
                        row.Append(
                            MakeTextCell("A", rowIndex, bomRows[i].Part),
                            MakeNumberCell("B", rowIndex, bomRows[i].Qty),
                            MakeTextCell("C", rowIndex, FormatLengthFromMetersString(bomRows[i].TubeLengthMeters)),
                            MakeTextCell("D", rowIndex, bomRows[i].CutAngles)
                        );
                        if (includeMaterial)
                            row.Append(MakeTextCell("E", rowIndex, string.IsNullOrWhiteSpace(bomRows[i].Material) ? "" : bomRows[i].Material));
                        sheetData.Append(row);
                    }

                    var sheets = wbPart.Workbook.AppendChild(new Sheets());
                    sheets.Append(new Sheet
                    {
                        Id = wbPart.GetIdOfPart(sheetPart),
                        SheetId = 1,
                        Name = "BOM"
                    });

                    wbPart.Workbook.Save();
                }

                Application.ReportStatus("BOM exported:\n" + excelPath, StatusMessageType.Information, null);

                // Open the file
                Process.Start(new ProcessStartInfo(excelPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to export to Excel:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Exports each non-main Part as an individual STEP file to a user-selected folder, optionally appending material to the filename.
        public static void ExportSTEP(Window window)
        {
            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                Application.ReportStatus("No active document.", StatusMessageType.Error, null);
                return;
            }

            bool matSTEP = Settings.Default.MatInSTEP;

            // 1) Gather all parts (excluding the main assembly) that have at least one body
            var partsToExport = doc.Parts
                .Where(p => p != doc.MainPart && p.Bodies.Any())
                .ToList();
            if (!partsToExport.Any())
            {
                Application.ReportStatus("No parts found to export.", StatusMessageType.Warning, null);
                return;
            }

            // 2) Ask user where to save the STEP files
            string folderPath;
            using (var bfb = new BetterFolderBrowser(new Container()))
            {
                bfb.RootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                bfb.Title = "Select folder to export STEP files";
                if (bfb.ShowDialog() != DialogResult.OK)
                    return;
                folderPath = bfb.SelectedFolder;
            }

            try
            {
                WriteBlock.ExecuteTask("Convert DXF Profile", () =>
                {
                    // 3) Loop through each part and export to a uniquely named .stp
                    foreach (var part in partsToExport)
                    {
                        // base name from display name
                        var name = part.DisplayName;

                        // if checkbox set, append material (underscored, no spaces)
                        if (matSTEP && part.Material != null)
                        {
                            var mat = part.Material.Name.Replace(" ", "_");
                            name = $"{name}_{mat}";
                        }

                        var baseFileName = $"{name}.stp";
                        var dest = Path.Combine(folderPath, baseFileName);

                        // If file exists, append a numeric suffix
                        if (File.Exists(dest))
                        {
                            int suffix = 1;
                            string candidate;
                            do
                            {
                                candidate = Path.Combine(folderPath, $"{name} ({suffix}).stp");
                                suffix++;
                            }
                            while (File.Exists(candidate) && suffix < 1000);

                            dest = candidate;
                        }

                        try
                        {
                            part.Export(PartExportFormat.Step, dest, false, null);
                        }
                        catch
                        {
                        }
                    }

                    // 4) Open the export folder in Explorer
                    Process.Start(new ProcessStartInfo(folderPath)
                    {
                        UseShellExecute = true,
                        Verb = "open"
                    });
                });
            }
            catch (Exception ex)
            {
                Application.ReportStatus($"Failed to export STEP files:\n{ex.Message}", StatusMessageType.Error, null);
            }
        }

        // Creates an OpenXML text cell at (col,rowIndex) with string content for Excel export.
        static Cell MakeTextCell(string col, uint rowIndex, string text)
        {
            return new Cell
            {
                CellReference = col + rowIndex,
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };
        }

        // Creates an OpenXML numeric cell at (col,rowIndex) with an integer value for Excel export.
        static Cell MakeNumberCell(string col, uint rowIndex, int number)
        {
            return new Cell
            {
                CellReference = col + rowIndex,
                DataType = CellValues.Number,
                CellValue = new CellValue(number.ToString())
            };
        }

        // Computes the start/end miter and bevel angles (X/Z) on the profile's end faces for a given component.
        public static (double xStart, double zStart, double xEnd, double zEnd) GetProfileCutAngles(Component comp)
        {
            const double tol = 1e-6;

            // — 1) Extract sweep direction and end-cap normals —
            var dc = comp.Template.Curves
                         .OfType<DesignCurve>()
                         .FirstOrDefault()
                     ?? throw new InvalidOperationException("No construction curve");
            var seg = (CurveSegment)dc.Shape;
            Vector sweepLocal = (seg.EndPoint - seg.StartPoint).Direction.ToVector();

            var body = comp.Template.Bodies
                          .FirstOrDefault(b => b.Name == "ExtrudedProfile")
                      ?? throw new InvalidOperationException("No ExtrudedProfile");
            var profileBody = (Body)body.Shape;

            var endCaps = profileBody.Faces
                .Select(f => f.Geometry as Plane)
                .Where(pl => pl != null)
                .Select(pl => pl.Frame.DirZ.ToVector())
                .Where(n => Math.Abs(Vector.Dot(n, sweepLocal)) > tol)
                .ToList();

            if (endCaps.Count != 2)
                throw new InvalidOperationException($"Expected 2 end-cap faces but found {endCaps.Count}");

            var n0 = endCaps.OrderByDescending(n => Vector.Dot(n, -sweepLocal)).First();
            var n1 = endCaps.OrderByDescending(n => Vector.Dot(n, sweepLocal)).First();

            // — 2) Compute signed X-cut: 90°–|clocking|, with sign of clocking ▷ in [–90,90]
            double ComputeXcut(Vector n)
            {
                double ang = Math.Atan2(n.Y, n.X) * 180.0 / Math.PI;
                if (ang > 90.0) ang -= 180.0;
                if (ang < -90.0) ang += 180.0;
                if (Math.Abs(ang) < tol)
                    return 0.0;
                double cut = (Math.Abs(ang) % 90);
                double signedCut = cut * Math.Sign(ang);
                return Math.Round(signedCut, 1);
            }

            // — 3) Compute signed Z-cut: 90°–bevel, preserving sign (bevel>0 if normal tilted up, <0 if down)
            double ComputeZcut(Vector n)
            {
                double raw = Math.Acos(Math.Max(-1.0, Math.Min(1.0, n.Z))) * 180.0 / Math.PI;
                double cut = Math.Abs(raw) % 90;
                if (Math.Abs(n.Z - 1.0) < tol || Math.Abs(n.Z + 1.0) < tol)
                    cut = 0.0;
                return Math.Round(cut, 1);
            }

            double x0 = ComputeXcut(n0), z0 = ComputeZcut(n0);
            double x1 = ComputeXcut(n1), z1 = ComputeZcut(n1);

            return (x0, z0, x1, z1);
        }

        // Returns a formatted user-facing string with both end cut angles in "X/Z" notation, or "ERR" on failure.
        public static string GetCutString(Component comp)
        {
            try
            {
                var (x0, z0, x1, z1) = GetProfileCutAngles(comp);
                string result = $"X: {x0:F1}/Z: {z0:F1}, X: {x1:F1}/Z: {z1:F1}";
                return result;
            }
            catch (Exception)
            {
                return "ERR";
            }
        }

        // Returns the currently configured length unit from settings (defaults to "mm" if unset).
        static string GetSelectedUnit()
        {
            var u = Settings.Default.LengthUnit;
            return string.IsNullOrWhiteSpace(u) ? "mm" : u;
        }

        // Parses a length string that is stored as raw meters into a double value (culture tolerant).
        static bool TryParseMeters(string text, out double meters)
        {
            meters = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            // stored as raw numeric meters (e.g., "0.06")
            // be tolerant of culture
            if (double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ||
                double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out v))
            {
                meters = v;
                return true;
            }
            return false;
        }

        // Converts a length in meters to the configured output unit (mm/cm/m/inch) for BOM display.
        static double FromMeters(double meters, string unit)
        {
            switch (unit)
            {
                case "mm": return meters * 1000.0;
                case "cm": return meters * 100.0;
                case "m": return meters;
                case "inch": return meters / 0.0254;
                default: return meters;
            }
        }

        // Formats a numeric meters string into the selected unit with a fixed numeric format for BOM/Excel output.
        static string FormatLengthFromMetersString(string rawMeters)
        {
            if (TryParseMeters(rawMeters, out var m))
            {
                var unit = GetSelectedUnit();
                var val = FromMeters(m, unit);
                return val.ToString("0.###", CultureInfo.InvariantCulture);
            }
            return rawMeters;
        }
    }
}
