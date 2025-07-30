using AESCConstruct25.FrameGenerator.Utilities;
using AESCConstruct25.Properties;     // for Settings.Default
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WK.Libraries.BetterFolderBrowserNS;
using Body = SpaceClaim.Api.V242.Modeler.Body;
using Component = SpaceClaim.Api.V242.Component;
using Table = SpaceClaim.Api.V242.Table;
//using Component = SpaceClaim.Api.V242.Component;

namespace AESCConstruct25.FrameGenerator.Commands
{
    public static class ExportCommands
    {
        public static void ExportBOM(Window window, bool update)
        {
            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                MessageBox.Show("No active document.", "Export BOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool matBOM = Settings.Default.MatInBOM;

            DrawingSheet sheet = doc.DrawingSheets.FirstOrDefault();
            if (update)
            {
                if (sheet == null)
                {
                    MessageBox.Show(
                        "No drawing sheet open; skipping BOM update.",
                        "Export BOM",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }
                else
                {
                    var old = sheet
                        .GetDescendants<Table>()
                        .Where(t => t.TryGetTextAttribute("IsExportedBOM", out var tag) && tag == "true")
                        .ToList();
                    Logger.Log($"Deleting {old.Count} existing BOM table(s).");
                    foreach (var tbl in old)
                        tbl.Delete();
                }
            }

            CompareCommand.CompareSimple();

            var comps = doc.MainPart
                           .GetChildren<SpaceClaim.Api.V242.Component>()
                           .Where(c => c.Template.Bodies.Any())
                           .ToList();
            if (!comps.Any())
            {
                MessageBox.Show("No components found to export.", "Export BOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 2) Group into BOM rows
            var bomRows = comps
                .GroupBy(c => Regex.Replace(c.Template.Name, @"\(\d+\)$", ""))
                .Select(g =>
                {
                    var first = g.First();
                    string length = "N/A";
                    if (first.Template.CustomProperties.TryGetValue("Construct_Tubelength", out var p))
                        length = p.Value.ToString();

                    string material = "N/A";
                    if (matBOM && first.Template.Material != null)
                        material = first.Template.Material.Name;

                    string cuts;
                    try { cuts = GetCutString(first); }
                    catch (Exception ex)
                    {
                        Logger.Log($"ExportCommands: failed to get cuts for {first.Name}: {ex.Message}");
                        cuts = "N/A";
                    }

                    return new
                    {
                        Part = g.Key,
                        Qty = g.Count(),
                        TubeLength = length,
                        CutAngles = cuts,
                        Material = material
                    };
                })
                .OrderBy(x => x.Part)
                .ToList();

            // 3) Build tab-delimited payload
            int baseCols = 4;
            int cols = matBOM ? baseCols + 1 : baseCols;
            var sb = new StringBuilder();

            // Header row
            sb.Append("Part Name").Append('\t')
              .Append("Qty").Append('\t')
              .Append("Tube Length").Append('\t')
              .Append("Cut Angles");
            if (matBOM)
                sb.Append('\t').Append("Material");
            sb.Append("\r\n");

            // Data rows
            foreach (var row in bomRows)
            {
                Logger.Log($"Export row: {row.Part} → material={row.Material}");
                sb.Append(row.Part).Append('\t')
                  .Append(row.Qty).Append('\t')
                  .Append(row.TubeLength).Append('\t')
                  .Append(row.CutAngles);
                if (matBOM)
                    sb.Append('\t').Append(row.Material);
                sb.Append("\r\n");
            }

            var payload = sb.ToString();

            // 4) Copy to clipboard
            try { Clipboard.SetText(payload); }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to copy BOM to clipboard:\n" + ex.Message, "Export BOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 5) Paste into drawing sheet
            try
            {
                if (sheet == null)
                {
                    MessageBox.Show(
                        "Please activate a drawing sheet before exporting the BOM.",
                        "Export BOM",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                int rows = bomRows.Count + 1;
                var contents = new string[rows, cols];

                // Header
                contents[0, 0] = "Part Name";
                contents[0, 1] = "Qty";
                contents[0, 2] = "Tube Length";
                contents[0, 3] = "Cut Angles";
                if (matBOM)
                    contents[0, 4] = "Material";

                // Data
                for (int i = 0; i < bomRows.Count; i++)
                {
                    contents[i + 1, 0] = bomRows[i].Part;
                    contents[i + 1, 1] = bomRows[i].Qty.ToString();
                    contents[i + 1, 2] = bomRows[i].TubeLength;
                    contents[i + 1, 3] = bomRows[i].CutAngles;
                    if (matBOM)
                        contents[i + 1, 4] = bomRows[i].Material;
                }

                double margin = 0.02;
                double rowHeight = 0.008;
                double width = 0.30;
                double columnWidth = width / cols;
                double fontSize = rowHeight * 0.4;

                double anchorX = Settings.Default.TableAnchorX;
                double anchorY = Settings.Default.TableAnchorY;
                var anchorUV = new PointUV(anchorX, anchorY);

                var corner = (LocationPoint)Enum.Parse(
                    typeof(LocationPoint),
                    Settings.Default.TableLocationPoint
                );

                var table = Table.Create(
                    sheet,
                    anchorUV,
                    corner,  //LocationPoint.BottomRightCorner,
                    rowHeight,
                    columnWidth,
                    fontSize,
                    contents
                );

                table.SetTextAttribute("IsExportedBOM", "true");

                // Adjust column widths dynamically
                double[] columnWidths = matBOM
                    ? new[] { 0.06, 0.015, 0.03, 0.04, 0.08 }
                    : new[] { 0.06, 0.015, 0.03, 0.04 };

                for (int colIndex = 0; colIndex < columnWidths.Length && colIndex < table.Columns.Count; colIndex++)
                    table.Columns[colIndex].Width = columnWidths[colIndex];

                MessageBox.Show(
                    "BOM table has been written to the active drawing sheet.",
                    "Export BOM",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to build BOM table:\n" + ex.Message, "Export BOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public static void ExportExcel(Window window)
        {
            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                MessageBox.Show("No active document.", "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool matExcel = Settings.Default.MatInExcel;

            // normalize duplicate names & ensure Construct_Tubelength is set
            CompareCommand.CompareSimple();

            // gather components
            var comps = doc.MainPart
                           .GetChildren<SpaceClaim.Api.V242.Component>()
                           .Where(c => c.Template.Bodies.Any())
                           .ToList();
            if (!comps.Any())
            {
                MessageBox.Show("No components found to export.", "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // group by cleaned name and pull tube length + material if requested
            var bomRows = comps
              .GroupBy(c => Regex.Replace(c.Template.Name, @"\(\d+\)$", ""))
              .Select(g =>
              {
                  var first = g.First();

                  string length = "N/A";
                  if (first.Template.CustomProperties.TryGetValue("Construct_Tubelength", out var p))
                      length = p.Value.ToString();

                  string material = "N/A";
                  if (matExcel && first.Template.Material != null)
                      material = first.Template.Material.Name;

                  string cuts;
                  try
                  {
                      cuts = GetCutString(first);
                  }
                  catch (Exception ex)
                  {
                      Logger.Log($"ExportCommands: failed to get cuts for {first.Name}: {ex.Message}");
                      cuts = "N/A";
                  }

                  return new
                  {
                      Part = g.Key,
                      Qty = g.Count(),
                      TubeLength = length,
                      CutAngles = cuts,
                      Material = material
                  };
              })
              .OrderBy(x => x.Part)
              .ToList();

            // ask for folder
            string folderPath;
            using (var bfb = new BetterFolderBrowser(new Container()))
            {
                bfb.Title = "Select folder to save BOM Excel";
                bfb.RootFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (bfb.ShowDialog() != DialogResult.OK)
                    return;
                folderPath = bfb.SelectedFolder;
            }

            var excelPath = Path.Combine(folderPath, "BOM.xlsx");
            try
            {
                using (var docX = SpreadsheetDocument.Create(excelPath, SpreadsheetDocumentType.Workbook))
                {
                    var wbPart = docX.AddWorkbookPart();
                    wbPart.Workbook = new Workbook();
                    var sheetPart = wbPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    // header row
                    var header = new Row { RowIndex = 1 };
                    header.Append(
                        MakeTextCell("A", 1, "Part Name"),
                        MakeTextCell("B", 1, "Quantity"),
                        MakeTextCell("C", 1, "Tube Length"),
                        MakeTextCell("D", 1, "Cut Angles")
                    );
                    if (matExcel)
                        header.Append(MakeTextCell("E", 1, "Material"));
                    sheetData.Append(header);

                    // data rows
                    for (int i = 0; i < bomRows.Count; i++)
                    {
                        uint rowIndex = (uint)(i + 2);
                        var row = new Row { RowIndex = rowIndex };
                        row.Append(
                            MakeTextCell("A", rowIndex, bomRows[i].Part),
                            MakeNumberCell("B", rowIndex, bomRows[i].Qty),
                            MakeTextCell("C", rowIndex, bomRows[i].TubeLength),
                            MakeTextCell("D", rowIndex, bomRows[i].CutAngles)
                        );
                        if (matExcel)
                            row.Append(MakeTextCell("E", rowIndex, bomRows[i].Material));
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

                Process.Start(new ProcessStartInfo(excelPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to export to Excel:\n{ex.Message}", "Export to Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void ExportSTEP(Window window)
        {
            var doc = window?.Document;
            if (doc?.MainPart == null)
            {
                MessageBox.Show("No active document.", "Export to STEP",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool matSTEP = Settings.Default.MatInSTEP;

            // 1) Gather all parts (excluding the main assembly) that have at least one body
            var partsToExport = doc.Parts
                .Where(p => p != doc.MainPart && p.Bodies.Any())
                .ToList();
            if (!partsToExport.Any())
            {
                MessageBox.Show("No parts found to export.", "Export to STEP",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                        // Swallow export errors for individual parts
                    }
                }

                // 4) Open the export folder in Explorer
                Process.Start(new ProcessStartInfo(folderPath)
                {
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                  $"Failed to export STEP files:\n{ex.Message}",
                  "Export to STEP",
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Error
                );
            }
        }

        // Helper: create a text cell
        static Cell MakeTextCell(string col, uint rowIndex, string text)
        {
            return new Cell
            {
                CellReference = col + rowIndex,
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };
        }

        // Helper: create a numeric cell
        static Cell MakeNumberCell(string col, uint rowIndex, int number)
        {
            return new Cell
            {
                CellReference = col + rowIndex,
                DataType = CellValues.Number,
                CellValue = new CellValue(number.ToString())
            };
        }

        //public static (double aStart, double aEnd) GetProfileCutAngles(Component comp)
        //{
        //    // — 1) Extract the local sweep direction from the design curve —
        //    var dc = comp.Template.Curves
        //                 .OfType<DesignCurve>()
        //                 .FirstOrDefault()
        //             ?? throw new InvalidOperationException("No construction curve");
        //    var seg = (CurveSegment)dc.Shape;
        //    Vector sweepLocal = (seg.EndPoint - seg.StartPoint).Direction.ToVector();

        //    // — 2) Pull out only the planar faces of the extruded profile body —
        //    var body = comp.Template.Bodies
        //                  .FirstOrDefault(b => b.Name == "ExtrudedProfile")
        //              ?? throw new InvalidOperationException("No ExtrudedProfile");
        //    var profileBody = (Body)body.Shape;

        //    const double tol = 1e-6;
        //    var endCaps = profileBody.Faces
        //      .Select(f => f.Geometry as Plane)               // cast to Plane
        //      .Where(pl => pl != null)
        //      .Select(pl => pl.Frame.DirZ.ToVector())         // face normal in local
        //      .Where(n => Math.Abs(Vector.Dot(n, sweepLocal)) > tol)
        //      .ToList();

        //    if (endCaps.Count != 2)
        //        throw new InvalidOperationException(
        //          $"Expected 2 end-cap faces but found {endCaps.Count}"
        //        );

        //    // — 3) Identify start vs end by alignment with ±sweepLocal —
        //    var n0 = endCaps.OrderByDescending(n => Vector.Dot(n, -sweepLocal)).First();
        //    var n1 = endCaps.OrderByDescending(n => Vector.Dot(n, sweepLocal)).First();

        //    // — 4) “Up” axis in local (Z)
        //    Vector localUp = Vector.Create(0, 0, 1);

        //    // — 5) Measure deviation from 90°, **round to 0.1°** —
        //    double Clamp01(double x) => Math.Max(-1.0, Math.Min(1.0, x));
        //    double Measure(Vector normal)
        //    {
        //        // normal·up = cos(deviation), 1→0°, 0→90°
        //        var d = Clamp01(Vector.Dot(normal, localUp));
        //        double angle = Math.Acos(d) * 180.0 / Math.PI;
        //        return Math.Round(angle, 1);     // <-- 0.1° precision
        //    }

        //    return (Measure(n0), Measure(n1));
        //}

        //public static string GetCutString(Component comp)
        //{
        //    try
        //    {
        //        var (a0, a1) = GetProfileCutAngles(comp);
        //        return $"{a0:F1}/{a1:F1}";
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.Log($"GetCutString: FAILED on '{comp.Name}': {ex.Message}");
        //        return "ERR";
        //    }
        //}
        public static (double xStart, double zStart, double xEnd, double zEnd) GetProfileCutAngles(Component comp)
        {
            const double tol = 1e-6;
            Logger.Log($"GetProfileCutAngles: component='{comp.Name}'");

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

            Logger.Log($"  normals: n0=({n0.X:F3},{n0.Y:F3},{n0.Z:F3}), n1=({n1.X:F3},{n1.Y:F3},{n1.Z:F3})");

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
            Logger.Log($"  start cut X={x0:F1}, Z={z0:F1}");
            double x1 = ComputeXcut(n1), z1 = ComputeZcut(n1);
            Logger.Log($"  end   cut X={x1:F1}, Z={z1:F1}");

            return (x0, z0, x1, z1);
        }

        public static string GetCutString(Component comp)
        {
            try
            {
                var (x0, z0, x1, z1) = GetProfileCutAngles(comp);
                string result = $"X: {x0:F1}/Z: {z0:F1}, X: {x1:F1}/Z: {z1:F1}";
                Logger.Log($"GetCutString: {result}");
                return result;
            }
            catch (Exception ex)
            {
                Logger.Log($"GetCutString: FAILED on '{comp.Name}': {ex.Message}");
                return "ERR";
            }
        }

    }
}