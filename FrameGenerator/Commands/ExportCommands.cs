using System;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.ComponentModel;
using System.Windows.Forms;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using Component = SpaceClaim.Api.V242.Component;
using Body = SpaceClaim.Api.V242.Modeler.Body;
using Table = SpaceClaim.Api.V242.Table;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;
using WK.Libraries.BetterFolderBrowserNS;
using AESCConstruct25.FrameGenerator.Utilities;
using System.Text.RegularExpressions;
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

            DrawingSheet sheet = doc.DrawingSheets.FirstOrDefault();
            // if we're updating, delete any old BOM tables first
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

            // normalize duplicate names & ensure Construct_Tubelength is set
            CompareCommand.CompareSimple();

            // 1) Gather all components with bodies
            var comps = doc.MainPart
                           .GetChildren<SpaceClaim.Api.V242.Component>()
                           .Where(c => c.Template.Bodies.Any())
                           .ToList();
            if (!comps.Any())
            {
                MessageBox.Show("No components found to export.", "Export BOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 2) Group into BOM rows by cleaned name, grabbing tube length from the first in each group
            var bomRows = comps
            .GroupBy(c => Regex.Replace(c.Template.Name, @"\(\d+\)$", ""))
            .Select(g => {
                var first = g.First();
                // tube length is as before…
                string length = "N/A";
                if (first.Template.CustomProperties.TryGetValue("Construct_Tubelength", out var p))
                    length = p.Value.ToString();

                // now defensively compute cut-angles:
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
                    CutAngles = cuts
                };
            })
            .OrderBy(x => x.Part)
            .ToList();

            // 3) Build tab-delimited string with three columns
            var sb = new StringBuilder();
            sb.Append("Part Name").Append('\t')
              .Append("Qty").Append('\t')
              .Append("Tube Length").Append('\t')
              .Append("Cut Angles").Append("\r\n");
            foreach (var row in bomRows)
            {
                Logger.Log($"Export row: {row.Part} → cuts={row.CutAngles}");
                sb.Append(row.Part).Append('\t')
                  .Append(row.Qty).Append('\t')
                  .Append(row.TubeLength).Append('\t')
                  .Append(row.CutAngles).Append("\r\n");
            }
            var payload = sb.ToString();

            // 4) Copy to clipboard
            try { Clipboard.SetText(payload); }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to copy BOM to clipboard:\n" + ex.Message, "Export BOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 5) Paste into active drawing sheet
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

                int rows = bomRows.Count + 1, cols = 4;
                var contents = new string[rows, cols];
                contents[0, 0] = "Part Name";
                contents[0, 1] = "Qty";
                contents[0, 2] = "Tube Length";
                contents[0, 3] = "Cut Angles";
                for (int i = 0; i < bomRows.Count; i++)
                {
                    contents[i + 1, 0] = bomRows[i].Part;
                    contents[i + 1, 1] = bomRows[i].Qty.ToString();
                    contents[i + 1, 2] = bomRows[i].TubeLength;
                    contents[i + 1, 3] = bomRows[i].CutAngles;
                }


                // margins (in model units)
                double margin = 0.02;
                double rowHeight = 0.008;
                double width = 0.30;            // total table width
                double columnWidth = width / cols;
                double fontSize = rowHeight * 0.4;

                // 1) Compute the anchor UV so that the table’s bottom-right corner sits
                //    margin units in from the sheet’s right edge and margin units up from bottom
                double anchorX = sheet.Width - margin;
                double anchorY = margin * 4.5;
                var anchorUV = new PointUV(anchorX, anchorY);

                // 2) Create the table with BottomRightCorner as the alignment point
                var table = Table.Create(
                    sheet,
                    anchorUV,
                    LocationPoint.BottomRightCorner,
                    rowHeight,
                    columnWidth,
                    fontSize,
                    contents
                );

                // 3) Tag it for future updates
                table.SetTextAttribute("IsExportedBOM", "true");

                // 4) Adjust column widths as before
                double[] columnWidths = { 0.06, 0.01, 0.03, 0.04 };
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

            // group by cleaned name and pull tube length
            var bomRows = comps
              .GroupBy(c => Regex.Replace(c.Template.Name, @"\(\d+\)$", ""))
              .Select(g => {
                  var first = g.First();
                  // tube length is as before…
                  string length = "N/A";
                  if (first.Template.CustomProperties.TryGetValue("Construct_Tubelength", out var p))
                      length = p.Value.ToString();

                  // now defensively compute cut-angles:
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
                      CutAngles = cuts
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
                    var header = new Row();
                    header.Append(
                        MakeTextCell("A", 1, "Part Name"),
                        MakeTextCell("B", 1, "Quantity"),
                        MakeTextCell("C", 1, "Tube Length"),
                        MakeTextCell("D", 1, "Cut Angles")
                    );
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
                // 1) Seed the initial path
                bfb.RootFolder = Environment.GetFolderPath(
                    Environment.SpecialFolder.MyDocuments
                );

                // 2) Set your title
                bfb.Title = "Select folder to export STEP files";

                // 3) Show it
                if (bfb.ShowDialog() != DialogResult.OK)
                    return;

                // 4) Read back the picked folder
                folderPath = bfb.SelectedFolder;
            }

            try
            {
                // 3) Loop through each part and export to a uniquely named .stp
                foreach (var part in partsToExport)
                {
                    var name = part.DisplayName;
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

        public static (double aStart, double aEnd) GetProfileCutAngles(Component comp)
        {
            // — 1) Extract the local sweep direction from the design curve —
            var dc = comp.Template.Curves
                         .OfType<DesignCurve>()
                         .FirstOrDefault()
                     ?? throw new InvalidOperationException("No construction curve");
            var seg = (CurveSegment)dc.Shape;
            Vector sweepLocal = (seg.EndPoint - seg.StartPoint).Direction.ToVector();

            // — 2) Pull out only the planar faces of the extruded profile body —
            var body = comp.Template.Bodies
                          .FirstOrDefault(b => b.Name == "ExtrudedProfile")
                      ?? throw new InvalidOperationException("No ExtrudedProfile");
            var profileBody = (Body)body.Shape;

            const double tol = 1e-6;
            var endCaps = profileBody.Faces
              .Select(f => f.Geometry as Plane)               // cast to Plane
              .Where(pl => pl != null)
              .Select(pl => pl.Frame.DirZ.ToVector())         // face normal in local
              .Where(n => Math.Abs(Vector.Dot(n, sweepLocal)) > tol)
              .ToList();

            if (endCaps.Count != 2)
                throw new InvalidOperationException(
                  $"Expected 2 end-cap faces but found {endCaps.Count}"
                );

            // — 3) Identify start vs end by alignment with ±sweepLocal —
            var n0 = endCaps.OrderByDescending(n => Vector.Dot(n, -sweepLocal)).First();
            var n1 = endCaps.OrderByDescending(n => Vector.Dot(n, sweepLocal)).First();

            // — 4) “Up” axis in local (Z)
            Vector localUp = Vector.Create(0, 0, 1);

            // — 5) Measure deviation from 90°, **round to 0.1°** —
            double Clamp01(double x) => Math.Max(-1.0, Math.Min(1.0, x));
            double Measure(Vector normal)
            {
                // normal·up = cos(deviation), 1→0°, 0→90°
                var d = Clamp01(Vector.Dot(normal, localUp));
                double angle = Math.Acos(d) * 180.0 / Math.PI;
                return Math.Round(angle, 1);     // <-- 0.1° precision
            }

            return (Measure(n0), Measure(n1));
        }

        public static string GetCutString(Component comp)
        {
            try
            {
                var (a0, a1) = GetProfileCutAngles(comp);
                return $"{a0:F1}/{a1:F1}";
            }
            catch (Exception ex)
            {
                Logger.Log($"GetCutString: FAILED on '{comp.Name}': {ex.Message}");
                return "ERR";
            }
        }

    }
}