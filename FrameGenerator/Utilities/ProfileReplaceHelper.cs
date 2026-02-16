/*
 ProfileReplaceHelper handles replacing Construct-generated profile components with their original driving curves.
 It detects selected Construct components, prompts the user, unhides matching original lines, updates selection, and deletes the component instances.
*/

using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Application = SpaceClaim.Api.V242.Application;
using Document = SpaceClaim.Api.V242.Document;
using Point = SpaceClaim.Api.V242.Geometry.Point;
using Window = SpaceClaim.Api.V242.Window;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    public static class ProfileReplaceHelper
    {
        // Entry point: optionally restores original curves for selected Construct profile components and deletes the components.
        public static (bool proceed, List<DesignCurve> originals)
            PromptAndReplaceIfAny(Window win, Action<IEnumerable<Component>> deleteAction = null)
        {
            var comps = FindConstructComponentsFromSelection(win);
            if (comps.Count == 0)
                return (true, new List<DesignCurve>()); // nothing to replace → continue as usual

            var answer = MessageBox.Show(
                "Do you want to replace the selected profile(s)?",
                "Replace profile(s)?",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question,
                MessageBoxResult.No);

            if (answer != MessageBoxResult.Yes)
                return (false, new List<DesignCurve>()); // user chose No → exit caller

            var originals = new List<DesignCurve>();

            WriteBlock.ExecuteTask("Replace Construct Profile(s)", () =>
            {
                foreach (var comp in comps)
                    originals.AddRange(UnhideAndGetOriginalCurves(win.Document, comp));

                if (deleteAction != null) deleteAction(comps);
                else DeleteComponents(comps);
            });

            if (originals.Count == 0)
            {
                Application.ReportStatus(
                    "No original lines could be restored for the selected profile(s).",
                    StatusMessageType.Warning, null);
            }

            // Try to show them in selection (visual feedback)
            try
            {
                if (originals.Count > 0 && win?.ActiveContext != null)
                {
                    // Cast DesignCurve -> IDocObject
                    var toSelect = originals
                        .Cast<SpaceClaim.Api.V242.IDocObject>()
                        .ToList(); // List<IDocObject> satisfies ICollection<IDocObject>

                    // Replace the selection in one go (never Clear/Add on the existing selection)
                    win.ActiveContext.Selection = toSelect;
                }
            }
            catch
            {
                // ignore UI/selection issues
            }

            return (true, originals);
        }

        // Finds all selected components whose template has an "AESC_Construct" custom property (Construct-generated profiles).
        public static HashSet<Component> FindConstructComponentsFromSelection(Window win)
        {
            var result = new HashSet<Component>();
            var sel = win?.ActiveContext?.Selection;
            var doc = win?.Document;
            if (doc == null || sel == null || sel.Count == 0) return result;

            foreach (var picked in sel)
            {
                var comp = TryGetOwningComponent(doc, picked);
                if (comp == null) continue;

                var part = comp.Template;
                if (part == null) continue;

                bool isConstruct = part.CustomProperties.Any(kv =>
                    string.Equals(kv.Key, "AESC_Construct", StringComparison.OrdinalIgnoreCase));

                if (isConstruct) result.Add(comp);
            }
            return result;
        }

        // Unhides original main-part DesignCurves that match the stored ConstructCurve endpoints of the given component.
        public static List<DesignCurve> UnhideAndGetOriginalCurves(Document doc, Component comp)
        {
            var restored = new List<DesignCurve>();
            if (doc == null || comp == null || comp.Template == null) return restored;

            var viewContexts = Window.GetWindows(doc)
                                     .Where(w => w != null && w.Document == doc)
                                     .Select(w => w.ActiveContext as IAppearanceContext)
                                     .Where(ctx => ctx != null)
                                     .ToList();

            // Collect world endpoints of ConstructCurve(s) stored inside the component
            var placement = comp.Placement;
            var constructEnds = new List<(Point A, Point B)>();

            foreach (var c in comp.Template.GetChildren<DesignCurve>())
            {
                if (!string.Equals(c.Name, "ConstructCurve", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (TryGetCurveEndpointsLocal(c, out var aLocal, out var bLocal))
                {
                    var aWorld = placement * aLocal;
                    var bWorld = placement * bLocal;
                    constructEnds.Add((aWorld, bWorld));
                }
            }

            if (constructEnds.Count == 0) return restored;

            const double tol = 1e-5;

            // Unhide matching original line(s) that live at MainPart level
            foreach (var curve in doc.MainPart.GetChildren<DesignCurve>())
            {
                if (string.Equals(curve.Name, "ConstructCurve", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!TryGetCurveEndpointsLocal(curve, out var p0, out var p1))
                    continue;

                if (MatchesAnyConstructEnds(p0, p1, constructEnds, tol))
                {
                    curve.SetVisibility(null, true);
                    foreach (var ctx in viewContexts) curve.SetVisibility(ctx, true);
                    restored.Add(curve);
                }
            }
            return restored;
        }

        // Deletes all provided components, swallowing any API exceptions during deletion.
        public static void DeleteComponents(IEnumerable<Component> components)
        {
            foreach (var comp in components.Where(c => c != null).ToList())
            {
                try { comp.Delete(); } catch { /* ignore */ }
            }
        }

        // Converts a set of DesignCurves to simple ITrimmedCurve segments based on their endpoints.
        public static List<ITrimmedCurve> ToSegments(IEnumerable<DesignCurve> curves)
        {
            var list = new List<ITrimmedCurve>();
            foreach (var dc in curves)
            {
                if (dc?.Shape == null) continue;
                if (TryGetCurveEndpointsLocal(dc, out var a, out var b))
                    list.Add(CurveSegment.Create(a, b)); // world-space segment
            }
            return list;
        }

        // Tries to extract curve endpoints (Start/End) from a DesignCurve shape into local coordinates.
        private static bool TryGetCurveEndpointsLocal(DesignCurve dc, out Point start, out Point end)
        {
            start = default(Point);
            end = default(Point);
            var shape = dc.Shape;

            if (shape is CurveSegment seg)
            {
                start = seg.StartPoint; end = seg.EndPoint; return true;
            }
            if (shape is ITrimmedCurve trimmed)
            {
                start = trimmed.StartPoint; end = trimmed.EndPoint; return true;
            }
            return false;
        }

        // Checks if a given segment matches any ConstructCurve endpoints (in either direction) within a tolerance.
        private static bool MatchesAnyConstructEnds(Point p0, Point p1, List<(Point A, Point B)> ends, double tol)
        {
            foreach (var (A, B) in ends)
            {
                if ((p0 - A).Magnitude < tol && (p1 - B).Magnitude < tol) return true;
                if ((p0 - B).Magnitude < tol && (p1 - A).Magnitude < tol) return true;
            }
            return false;
        }

        // Resolves the owning Component for a picked item (component, face, edge, body, curve, etc.) in the active document.
        public static Component TryGetOwningComponent(Document doc, object picked)
        {
            if (picked == null || doc == null) return null;

            DesignCurve dc = null; DesignEdge de = null; DesignFace df = null; DesignBody db = null;

            if (picked is Component c0) return c0;

            if (picked is IDesignCurve idc) dc = idc.Master; else if (picked is DesignCurve dc0) dc = dc0;
            if (picked is IDesignEdge ide) de = ide.Master; else if (picked is DesignEdge de0) de = de0;
            if (picked is IDesignFace idf) df = idf.Master; else if (picked is DesignFace df0) df = df0;
            if (picked is IDesignBody idb) db = idb.Master; else if (picked is DesignBody db0) db = db0;

            foreach (var comp in doc.MainPart.GetChildren<Component>())
            {
                var part = comp.Template;
                if (part == null) continue;

                if (dc != null && part.GetChildren<DesignCurve>().Any(c => ReferenceEquals(c, dc)))
                    return comp;

                if (de != null && part.Bodies.Any(b => b.Shape == de.Shape.Body))
                    return comp;

                if (df != null && part.Bodies.Any(b => b.Shape == df.Shape.Body))
                    return comp;

                if (db != null && part.Bodies.Any(b => ReferenceEquals(b, db)))
                    return comp;
            }
            return null;
        }
    }
}
