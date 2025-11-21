/*
 RotateComponentCommand centralizes all rotation logic for Construct components.

 Active implementation:
 - Rotates each component around the axis defined by its (ConstructCurve) profile in world space.
 - Keeps that axis curve visually fixed by rebuilding it in the new local frame after rotation.
 - Accumulates the total rotation in the Part’s "RotationAngle" custom property.

 Differences with commented variants:
 - Top legacy block (//////...): rotates only the DesignBody.Shape around a local curve axis; it does not change Component.Placement and only touches the body, so the component transform stays unchanged.
 - Middle legacy ApplyRotation (under the XML summary): rotates the component around its local Z axis through the component origin, not around the actual profile/ConstructCurve axis.
 - Bottom legacy block (after the active class): rotates the component around the first curve axis and additionally "bakes" the same rotation into all template DesignCurves; it also has helpers to reapply the stored RotationAngle to components or detached Body instances.
*/

//////using System;
//////using System.Collections.Generic;
//////using System.Linq;
//////using AESCConstruct25.FrameGenerator.Modules;
//////using AESCConstruct25.FrameGenerator.Utilities;
//////using SpaceClaim.Api.V242;
//////using SpaceClaim.Api.V242.Geometry;
//////using SpaceClaim.Api.V242.Modeler;

//////namespace AESCConstruct25.Commands
//////{
//////    public static class RotateComponentCommand
//////    {
//////        public static void Execute(Window window, double rotationAngleDegrees)
//////        {
//////            if (window == null) return;

//////            List<Component> components = JointSelectionHelper.GetSelectedComponents(window);
//////            ApplyRotation(window, components, rotationAngleDegrees);
//////        }

//////        public static void ApplyStoredRotation(Window window, List<Component> components)
//////        {
//////            if (window == null || components == null) return;

//////            foreach (Component comp in components)
//////            {
//////                double storedRotation = 0;
//////                if (comp.Template.CustomProperties.TryGetValue("RotationAngle", out CustomPartProperty prop))
//////                {
//////                    double.TryParse(prop.Value.ToString(), out storedRotation);
//////                }

//////                ApplyRotation(window, new List<Component> { comp }, storedRotation, storeProperty: false);
//////            }
//////        }

//////        private static void ApplyRotation(Window window, List<Component> components, double angleDegrees, bool storeProperty = true)
//////        {
//////            double angleRadians = angleDegrees * Math.PI / 180.0;

//////            foreach (Component comp in components)
//////            {
//////                DesignCurve axisCurve = comp.Template.Curves.FirstOrDefault() as DesignCurve;
//////                DesignBody body = comp.Template.Bodies.FirstOrDefault() as DesignBody;

//////                if (axisCurve == null || body == null) continue;

//////                CurveSegment segment = axisCurve.Shape as CurveSegment;
//////                if (segment == null) continue;

//////                Line axis = Line.Create(segment.StartPoint, (segment.EndPoint - segment.StartPoint).Direction);
//////                Matrix rotation = Matrix.CreateRotation(axis, angleRadians);
//////                body.Shape.Transform(rotation);

//////                if (storeProperty)
//////                {
//////                    double existingRotation = 0.0;

//////                    if (comp.Template.CustomProperties.TryGetValue("RotationAngle", out CustomPartProperty existingProp))
//////                    {
//////                        double.TryParse(existingProp.Value.ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out existingRotation);
//////                    }

//////                    double totalRotation = existingRotation + angleDegrees;
//////                    string angleText = totalRotation.ToString(System.Globalization.CultureInfo.InvariantCulture);

//////                    if (comp.Template.CustomProperties.ContainsKey("RotationAngle"))
//////                    {
//////                        comp.Template.CustomProperties["RotationAngle"].Value = angleText;
//////                    }
//////                    else
//////                    {
//////                        CustomPartProperty.Create(comp.Template, "RotationAngle", angleText);
//////                    }

//////                    //Logger.Log($"Updated RotationAngle = {angleText} (was {existingRotation}) for {comp.Name}");
//////                }

//////            }
//////        }
//////    }
//////}
using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AESCConstruct25.Commands
{
    public static class RotateComponentCommand
    {
        // Legacy Execute overload: rotates all currently selected components by a given angle using the old ApplyRotation implementation above.
        //public static void Execute(Window window, double rotationAngleDegrees)
        //{
        //    if (window == null) return;

        //    List<Component> components = JointSelectionHelper.GetSelectedComponents(window);
        //    ApplyRotation(window, components, rotationAngleDegrees);
        //}

        // Legacy helper: reads "RotationAngle" from each component and reapplies it using the old ApplyRotation overload without re-storing the property.
        //public static void ApplyStoredRotation(Window window, List<Component> components)
        //{
        //    if (window == null || components == null) return;

        //    foreach (Component comp in components)
        //    {
        //        double storedRotation = 0;
        //        if (comp.Template.CustomProperties.TryGetValue("RotationAngle", out CustomPartProperty prop))
        //        {
        //            double.TryParse(prop.Value.ToString(), out storedRotation);
        //        }

        //        ApplyRotation(window, new List<Component> { comp }, storedRotation, storeProperty: false);
        //    }
        //}

        /// <summary>
        /// Rotates each component in world space by angleDegrees around the curve axis in its template,
        /// then counter–rotates that curve so it remains fixed in world space.
        /// Also accumulates a "RotationAngle" custom property on the Part.
        /// </summary>
        // Legacy ApplyRotation: rotates around local Z through the component origin, not around the ConstructCurve axis, and then rebuilds only the first DesignCurve.
        //public static void ApplyRotation(
        //    Window window,
        //    List<Component> components,
        //    double angleDegrees,
        //    bool storeProperty = true
        //)
        //{
        //    if (window == null || components == null || components.Count == 0)
        //        return;

        //    // Convert once to radians
        //    double angleRad = angleDegrees * Math.PI / 180.0;

        //    foreach (var comp in components)
        //    {
        //        Part part = comp.Template;
        //        var dc = part.Curves.OfType<DesignCurve>().FirstOrDefault();
        //        var body = part.Bodies.OfType<DesignBody>().FirstOrDefault();
        //        if (dc == null || body == null)
        //            continue;

        //        // — 1) Build world-space rotation axis along local Z through component origin —
        //        Point originLocal = Point.Origin;
        //        Vector zLocal = Vector.Create(0, 0, 1);
        //        Point originWorld = comp.Placement * originLocal;
        //        Direction axisDir = (comp.Placement * zLocal).Direction;
        //        Line axis = Line.Create(originWorld, axisDir);

        //        // — 2) Snapshot the curve’s world endpoints BEFORE rotation —
        //        var seg = dc.Shape as CurveSegment;
        //        if (seg == null)
        //            continue;
        //        Point worldStart = comp.Placement * seg.StartPoint;
        //        Point worldEnd = comp.Placement * seg.EndPoint;

        //        // — 3) Rotate the component placement around axis —
        //        var rotation = Matrix.CreateRotation(axis, angleRad);
        //        comp.Placement = rotation * comp.Placement;

        //        // — 4) Reproject those saved world points back into the NEW local frame —
        //        var invPlacement = comp.Placement.Inverse;
        //        Point newLocalStart = invPlacement * worldStart;
        //        Point newLocalEnd = invPlacement * worldEnd;

        //        // — 5) Delete old curve and recreate it exactly at the same world location —
        //        string oldName = dc.Name;
        //        var oldColor = dc.GetColor(null);
        //        dc.Delete();

        //        var newSeg = CurveSegment.Create(newLocalStart, newLocalEnd);
        //        var newDc = DesignCurve.Create(part, newSeg);
        //        newDc.Name = oldName;
        //        newDc.SetColor(null, oldColor);

        //        // — 6) Optionally store cumulative RotationAngle in the Part’s properties —
        //        if (storeProperty)
        //        {
        //            double prev = 0;
        //            if (part.CustomProperties.TryGetValue("RotationAngle", out var prop) &&
        //                double.TryParse(prop.Value.ToString(),
        //                                NumberStyles.Float,
        //                                CultureInfo.InvariantCulture,
        //                                out var parsed))
        //            {
        //                prev = parsed;
        //            }

        //            double total = prev + angleDegrees;
        //            string text = total.ToString(CultureInfo.InvariantCulture);

        //            if (part.CustomProperties.ContainsKey("RotationAngle"))
        //                part.CustomProperties["RotationAngle"].Value = text;
        //            else
        //                CustomPartProperty.Create(part, "RotationAngle", text);
        //        }
        //    }
        //}

        // Rotates each component around its (ConstructCurve) axis in world space, keeps that axis curve visually fixed, and updates the "RotationAngle" part property.
        public static void ApplyRotation(
            Window window,
            List<Component> components,
            double angleDegrees,
            bool storeProperty = true
        )
        {
            if (window == null || components == null || components.Count == 0)
                return;

            double angleRad = angleDegrees * Math.PI / 180.0;

            foreach (var comp in components)
            {
                var part = comp.Template;

                // pick the axis curve inside the component (prefer a named one if present)
                var dc = part.Curves
                             .OfType<DesignCurve>()
                             .FirstOrDefault(c => string.Equals(c.Name, "ConstructCurve", StringComparison.OrdinalIgnoreCase))
                      ?? part.Curves.OfType<DesignCurve>().FirstOrDefault();

                var body = part.Bodies.OfType<DesignBody>().FirstOrDefault();
                if (dc == null || body == null)
                    continue;

                if (dc.Shape is not CurveSegment seg)
                    continue;

                // world-space endpoints of the axis curve BEFORE rotation
                var place = comp.Placement;
                Point w0 = place * seg.StartPoint;
                Point w1 = place * seg.EndPoint;

                // guard: degenerate axis → skip
                var dirVec = (w1 - w0);
                if (dirVec.Magnitude < 1e-9)
                    continue;

                var axisDir = dirVec.Direction;
                var midWorld = Point.Create(
                    0.5 * (w0.X + w1.X),
                    0.5 * (w0.Y + w1.Y),
                    0.5 * (w0.Z + w1.Z)
                );
                var axisWorld = Line.Create(midWorld, axisDir);

                // rotate the COMPONENT placement around the axis defined by its own curve
                var rotation = Matrix.CreateRotation(axisWorld, angleRad);
                comp.Placement = rotation * comp.Placement;

                // keep the axis curve fixed in world space: rebuild it in the new local frame
                var invNew = comp.Placement.Inverse;
                Point newLocalStart = invNew * w0;
                Point newLocalEnd = invNew * w1;

                string oldName = dc.Name;
                var oldColor = dc.GetColor(null);
                bool oldVis = true;
                try { oldVis = dc.GetVisibility(null) == true; } catch { }
                Logger.Log($"RotateComponent: oldVis={oldVis}");
                var oldLayer = dc.Layer;

                dc.Delete();

                var newSeg = CurveSegment.Create(newLocalStart, newLocalEnd);
                var newDc = DesignCurve.Create(part, newSeg);
                newDc.Name = oldName;
                newDc.SetColor(null, oldColor);
                var ctx = window?.ActiveContext as IAppearanceContext;
                try { newDc.SetVisibility(ctx, oldVis); } catch { }
                try { newDc.Layer = oldLayer; } catch { }

                // accumulate RotationAngle on the Part
                if (storeProperty)
                {
                    double prev = 0;
                    if (part.CustomProperties.TryGetValue("RotationAngle", out var prop)
                        && double.TryParse(prop.Value.ToString(),
                                           NumberStyles.Float,
                                           CultureInfo.InvariantCulture,
                                           out var parsed))
                    {
                        prev = parsed;
                    }

                    double total = prev + angleDegrees;
                    string text = total.ToString(CultureInfo.InvariantCulture);

                    if (part.CustomProperties.ContainsKey("RotationAngle"))
                        part.CustomProperties["RotationAngle"].Value = text;
                    else
                        CustomPartProperty.Create(part, "RotationAngle", text);
                }
            }
        }

    }
}

//// ^^^ rotation around local axis instead of component axis

//using AESCConstruct25.FrameGenerator.Utilities;
//using SpaceClaim.Api.V242;
//using SpaceClaim.Api.V242.Geometry;
//using SpaceClaim.Api.V242.Modeler;
//using System;
//using System.Collections.Generic;
//using System.Globalization;
//using System.Linq;

//namespace AESCConstruct25.FrameGenerator.Commands
//{
//    public static class RotateComponentCommand
//    {
//        public static void Execute(Window window, double rotationAngleDegrees)
//        {
//            if (window == null) return;

//            List<Component> components = JointSelectionHelper.GetSelectedComponents(window);
//            ApplyRotation(window, components, rotationAngleDegrees);
//        }

//        private static void ApplyRotation(
//            Window window,
//            List<Component> components,
//            double angleDegrees,
//            bool storeProperty = true
//        )
//        {
//            if (window == null || components == null) return;
//            double angleRad = angleDegrees * Math.PI / 180.0;

//            foreach (var comp in components)
//            {
//                // 1) find the axis‐curve in local space
//                var dc = comp.Template.Curves.OfType<DesignCurve>().FirstOrDefault();
//                if (dc?.Shape is not CurveSegment seg)
//                    continue;

//                // 2) build the world‐axis
//                var midLocal = Point.Create(
//                    0.5 * (seg.StartPoint.X + seg.EndPoint.X),
//                    0.5 * (seg.StartPoint.Y + seg.EndPoint.Y),
//                    0.5 * (seg.StartPoint.Z + seg.EndPoint.Z)
//                );
//                var place = comp.Placement;
//                var midWorld = place * midLocal;
//                var dirWorld = (place * seg.EndPoint - place * seg.StartPoint).Direction;
//                var axisWorld = Line.Create(midWorld, dirWorld);

//                // 3) rotate the component placement
//                var rot = Matrix.CreateRotation(axisWorld, angleRad);
//                comp.Placement = rot * comp.Placement;

//                // 4) now “bake” that same rotation into each DesignCurve in the template
//                var invNew = comp.Placement.Inverse;
//                var part = comp.Template;
//                // take a snapshot since we'll be deleting from the collection
//                foreach (var curve in part.Curves.OfType<DesignCurve>().ToList())
//                {
//                    if (curve.Shape is not CurveSegment oldSeg)
//                        continue;

//                    // compute the new local endpoints
//                    // a) local → world (before rotation)
//                    var w0 = place * oldSeg.StartPoint;
//                    var w1 = place * oldSeg.EndPoint;
//                    // b) apply rotation in world
//                    var w0r = rot * w0;
//                    var w1r = rot * w1;
//                    // c) back into local (after rotation)
//                    var nl0 = invNew * w0r;
//                    var nl1 = invNew * w1r;

//                    // remember its name, visibility and color (optional)
//                    var name = curve.Name;
//                    //bool vis = curve.GetVisibility(null);
//                    var color = curve.GetColor(null);

//                    // delete the old curve…
//                    curve.Delete();
//                    // …and recreate it with the rotated segment
//                    var newSeg = CurveSegment.Create(nl0, nl1);
//                    var newDc = DesignCurve.Create(part, newSeg);
//                    newDc.Name = name;
//                    //newDc.SetVisibility(null, vis);
//                    newDc.SetColor(null, color);
//                }

//                // 5) store cumulative angle as before
//                if (storeProperty)
//                {
//                    double prev = 0;
//                    if (comp.Template.CustomProperties
//                             .TryGetValue("RotationAngle", out var p))
//                    {
//                        double.TryParse(p.Value.ToString(),
//                                        NumberStyles.Float,
//                                        CultureInfo.InvariantCulture,
//                                        out prev);
//                    }

//                    double total = prev + angleDegrees;
//                    string text = total.ToString(CultureInfo.InvariantCulture);

//                    if (comp.Template.CustomProperties.ContainsKey("RotationAngle"))
//                        comp.Template.CustomProperties["RotationAngle"].Value = text;
//                    else
//                        CustomPartProperty.Create(comp.Template, "RotationAngle", text);

//                    //Logger.Log($"Updated RotationAngle = {text} (was {prev}) for {comp.Name}");
//                }
//            }
//        }


//        public static void ApplyStoredRotation(
//            Window window,
//            List<Component> components
//        )
//        {
//            if (window == null || components == null) return;
//            foreach (var c in components)
//            {
//                if (!c.Template.CustomProperties.TryGetValue("RotationAngle", out var p)
//                 || !double.TryParse(p.Value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var a))
//                    continue;
//                // apply once, and do not re-store the property
//                ApplyRotation(window, new List<Component> { c }, a, storeProperty: false);
//            }
//        }

//        /// <summary>
//        /// If you do still need to rotate a *detached* Body (e.g. a cutter or a half you
//        /// copied out), this will re-apply exactly the component’s stored RotationAngle
//        /// to that body, so that it matches the component’s world orientation.
//        /// But *normal* joints no longer need this, because the component itself is rotated.
//        /// </summary>
//        public static Body ApplyStoredRotationOnBody(Component component, Body body)
//        {
//            if (component == null || body == null)
//                return body;

//            if (!component.Template.CustomProperties.TryGetValue("RotationAngle", out var p)
//             || !double.TryParse(p.Value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var angleDeg)
//             || Math.Abs(angleDeg) < 1e-6)
//                return body;

//            // same axis calc as above:
//            var dc = component.Template.Curves.OfType<DesignCurve>().FirstOrDefault();
//            if (dc?.Shape is not CurveSegment seg)
//                return body;

//            var midLocal = Point.Create(
//                0.5 * (seg.StartPoint.X + seg.EndPoint.X),
//                0.5 * (seg.StartPoint.Y + seg.EndPoint.Y),
//                0.5 * (seg.StartPoint.Z + seg.EndPoint.Z)
//            );
//            var place = component.Placement;
//            var midWorld = place * midLocal;
//            var dirWorld = (place * seg.EndPoint - place * seg.StartPoint).Direction;
//            var axisWorld = Line.Create(midWorld, dirWorld);

//            double angleRad = angleDeg * Math.PI / 180.0;
//            var rot = Matrix.CreateRotation(axisWorld, angleRad);
//            body.Transform(rot);
//            return body;
//        }
//    }
//}
