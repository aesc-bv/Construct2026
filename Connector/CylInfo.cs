using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

public static class CylInfo
{
    /// <summary>
    /// Find the coaxial mate of a cylindrical face on the same body and return
    /// inner/outer faces, thickness, outer radius, and the (outer) axis.
    /// </summary>
    /// <param name="iBody">The body that owns the face (occurrence or master).</param>
    /// <param name="desFace">A cylindrical DesignFace on that body.</param>
    /// <param name="linearTol">Distance tolerance for axis coincidence test.</param>
    /// <param name="angTol">Angle tolerance (radians) for axis parallelism.</param>
    public static (DesignFace InnerFace,
                   DesignFace OuterFace,
                   double Thickness,
                   double OuterRadius,
                   Line Axis)
        GetCoaxialCylPair(IDesignBody iBody, DesignFace desFace)
    {
        if (iBody == null) throw new ArgumentNullException(nameof(iBody));
        if (desFace == null) throw new ArgumentNullException(nameof(desFace));

        // Work with masters for robust topology/geometry access
        var bodyM = iBody as DesignBody ?? (iBody.Master as DesignBody) ?? iBody.Master as DesignBody;
        if (bodyM == null) throw new ArgumentException("Cannot resolve DesignBody master from IDesignBody.", nameof(iBody));

        var faceM = desFace; // in case we were already passed a master
        var seedCyl = faceM.Shape.GetGeometry<Cylinder>();
        if (seedCyl == null) throw new ArgumentException("The provided DesignFace is not cylindrical.", nameof(desFace));

        var seedAxis = seedCyl.Axis;
        var seedR = seedCyl.Radius;

        // Collect all OTHER cylindrical faces on the same body that are coaxial
        var coaxial = new List<(DesignFace F, Cylinder C, double R)>();
        foreach (var f in bodyM.Faces)
        {
            if (ReferenceEquals(f, faceM)) continue;
            var cyl = f.Shape.GetGeometry<Cylinder>();
            if (cyl == null) continue;
            if (AreCoaxial(seedAxis, cyl.Axis))
            {
                coaxial.Add((f, cyl, cyl.Radius));
            }
        }

        if (coaxial.Count == 0)
        {
            // No mate found — return the given face as both (thickness 0)
            var seedOuterAxis = seedAxis;
            return (InnerFace: null,
                    OuterFace: faceM,
                    Thickness: 0.0,
                    OuterRadius: seedR,
                    Axis: seedOuterAxis);
        }

        // Choose the mate with the closest radius difference (handles multiple coaxial cylinders)
        var mate = coaxial.OrderBy(t => Math.Abs(t.R - seedR)).First();

        // Decide inner/outer purely by radius
        DesignFace inner, outer;
        Cylinder innerCyl, outerCyl;
        if (seedR <= mate.R)
        {
            inner = faceM; innerCyl = seedCyl;
            outer = mate.F; outerCyl = mate.C;
        }
        else
        {
            inner = mate.F; innerCyl = mate.C;
            outer = faceM; outerCyl = seedCyl;
        }

        double thickness = Math.Abs(outerCyl.Radius - innerCyl.Radius);
        // As requested, report axis & radius from the OUTER face
        return (InnerFace: inner,
                OuterFace: outer,
                Thickness: thickness,
                OuterRadius: outerCyl.Radius,
                Axis: outerCyl.Axis);
    }

    /// <summary>
    /// Returns true if two (infinite) lines are coaxial (colinear) within tolerances.
    /// Requires axes to be parallel and the minimum separation between lines to be small.
    /// </summary>
    private static bool AreCoaxial(Line a, Line b)
    {
        // 1) parallel check via dot product of directions

        bool isCoincident = a.IsCoincident(b);

        Direction cross = Direction.Cross(a.Direction, b.Direction);

        double cos = Math.Abs(Vector.Dot(a.Direction.UnitVector, b.Direction.UnitVector));

        return isCoincident;
    }


    // Get inner and outer edges based on inner and outer faces and points
    public static (bool Success, DesignEdge InnerEdge, DesignEdge OuterEdge)
        GetEdges(DesignFace fInner, Point pInner, DesignFace fOuter, Point pOuter)
    {
        if (fInner == null || fOuter == null) return (false, null, null);

        // Find closest edge on the inner face
        DesignEdge inner = null;
        double bestInner = double.PositiveInfinity;
        foreach (var e in fInner.Edges)
        {
            var proj = e.Shape.ProjectPoint(pInner);
            if (proj == null) continue; // just in case
            double d = (proj.Point - pInner).Magnitude;
            if (d < bestInner)
            {
                bestInner = d;
                inner = e;
                if (bestInner == 0) break;
            }
        }

        // Find closest edge on the outer face
        DesignEdge outer = null;
        double bestOuter = double.PositiveInfinity;
        foreach (var e in fOuter.Edges)
        {
            var proj = e.Shape.ProjectPoint(pOuter);
            if (proj == null) continue;
            double d = (proj.Point - pOuter).Magnitude;
            if (d < bestOuter)
            {
                bestOuter = d;
                outer = e;
                if (bestOuter == 0) break;
            }
        }

        bool success = inner != null && outer != null;
        return (success, inner, outer);
    }

    public static (bool Success, Point PointEdge) GetClosestPoint(DesignEdge designEdge, Point point, Direction direction)
    {
        Point returnPoint = Point.Origin;
        bool success = false;

        var separation = designEdge.Shape.GetClosestSeparation(CurveSegment.Create(point + 100 * direction, point - 100 * direction));
        double dist = separation.Distance;

        if (dist < 1e-6)
        {
            success = true;
            returnPoint = separation.PointB;
        }

        return (success, returnPoint);
    }

}
