/*
 RibCutOutSelectionHelper provides helper logic to derive overlapping body pairs.
 It inspects a list of SpaceClaim Body instances and returns all pairs whose bounding boxes intersect.
*/

using SpaceClaim.Api.V242.Geometry;
using SpaceClaim.Api.V242.Modeler;
using System.Collections.Generic;

namespace AESCConstruct2026.FrameGenerator.Utilities
{
    public static class RibCutOutSelectionHelper
    {
        // Returns all unique body pairs from the input list whose axis-aligned bounding boxes overlap in 3D.
        public static List<(Body A, Body B)> GetOverlappingPairs(List<Body> bodies)
        {
            var pairs = new List<(Body A, Body B)>();
            for (int i = 0; i < bodies.Count; i++)
            {
                for (int j = i + 1; j < bodies.Count; j++)
                {
                    var a = bodies[i];
                    var b = bodies[j];
                    var boxA = a.GetBoundingBox(Matrix.Identity, true);
                    var boxB = b.GetBoundingBox(Matrix.Identity, true);

                    if (boxA.MinCorner.X <= boxB.MaxCorner.X && boxA.MaxCorner.X >= boxB.MinCorner.X &&
                        boxA.MinCorner.Y <= boxB.MaxCorner.Y && boxA.MaxCorner.Y >= boxB.MinCorner.Y &&
                        boxA.MinCorner.Z <= boxB.MaxCorner.Z && boxA.MaxCorner.Z >= boxB.MinCorner.Z)
                    {
                        pairs.Add((a, b));
                    }
                }
            }
            return pairs;
        }
    }
}
