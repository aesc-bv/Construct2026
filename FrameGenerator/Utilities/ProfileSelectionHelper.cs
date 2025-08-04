using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System.Collections.Generic;

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class ProfileSelectionHelper
    {
        /// <summary>
        /// Retrieves all selected curves (either DesignCurves or DesignEdges) as ITrimmedCurve objects.
        /// </summary>
        public static List<ITrimmedCurve> GetSelectedCurves(Window window)
        {
            List<ITrimmedCurve> selectedCurves = new();
            var selectedObjects = window.ActiveContext.Selection;

            foreach (var obj in selectedObjects)
            {
                ITrimmedCurve curve = null;

                if (obj is not NurbsCurve)
                {
                    if (obj is DesignCurve genericCurve)
                    {
                        curve = genericCurve.Shape;
                    }
                    else if (obj is DesignEdge edge)
                    {
                        curve = edge.Shape;
                    }
                    else if (obj.Master is DesignEdge componentEdge)
                    {
                        curve = componentEdge.Shape;
                    }

                    if (curve != null && curve.Geometry is Line)
                    {
                        Logger.Log("→ Valid straight line curve added.");
                        selectedCurves.Add(curve);
                    }
                    else
                    {
                        Logger.Log($"[Selection] obj.GetType() = {obj?.GetType().FullName}");
                        Logger.Log("→ Skipped: null curve or non-line.");
                    }
                }
            }

            return selectedCurves;
        }

    }
}
