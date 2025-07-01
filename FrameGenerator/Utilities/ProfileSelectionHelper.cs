using System;
using System.Collections.Generic;
using AESCConstruct25.FrameGenerator.Modules.Profiles;
using AESCConstruct25.FrameGenerator.Utilities;
using SpaceClaim.Api.V251;
using SpaceClaim.Api.V251.Geometry;
using SpaceClaim.Api.V251.Modeler;

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class ProfileSelectionHelper
    {
        /// <summary>
        /// Retrieves all selected curves (either DesignCurves or DesignEdges) as ITrimmedCurve objects.
        /// </summary>
        public static List<ITrimmedCurve> GetSelectedCurves(Window window)
        {
            List<ITrimmedCurve> selectedCurves = new List<ITrimmedCurve>();
            var selectedObjects = window.ActiveContext.Selection;

            foreach (var obj in selectedObjects)
            {
                ITrimmedCurve curve = null;

                if (obj is DesignCurve designCurve)
                {
                    curve = designCurve.Shape;
                    //Logger.Log("Selected DesignCurve.");
                }
                else if (obj is DesignEdge edge)
                {
                    curve = edge.Shape;
                    //Logger.Log("Selected DesignEdge.");
                }
                else if (obj.Master is DesignEdge componentEdge)
                {
                    curve = componentEdge.Shape;
                    //Logger.Log("Selected ComponentEdge via Master.");
                }

                if (curve != null)
                {
                    selectedCurves.Add(curve);
                }
                else
                {
                    //Logger.Log("Skipped object - no valid curve shape.");
                }
            }

            //Logger.Log($"Total selected curves: {selectedCurves.Count}");
            return selectedCurves;
        }
    }
}
