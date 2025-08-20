using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace AESCConstruct25.FrameGenerator.Utilities
{
    public static class ProfileSelectionHelper
    {
        /// <summary>
        /// Retrieves all selected curves (either DesignCurves or DesignEdges) as ITrimmedCurve objects in world space.
        /// </summary>
        public static List<ITrimmedCurve> GetSelectedCurves(Window window)
        {
            var list = new List<ITrimmedCurve>();
            var sel = window?.ActiveContext?.Selection;
            if (sel == null || sel.Count == 0) return list;

            var mainPart = window.Document?.MainPart;

            foreach (var o in sel)
            {
                ITrimmedCurve cur = null;
                Part owningPart = null;

                // Direct curve
                if (o is DesignCurve dc)
                {
                    cur = dc.Shape;
                    owningPart = dc.Parent as Part;
                }
                // Edge of a body
                else if (o is DesignEdge de)
                {
                    cur = de.Shape;
                    var parentBody = de.Parent as DesignBody;
                    if (parentBody != null)
                        owningPart = parentBody.Parent as Part;
                }
                // Instance referencing a master edge
                else if (o is IDocObject dobj && dobj.Master is DesignEdge mde)
                {
                    cur = mde.Shape;
                    var parentBody = mde.Parent as DesignBody;
                    if (parentBody != null)
                        owningPart = parentBody.Parent as Part;
                }
                // Instance referencing a master curve
                else if (o is IDocObject dobj2 && dobj2.Master is DesignCurve mdc)
                {
                    cur = mdc.Shape;
                    owningPart = mdc.Parent as Part;
                }

                if (cur == null) continue;
                if (!(cur.Geometry is Line)) continue;

                // Locate the displayed component instance for this Part (fallback: identity)
                Matrix toWorld = Matrix.Identity;
                if (owningPart != null && mainPart != null)
                {
                    var comp =
                        mainPart.GetDescendants<IComponent>()
                                .OfType<Component>()
                                .FirstOrDefault(c => c.Template == owningPart);
                    if (comp != null)
                        toWorld = comp.Placement;
                }

                // Build a world-space segment
                Point w0 = toWorld * cur.StartPoint;
                Point w1 = toWorld * cur.EndPoint;
                list.Add(CurveSegment.Create(w0, w1));
            }

            return list;
        }
    }
}
