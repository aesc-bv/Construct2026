using SpaceClaim.Api.V242.Geometry;

namespace AESCConstruct2026.Fastener.Module
{
    /// <summary>Holds the placement context (circle, origin, direction, depth) for a single fastener insertion point.</summary>
    public class PlacementData
    {
        public Circle Circle { get; set; }
        public Point Origin { get; set; }
        public Direction Direction { get; set; }
        public double Depth { get; set; }
    }
}
