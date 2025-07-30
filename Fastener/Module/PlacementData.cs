using SpaceClaim.Api.V242.Geometry;

namespace AESCConstruct25.Fastener.Module
{
    public class PlacementData
    {
        public Circle Circle { get; set; }
        public Point Origin { get; set; }
        public Direction Direction { get; set; }
        public double Depth { get; set; }
    }
}
