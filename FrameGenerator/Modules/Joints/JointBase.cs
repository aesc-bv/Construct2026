using SpaceClaim.Api.V242;
using SpaceClaim.Api.V242.Modeler;

namespace AESCConstruct2026.FrameGenerator.Modules.Joints
{
    /// <summary>Abstract base class for all joint types (Miter, Straight, T, Trim).</summary>
    public abstract class JointBase
    {
        public abstract string Name { get; }

        public abstract void Execute(Component componentA, Component componentB, double spacing, Body bodyA, Body bodyB);
    }
}
