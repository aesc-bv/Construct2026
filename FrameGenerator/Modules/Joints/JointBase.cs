using SpaceClaim.Api.V251;
using SpaceClaim.Api.V251.Modeler;

namespace AESCConstruct25.FrameGenerator.Modules.Joints
{
    public abstract class JointBase
    {
        public abstract string Name { get; }

        public abstract void Execute(Component componentA, Component componentB, double spacing, Body bodyA, Body bodyB);
    }
}
