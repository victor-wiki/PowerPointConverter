using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPointConverter.Model
{
    public class TransformInfo
    {
        public A.Offset? Offset { get; set; }
        public A.Extents Extents { get; set; }
        public A.ChildOffset? ChildOffset { get; set; }
        public A.ChildExtents? ChildExtents { get; set; }
        public Int32Value? Rotation { get; set; }
        
    }
}
