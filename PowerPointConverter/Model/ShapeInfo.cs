using DocumentFormat.OpenXml;
using ShapeCrawler;

namespace PowerPointConverter.Model
{
    public class ShapeInfo
    {
        public OpenXmlElement OpenXmlElement { get; set; }
        public PlaceholderType? PlaceholderType { get; set; }        
        public Geometry? GeometryType { get; set; }
        public ShapePropertiesInfo ShapeProperties { get; set; }
        public TextBodyInfo TextBody { get; set; }
    }
}
