using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointConverter.Helper
{
    public class OpenXmlHelper
    {
        public static readonly string[] UniquePlaceholderTypes = ["title", "ctrTitle", "subTitle", "dt", "ftr", "sldNum"];
        public static readonly string[] DefaultButNotApplyPlaceholderTypes = ["title", "ctrTitle", "subTitle", "dt", "ftr", "sldNum"];

        public static string GetAttributeValue(OpenXmlElement element, string name)
        {
            if (element == null || name == null)
            {
                return null;
            }

            var attribute = element.GetAttributes()?.FirstOrDefault(item => item.LocalName == name);

            if (attribute != null && attribute.Value != null)
            {
                return attribute.Value.Value;
            }

            return null;
        }

        public static PlaceholderShape GetPlaceholderTypeNode(OpenXmlElement element)
        {
            return element?.Descendants<P.PlaceholderShape>()?.FirstOrDefault();
        }

        public static string GetPlaceholderType(OpenXmlElement element)
        {
            var ph = GetPlaceholderTypeNode(element);

            return ph?.Type;
        }

        public static bool IsBody(OpenXmlElement element)
        {
            var node = GetPlaceholderTypeNode(element);

            if(node == null)
            {
                return false;
            }

            var type = node.Type;

            return (type == null || type == "body");            
        }

        public static Geometry? GetGeometryType(P.ShapeProperties shapeProperties)
        {
            var preset = shapeProperties.GetFirstChild<A.PresetGeometry>()?.Preset;

            if (preset is null)
            {
                if (shapeProperties.OfType<CustomGeometry>().Any())
                {
                    return Geometry.Custom;
                }

                return Geometry.Rectangle;
            }

            Geometry geom;

            if (!ShapeGeometry.ShapeTypeValuesToGeometryMap.TryGetValue(preset, out geom))
            {
                var presetString = preset.ToString()!;
                var name = presetString.ToLowerInvariant().Replace("rect", "rectangle").Replace("diag", "diagonal");
                return (Geometry)Enum.Parse(typeof(Geometry), name, true);
            }

            return geom;
        }      
    }
}
