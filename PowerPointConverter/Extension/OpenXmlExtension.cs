using DocumentFormat.OpenXml;
using PowerPointConverter.Helper;
using PowerPointConverter.Model;
using A = DocumentFormat.OpenXml.Drawing;
using O = DocumentFormat.OpenXml.Office.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPointConverter.Extension
{
    public static class OpenXmlExtension
    {
        public static int? GetId(this OpenXmlElement element)
        {
            string id = null;

            if (element is P.Shape s)
            {
                id = s?.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id;
            }
            if (element is O.Shape os)
            {
                id = os?.ShapeNonVisualProperties?.NonVisualDrawingProperties?.Id;
            }
            if (element is P.Picture p)
            {
                id = p?.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id;
            }            

            if (id != null)
            {
                return int.Parse(id);
            }

            return null;
        }

        public static TextBodyInfo GetTextBody(this OpenXmlElement element)
        {
            if(element == null)
            {
                return null;
            }

            OpenXmlElement textBody = (element is P.Shape) ? (element as P.Shape).TextBody : (element as O.Shape).TextBody;

            return new TextBodyInfo()
            {
                BodyProperties = textBody.GetFirstChild<A.BodyProperties>(),
                Paragraphs = textBody.Elements<A.Paragraph>(),
                ListStyle = textBody.GetFirstChild<A.ListStyle>()
            };
        }

        public static TransformInfo GetTransform(this OpenXmlElement element)
        {
            if(element == null)
            {
                return null;
            }

            if(element is P.GroupShape gp)
            {
                var transform = gp.GroupShapeProperties?.TransformGroup;

                return new TransformInfo() { Offset = transform?.Offset, Extents = transform?.Extents, Rotation = transform?.Rotation, ChildOffset = transform?.ChildOffset, ChildExtents=transform?.ChildExtents };
            }
            else if (element is P.GraphicFrame gf)
            {
                var transform = gf.Transform;

                return new TransformInfo() { Offset = transform?.Offset, Extents = transform?.Extents, Rotation = transform?.Rotation };
            }     
            else
            {
                string shapePropertiesName = nameof(P.ShapeProperties);

                if (ObjectHelper.HasProperty(element, shapePropertiesName))
                {
                    var shapeProperties = ObjectHelper.GetValue(element, shapePropertiesName);

                    if(ObjectHelper.HasProperty(shapeProperties, nameof(A.Transform2D)))
                    {
                        var transform = ObjectHelper.GetValue(shapeProperties, nameof(A.Transform2D)) as A.Transform2D;

                        return new TransformInfo() { Offset = transform?.Offset, Extents = transform?.Extents, Rotation = transform?.Rotation };
                    }
                }                    
            }

            return null;
        }
    }
}
