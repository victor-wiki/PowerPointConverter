using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPointConverter.Model
{
    public class TextBodyInfo
    {
        public A.BodyProperties? BodyProperties { get; set; }
        public IEnumerable<A.Paragraph> Paragraphs { get; set; }
        public A.ListStyle ListStyle { get; set; }
    }
}
