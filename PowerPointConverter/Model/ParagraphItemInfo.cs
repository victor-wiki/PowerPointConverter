namespace PowerPointConverter.Model
{
    public class ParagraphItemInfo
    {
        public double MarginLeft { get; set; }
        public double MarginRight { get; set; }
        public double Indent { get; set; }
        public bool IsBullet { get; set; }        
        public bool LastItemIsLineBreak { get; set; }
    }
}
