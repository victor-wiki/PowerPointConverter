namespace PowerPointConverter.Model
{
    public class TextStyle
    {
        public string Color { get; set; }       
        public double? FontSize { get; set; }
        public string FontFamily { get; set; }
        public List<string> FontFamilyStack { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }   
        public bool IsStrike { get; set; }
        public string HighlightColor { get; set; }
        public string UnderlineColor { get; set; }
        public bool UnderlineFollowsText { get; set; }
        public string Alignment { get; set; }
        public bool RightToLeft { get; set; }
        public double? MarginLeft { get; set; }
        public double? MarginRight { get; set; }
        public double? LetterSpacingPoints { get; set; }
        public double? Indent { get; set; }
        public string LineHeight { get; set; }
        public bool IsAbsoluteLineHeight { get; set; }
        public double? SpaceBeforePercent { get; set; }
        public double? SpaceBeforePoints { get; set; }
        public double? SpaceAfterPercent { get; set; }
        public double? SpaceAfterPoints { get; set; }
        public bool IsTextNoFill { get; set; }
        public string GradientFillCss { get; set; }
        public string PatternFillCss { get; set; }
        public double? Kern { get; set; }
        public string Capital { get; set; }
        public int? Baseline { get; set; }
        public string TextShadow { get; set; }
        public double? OutlineWidth { get; set; }
        public string OutlineColor { get; set; }
        public string OutlineGradientCss { get; set; }
        public bool? BulletNone { get; set; }
        public string? BulletChar { get; set; }
        public string? BulletFontName { get; set; }
        public string? BulletAutoNumber { get; set; }
        public int? BulletAutoNumberStartAt { get; set; }
        public double? BulletSizePercent { get; set; }
        public double? BulletSizePoints { get; set; }
        public string? BulletColor { get; set; }
        public bool? BulletColorFollowsText { get; set; }
    }
}
