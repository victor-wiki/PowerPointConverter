namespace PowerPointConverter.Model
{
    public class PresetGeometryInfo
    {
        public string PathD { get; set; }
        public List<ArrowPathInfo> MultiPaths { get; set; }
        public string EffectivePreset { get; set; }
        public string AdjustmentKey { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }
}
