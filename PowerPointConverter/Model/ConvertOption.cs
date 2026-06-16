namespace PowerPointConverter.Model
{
    public class ConvertOption
    {
        public List<int> SlideNumbers { get; set; }
        public bool ReduceImageQuality { get; set; }
        public bool EnableLog { get; set; }
        public string DefaultLogFolder { get; set; }
    }
}
