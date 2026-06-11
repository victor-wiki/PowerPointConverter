namespace PowerPointConverter.Model
{
    public class ConvertResult
    {
        public bool IsOK => this.Infos?.All(item => item.IsOK) == true;

        public List<HtmlConvertInfo> Infos { get; set; }
        public string Message
        {
            get
            {
                return string.Join(Environment.NewLine, this.Infos?.Select(item=> $"Slide{(item.Index)}:{item.Message}" ));
            }
        }
    }
}
