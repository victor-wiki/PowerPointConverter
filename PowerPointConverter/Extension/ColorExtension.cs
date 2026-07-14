using D = System.Drawing;

namespace PowerPointConverter.Extension
{
    public static class ColorExtension
    {
        public static string ToHex(this D.Color color)
        {
            string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";            

            return hex;
        }
    }
}
