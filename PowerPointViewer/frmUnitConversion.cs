using PowerPointConverter.Helper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace PowerPointViewer
{
    public partial class frmUnitConversion : Form
    {
        public frmUnitConversion()
        {
            InitializeComponent();
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            string value = this.GetCleanValue(this.txtValue.Text);

            if (!this.IsIntegerValue(value))
            {
                MessageBox.Show("Value must be an integer!");
                return;
            }

            long integerValue = Convert.ToInt64(value);

            var emus = new ShapeCrawler.Units.Emus(integerValue);
            decimal points = Math.Round(emus.AsPoints(), 2);
            decimal pixels = Math.Round(emus.AsPixels(), 2);

            this.txtPoints.Text = points.ToString();
            this.txtPixels.Text = pixels.ToString();

            if (this.txtPixels.Text.Length > 0)
            {
                this.ConvertPixelsToCentimeter((double)pixels);
            }
        }

        private float GetDpi()
        {
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = graphics.DpiX;
                float dpiY = graphics.DpiY;

                return dpiX;
            }
        }

        private string GetCleanValue(string value)
        {
            return value.Replace(",", "").Trim();
        }

        private bool IsIntegerValue(string value)
        {
            return long.TryParse(value, out _);
        }

        private bool IsNumber(string value)
        {
            return decimal.TryParse(value, out _);
        }

        private void txtPixels_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string value = this.GetCleanValue(this.txtPixels.Text);

            if (value.Length > 0 && this.IsNumber(value))
            {
                this.ConvertPixelsToCentimeter(Convert.ToDouble(value));
            }
        }

        private void ConvertPixelsToCentimeter(double pixels)
        {
            float dpi = this.GetDpi();

            double centimeter = Math.Round((float)pixels / dpi * 2.54, 2);

            this.txtCentimeter.Text = centimeter.ToString();
        }

        private void ConvertCentimeterToPixels(double centimeter)
        {
            float dpi = this.GetDpi();

            double pixels = Math.Round((float)centimeter / 2.54 * dpi , 2);

            this.txtPixels.Text = pixels.ToString();

            this.txtPoints.Text = ValueHelper.PixelsValueToPointsValue((decimal)pixels).ToString();
        }

        private void txtCentimeter_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string value = this.GetCleanValue(this.txtCentimeter.Text);

            if (value.Length > 0 && this.IsNumber(value))
            {
                this.ConvertCentimeterToPixels(Convert.ToDouble(value));
            }
        }
    }
}
