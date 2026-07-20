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
    public partial class frmColorTransform : Form
    {
        public frmColorTransform()
        {
            InitializeComponent();
        }

        private void btnTransform_Click(object sender, EventArgs e)
        {
            string colorHex = this.txtColorHex.Text.Trim();
            string luminanceModulation = this.txtLuminanceModulation.Text.Trim();
            string luminanceOffset = this.txtLuminanceOffset.Text.Trim();

            if (string.IsNullOrEmpty(colorHex))
            {
                MessageBox.Show("Please enter a color hex value.");
                return;
            }

            if (ColorHelper.IsColorModelHex(colorHex) == false)
            {
                MessageBox.Show("Please enter a valid color hex value.");
                return;
            }

            double? luminanceModulationValue = default(double?);
            double? luminanceOffsetValue = default(double?);

            if (!string.IsNullOrEmpty(luminanceModulation))
            {
                if (!double.TryParse(luminanceModulation, out _))
                {
                    MessageBox.Show("Please enter a valid number for luminance modulation.");
                    return;
                }

                luminanceModulationValue = double.Parse(luminanceModulation);
            }


            if (!string.IsNullOrEmpty(luminanceOffset))
            {
                if (!double.TryParse(luminanceOffset, out _))
                {
                    MessageBox.Show("Please enter a valid number for luminance offset.");
                    return;
                }

                luminanceOffsetValue = double.Parse(luminanceOffset);
            }

            try
            {
                string transformedColor = colorHex;

                if (luminanceModulationValue.HasValue && luminanceModulationValue!=1)
                {
                    transformedColor = ColorHelper.TransformLumMod(colorHex, (long)luminanceModulationValue);
                }

                if (luminanceOffsetValue.HasValue && luminanceOffsetValue != 0)
                {
                    transformedColor = ColorHelper.TransformLumMod(transformedColor, (long)luminanceOffsetValue);
                }                

                this.txtResult.Text = transformedColor.ToUpper();                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            this.txtColorHex.Text = "";
            this.txtLuminanceModulation.Text = "";
            this.txtLuminanceOffset.Text = "";
            this.txtResult.Text = "";          
        }
    }
}
