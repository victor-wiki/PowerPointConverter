namespace PowerPointViewer
{
    partial class frmUnitConversion
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            label1 = new Label();
            txtValue = new TextBox();
            txtPoints = new TextBox();
            label2 = new Label();
            txtPixels = new TextBox();
            label3 = new Label();
            btnConvert = new Button();
            txtCentimeter = new TextBox();
            label4 = new Label();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(23, 21);
            label1.Name = "label1";
            label1.Size = new Size(43, 17);
            label1.TabIndex = 0;
            label1.Text = "Value:";
            // 
            // txtValue
            // 
            txtValue.Location = new Point(101, 18);
            txtValue.Name = "txtValue";
            txtValue.Size = new Size(196, 23);
            txtValue.TabIndex = 1;
            // 
            // txtPoints
            // 
            txtPoints.Location = new Point(101, 57);
            txtPoints.Name = "txtPoints";
            txtPoints.Size = new Size(196, 23);
            txtPoints.TabIndex = 3;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(23, 60);
            label2.Name = "label2";
            label2.Size = new Size(46, 17);
            label2.TabIndex = 2;
            label2.Text = "Points:";
            // 
            // txtPixels
            // 
            txtPixels.Location = new Point(101, 95);
            txtPixels.Name = "txtPixels";
            txtPixels.Size = new Size(196, 23);
            txtPixels.TabIndex = 5;
            txtPixels.MouseDoubleClick += txtPixels_MouseDoubleClick;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(23, 98);
            label3.Name = "label3";
            label3.Size = new Size(43, 17);
            label3.TabIndex = 4;
            label3.Text = "Pixels:";
            // 
            // btnConvert
            // 
            btnConvert.Location = new Point(126, 181);
            btnConvert.Name = "btnConvert";
            btnConvert.Size = new Size(75, 23);
            btnConvert.TabIndex = 6;
            btnConvert.Text = "Convert";
            btnConvert.UseVisualStyleBackColor = true;
            btnConvert.Click += btnConvert_Click;
            // 
            // txtCentimeter
            // 
            txtCentimeter.Location = new Point(101, 133);
            txtCentimeter.Name = "txtCentimeter";
            txtCentimeter.Size = new Size(196, 23);
            txtCentimeter.TabIndex = 8;
            txtCentimeter.MouseDoubleClick += txtCentimeter_MouseDoubleClick;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(23, 136);
            label4.Name = "label4";
            label4.Size = new Size(74, 17);
            label4.TabIndex = 7;
            label4.Text = "Centimeter:";
            // 
            // frmUnitConversion
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(340, 223);
            Controls.Add(txtCentimeter);
            Controls.Add(label4);
            Controls.Add(btnConvert);
            Controls.Add(txtPixels);
            Controls.Add(label3);
            Controls.Add(txtPoints);
            Controls.Add(label2);
            Controls.Add(txtValue);
            Controls.Add(label1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "frmUnitConversion";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Unit Conversion";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private TextBox txtValue;
        private TextBox txtPoints;
        private Label label2;
        private TextBox txtPixels;
        private Label label3;
        private Button btnConvert;
        private TextBox txtCentimeter;
        private Label label4;
    }
}