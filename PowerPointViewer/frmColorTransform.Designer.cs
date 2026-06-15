namespace PowerPointViewer
{
    partial class frmColorTransform
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
            label2 = new Label();
            label3 = new Label();
            txtColorHex = new TextBox();
            txtLuminanceModulation = new TextBox();
            txtLuminanceOffset = new TextBox();
            btnTransform = new Button();
            btnReset = new Button();
            txtResult = new TextBox();
            label4 = new Label();
            label5 = new Label();
            label6 = new Label();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(21, 20);
            label1.Name = "label1";
            label1.Size = new Size(69, 17);
            label1.TabIndex = 0;
            label1.Text = "Color Hex:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(21, 56);
            label2.Name = "label2";
            label2.Size = new Size(143, 17);
            label2.TabIndex = 1;
            label2.Text = "Luminance Modulation:";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(21, 92);
            label3.Name = "label3";
            label3.Size = new Size(111, 17);
            label3.TabIndex = 2;
            label3.Text = "Luminance Offset:";
            // 
            // txtColorHex
            // 
            txtColorHex.Location = new Point(167, 17);
            txtColorHex.Name = "txtColorHex";
            txtColorHex.Size = new Size(118, 23);
            txtColorHex.TabIndex = 3;
            // 
            // txtLuminanceModulation
            // 
            txtLuminanceModulation.Location = new Point(167, 53);
            txtLuminanceModulation.Name = "txtLuminanceModulation";
            txtLuminanceModulation.Size = new Size(118, 23);
            txtLuminanceModulation.TabIndex = 4;
            // 
            // txtLuminanceOffset
            // 
            txtLuminanceOffset.Location = new Point(167, 88);
            txtLuminanceOffset.Name = "txtLuminanceOffset";
            txtLuminanceOffset.Size = new Size(118, 23);
            txtLuminanceOffset.TabIndex = 5;
            // 
            // btnTransform
            // 
            btnTransform.Location = new Point(131, 187);
            btnTransform.Name = "btnTransform";
            btnTransform.Size = new Size(88, 23);
            btnTransform.TabIndex = 6;
            btnTransform.Text = "Transform";
            btnTransform.UseVisualStyleBackColor = true;
            btnTransform.Click += btnTransform_Click;
            // 
            // btnReset
            // 
            btnReset.Location = new Point(229, 187);
            btnReset.Name = "btnReset";
            btnReset.Size = new Size(75, 23);
            btnReset.TabIndex = 7;
            btnReset.Text = "Reset";
            btnReset.UseVisualStyleBackColor = true;
            btnReset.Click += btnReset_Click;
            // 
            // txtResult
            // 
            txtResult.Location = new Point(167, 126);
            txtResult.Name = "txtResult";
            txtResult.Size = new Size(118, 23);
            txtResult.TabIndex = 9;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(21, 128);
            label4.Name = "label4";
            label4.Size = new Size(46, 17);
            label4.TabIndex = 8;
            label4.Text = "Result:";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(302, 57);
            label5.Name = "label5";
            label5.Size = new Size(160, 17);
            label5.TabIndex = 10;
            label5.Text = "(value is between 0 and 1)";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(302, 90);
            label6.Name = "label6";
            label6.Size = new Size(160, 17);
            label6.TabIndex = 11;
            label6.Text = "(value is between 0 and 1)";
            // 
            // frmColorTransform
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(465, 231);
            Controls.Add(label6);
            Controls.Add(label5);
            Controls.Add(txtResult);
            Controls.Add(label4);
            Controls.Add(btnReset);
            Controls.Add(btnTransform);
            Controls.Add(txtLuminanceOffset);
            Controls.Add(txtLuminanceModulation);
            Controls.Add(txtColorHex);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            MaximizeBox = false;
            Name = "frmColorTransform";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Color Transform";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Label label2;
        private Label label3;
        private TextBox txtColorHex;
        private TextBox txtLuminanceModulation;
        private TextBox txtLuminanceOffset;
        private Button btnTransform;
        private Button btnReset;
        private TextBox txtResult;
        private Label label4;
        private Label label5;
        private Label label6;
    }
}