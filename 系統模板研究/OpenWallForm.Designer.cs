namespace Modeling
{
    partial class OpenWallForm
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
            this.cb_level = new System.Windows.Forms.ComboBox();
            this.Text_Level = new System.Windows.Forms.Label();
            this.Text_kerb = new System.Windows.Forms.TextBox();
            this.Text1 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_OK = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.GridSizeText = new System.Windows.Forms.Label();
            this.GridSizeNumber = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cb_level
            // 
            this.cb_level.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cb_level.FormattingEnabled = true;
            this.cb_level.Location = new System.Drawing.Point(143, 112);
            this.cb_level.Name = "cb_level";
            this.cb_level.Size = new System.Drawing.Size(130, 32);
            this.cb_level.TabIndex = 3;
            this.cb_level.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // Text_Level
            // 
            this.Text_Level.AutoSize = true;
            this.Text_Level.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Text_Level.Location = new System.Drawing.Point(42, 115);
            this.Text_Level.Name = "Text_Level";
            this.Text_Level.Size = new System.Drawing.Size(57, 24);
            this.Text_Level.TabIndex = 4;
            this.Text_Level.Text = "Level";
            this.Text_Level.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Text_kerb
            // 
            this.Text_kerb.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Text_kerb.Location = new System.Drawing.Point(144, 174);
            this.Text_kerb.Name = "Text_kerb";
            this.Text_kerb.Size = new System.Drawing.Size(130, 33);
            this.Text_kerb.TabIndex = 5;
            this.Text_kerb.TextChanged += new System.EventHandler(this.textBox1_TextChanged_1);
            // 
            // Text1
            // 
            this.Text1.AutoSize = true;
            this.Text1.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Text1.Location = new System.Drawing.Point(14, 183);
            this.Text1.Name = "Text1";
            this.Text1.Size = new System.Drawing.Size(118, 24);
            this.Text1.TabIndex = 6;
            this.Text1.Text = "Kerb Height";
            this.Text1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(70, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(203, 27);
            this.label1.TabIndex = 7;
            this.label1.Text = "Opened Wall Editor";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btn_OK
            // 
            this.btn_OK.Location = new System.Drawing.Point(75, 292);
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.Size = new System.Drawing.Size(75, 23);
            this.btn_OK.TabIndex = 8;
            this.btn_OK.Text = "OK";
            this.btn_OK.UseVisualStyleBackColor = true;
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(187, 292);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_Cancel.TabIndex = 9;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(287, 183);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 24);
            this.label2.TabIndex = 10;
            this.label2.Text = "cm";
            // 
            // GridSizeText
            // 
            this.GridSizeText.AutoSize = true;
            this.GridSizeText.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.GridSizeText.Location = new System.Drawing.Point(12, 251);
            this.GridSizeText.Name = "GridSizeText";
            this.GridSizeText.Size = new System.Drawing.Size(120, 24);
            this.GridSizeText.TabIndex = 11;
            this.GridSizeText.Text = "Gridline size";
            this.GridSizeText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.GridSizeText.Click += new System.EventHandler(this.label3_Click);
            // 
            // GridSizeNumber
            // 
            this.GridSizeNumber.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.GridSizeNumber.Location = new System.Drawing.Point(144, 242);
            this.GridSizeNumber.Name = "GridSizeNumber";
            this.GridSizeNumber.Size = new System.Drawing.Size(130, 33);
            this.GridSizeNumber.TabIndex = 12;
            this.GridSizeNumber.TextChanged += new System.EventHandler(this.GridSizeNumber_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(280, 251);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 24);
            this.label4.TabIndex = 13;
            this.label4.Text = "mm";
            // 
            // OpenWallForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(336, 317);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.GridSizeNumber);
            this.Controls.Add(this.GridSizeText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_OK);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Text1);
            this.Controls.Add(this.Text_kerb);
            this.Controls.Add(this.Text_Level);
            this.Controls.Add(this.cb_level);
            this.Name = "OpenWallForm";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox cb_level;
        private System.Windows.Forms.Label Text_Level;
        private System.Windows.Forms.TextBox Text_kerb;
        private System.Windows.Forms.Label Text1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_OK;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label GridSizeText;
        private System.Windows.Forms.TextBox GridSizeNumber;
        private System.Windows.Forms.Label label4;
    }
}