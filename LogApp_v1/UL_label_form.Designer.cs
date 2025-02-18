namespace LogApp_v1
{
    partial class UL_label_form
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
            this.ul_label_panel = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.updateLblBtn = new System.Windows.Forms.Button();
            this.closeImgBtn = new System.Windows.Forms.Panel();
            this.ul_label_panel.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // ul_label_panel
            // 
            this.ul_label_panel.BackColor = System.Drawing.Color.Transparent;
            this.ul_label_panel.Controls.Add(this.panel1);
            this.ul_label_panel.Controls.Add(this.closeImgBtn);
            this.ul_label_panel.Location = new System.Drawing.Point(0, 0);
            this.ul_label_panel.Margin = new System.Windows.Forms.Padding(4);
            this.ul_label_panel.Name = "ul_label_panel";
            this.ul_label_panel.Size = new System.Drawing.Size(432, 395);
            this.ul_label_panel.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.SteelBlue;
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.updateLblBtn);
            this.panel1.Location = new System.Drawing.Point(5, 80);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(315, 213);
            this.panel1.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(101)))), ((int)(((byte)(114)))));
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Microsoft YaHei UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.SystemColors.Control;
            this.button2.Location = new System.Drawing.Point(36, 107);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(236, 65);
            this.button2.TabIndex = 1;
            this.button2.Text = "IMPORT AXMR432A";
            this.button2.UseVisualStyleBackColor = false;
            // 
            // updateLblBtn
            // 
            this.updateLblBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(101)))), ((int)(((byte)(114)))));
            this.updateLblBtn.FlatAppearance.BorderSize = 0;
            this.updateLblBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.updateLblBtn.Font = new System.Drawing.Font("Microsoft YaHei UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.updateLblBtn.ForeColor = System.Drawing.SystemColors.Control;
            this.updateLblBtn.Location = new System.Drawing.Point(36, 33);
            this.updateLblBtn.Margin = new System.Windows.Forms.Padding(4);
            this.updateLblBtn.Name = "updateLblBtn";
            this.updateLblBtn.Size = new System.Drawing.Size(236, 65);
            this.updateLblBtn.TabIndex = 0;
            this.updateLblBtn.Text = "UPDATE UL LABEL";
            this.updateLblBtn.UseVisualStyleBackColor = false;
            this.updateLblBtn.Click += new System.EventHandler(this.updateLblBtn_Click);
            // 
            // closeImgBtn
            // 
            this.closeImgBtn.BackgroundImage = global::LogApp_v1.Properties.Resources.icons8_triangle_80;
            this.closeImgBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.closeImgBtn.Location = new System.Drawing.Point(219, -4);
            this.closeImgBtn.Margin = new System.Windows.Forms.Padding(4);
            this.closeImgBtn.Name = "closeImgBtn";
            this.closeImgBtn.Size = new System.Drawing.Size(112, 82);
            this.closeImgBtn.TabIndex = 3;
            this.closeImgBtn.Click += new System.EventHandler(this.closeImgBtn_Click);
            // 
            // UL_label_form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ClientSize = new System.Drawing.Size(432, 392);
            this.ControlBox = false;
            this.Controls.Add(this.ul_label_panel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UL_label_form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "UL_label_form";
            this.TransparencyKey = System.Drawing.Color.Transparent;
            this.ul_label_panel.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel ul_label_panel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button updateLblBtn;
        private System.Windows.Forms.Panel closeImgBtn;
    }
}