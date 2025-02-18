namespace LogApp_v1
{
    partial class Form4
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            this.pnl_pulltktsys2 = new System.Windows.Forms.Panel();
            this.btnasper_done = new System.Windows.Forms.Button();
            this.btnasper_mark = new System.Windows.Forms.Button();
            this.btnasper_selall = new System.Windows.Forms.Button();
            this.asperfilterTxt = new System.Windows.Forms.TextBox();
            this.asperfilterBox = new System.Windows.Forms.ComboBox();
            this.seach = new System.Windows.Forms.Label();
            this.asperfilterdatagrid = new System.Windows.Forms.DataGridView();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.timer2 = new System.Windows.Forms.Timer(this.components);
            this.pnl_pulltktsys2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.asperfilterdatagrid)).BeginInit();
            this.SuspendLayout();
            // 
            // pnl_pulltktsys2
            // 
            this.pnl_pulltktsys2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnl_pulltktsys2.BackColor = System.Drawing.Color.DimGray;
            this.pnl_pulltktsys2.Controls.Add(this.btnasper_done);
            this.pnl_pulltktsys2.Controls.Add(this.btnasper_mark);
            this.pnl_pulltktsys2.Controls.Add(this.btnasper_selall);
            this.pnl_pulltktsys2.Controls.Add(this.asperfilterTxt);
            this.pnl_pulltktsys2.Controls.Add(this.asperfilterBox);
            this.pnl_pulltktsys2.Controls.Add(this.seach);
            this.pnl_pulltktsys2.Location = new System.Drawing.Point(0, 610);
            this.pnl_pulltktsys2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pnl_pulltktsys2.Name = "pnl_pulltktsys2";
            this.pnl_pulltktsys2.Size = new System.Drawing.Size(1264, 107);
            this.pnl_pulltktsys2.TabIndex = 1;
            // 
            // btnasper_done
            // 
            this.btnasper_done.BackColor = System.Drawing.Color.SteelBlue;
            this.btnasper_done.FlatAppearance.BorderSize = 0;
            this.btnasper_done.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnasper_done.Font = new System.Drawing.Font("Microsoft YaHei", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnasper_done.ForeColor = System.Drawing.Color.White;
            this.btnasper_done.Location = new System.Drawing.Point(1115, 34);
            this.btnasper_done.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnasper_done.Name = "btnasper_done";
            this.btnasper_done.Size = new System.Drawing.Size(123, 41);
            this.btnasper_done.TabIndex = 9;
            this.btnasper_done.Text = "DONE";
            this.btnasper_done.UseVisualStyleBackColor = false;
            this.btnasper_done.Click += new System.EventHandler(this.btnasper_done_Click);
            // 
            // btnasper_mark
            // 
            this.btnasper_mark.BackColor = System.Drawing.Color.SteelBlue;
            this.btnasper_mark.FlatAppearance.BorderSize = 0;
            this.btnasper_mark.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnasper_mark.Font = new System.Drawing.Font("Microsoft YaHei", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnasper_mark.ForeColor = System.Drawing.Color.White;
            this.btnasper_mark.Location = new System.Drawing.Point(653, 34);
            this.btnasper_mark.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnasper_mark.Name = "btnasper_mark";
            this.btnasper_mark.Size = new System.Drawing.Size(123, 41);
            this.btnasper_mark.TabIndex = 8;
            this.btnasper_mark.Text = "MARK";
            this.btnasper_mark.UseVisualStyleBackColor = false;
            this.btnasper_mark.Click += new System.EventHandler(this.btnasper_mark_Click);
            // 
            // btnasper_selall
            // 
            this.btnasper_selall.BackColor = System.Drawing.Color.SteelBlue;
            this.btnasper_selall.FlatAppearance.BorderSize = 0;
            this.btnasper_selall.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnasper_selall.Font = new System.Drawing.Font("Microsoft YaHei", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnasper_selall.ForeColor = System.Drawing.Color.White;
            this.btnasper_selall.Location = new System.Drawing.Point(509, 34);
            this.btnasper_selall.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnasper_selall.Name = "btnasper_selall";
            this.btnasper_selall.Size = new System.Drawing.Size(123, 41);
            this.btnasper_selall.TabIndex = 7;
            this.btnasper_selall.Text = "SELECT ALL";
            this.btnasper_selall.UseVisualStyleBackColor = false;
            this.btnasper_selall.Click += new System.EventHandler(this.btnasper_selall_Click);
            // 
            // asperfilterTxt
            // 
            this.asperfilterTxt.Location = new System.Drawing.Point(261, 54);
            this.asperfilterTxt.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.asperfilterTxt.Name = "asperfilterTxt";
            this.asperfilterTxt.Size = new System.Drawing.Size(199, 22);
            this.asperfilterTxt.TabIndex = 6;
            this.asperfilterTxt.TextChanged += new System.EventHandler(this.asperfilterTxt_TextChanged);
            // 
            // asperfilterBox
            // 
            this.asperfilterBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.asperfilterBox.FormattingEnabled = true;
            this.asperfilterBox.Items.AddRange(new object[] {
            "ALL",
            "FACILITY",
            "PARTNUMBER",
            "PULLTICKET",
            "REMARKS"});
            this.asperfilterBox.Location = new System.Drawing.Point(44, 52);
            this.asperfilterBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.asperfilterBox.Name = "asperfilterBox";
            this.asperfilterBox.Size = new System.Drawing.Size(211, 24);
            this.asperfilterBox.TabIndex = 5;
            
            // 
            // seach
            // 
            this.seach.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.seach.AutoSize = true;
            this.seach.Font = new System.Drawing.Font("Microsoft YaHei UI", 12F, System.Drawing.FontStyle.Bold);
            this.seach.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.seach.Location = new System.Drawing.Point(39, 14);
            this.seach.Name = "seach";
            this.seach.Size = new System.Drawing.Size(132, 27);
            this.seach.TabIndex = 1;
            this.seach.Text = "SEARCH BY:";
            // 
            // asperfilterdatagrid
            // 
            dataGridViewCellStyle6.BackColor = System.Drawing.Color.SteelBlue;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle6.ForeColor = System.Drawing.Color.White;
            this.asperfilterdatagrid.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle6;
            this.asperfilterdatagrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft YaHei UI", 9.75F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.asperfilterdatagrid.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.asperfilterdatagrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle8.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft YaHei UI", 9.75F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.asperfilterdatagrid.DefaultCellStyle = dataGridViewCellStyle8;
            this.asperfilterdatagrid.GridColor = System.Drawing.Color.SkyBlue;
            this.asperfilterdatagrid.Location = new System.Drawing.Point(5, 42);
            this.asperfilterdatagrid.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.asperfilterdatagrid.Name = "asperfilterdatagrid";
            this.asperfilterdatagrid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle9.Font = new System.Drawing.Font("Microsoft YaHei UI", 9.75F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.asperfilterdatagrid.RowHeadersDefaultCellStyle = dataGridViewCellStyle9;
            this.asperfilterdatagrid.RowHeadersVisible = false;
            this.asperfilterdatagrid.RowHeadersWidth = 51;
            dataGridViewCellStyle10.Font = new System.Drawing.Font("Microsoft YaHei UI", 9.75F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle10.SelectionBackColor = System.Drawing.Color.Navy;
            this.asperfilterdatagrid.RowsDefaultCellStyle = dataGridViewCellStyle10;
            this.asperfilterdatagrid.RowTemplate.Height = 24;
            this.asperfilterdatagrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.asperfilterdatagrid.Size = new System.Drawing.Size(1252, 562);
            this.asperfilterdatagrid.TabIndex = 2;
            this.asperfilterdatagrid.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.asperfilterdatagrid_CellMouseClick);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 45;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // timer2
            // 
            this.timer2.Interval = 45;
            this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DimGray;
            this.ClientSize = new System.Drawing.Size(1264, 718);
            this.Controls.Add(this.asperfilterdatagrid);
            this.Controls.Add(this.pnl_pulltktsys2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form4";
            this.Opacity = 0D;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form4";
            this.Load += new System.EventHandler(this.Form4_Load);
            this.pnl_pulltktsys2.ResumeLayout(false);
            this.pnl_pulltktsys2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.asperfilterdatagrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel pnl_pulltktsys2;
        private System.Windows.Forms.Label seach;
        private System.Windows.Forms.ComboBox asperfilterBox;
        private System.Windows.Forms.TextBox asperfilterTxt;
        private System.Windows.Forms.Button btnasper_done;
        private System.Windows.Forms.Button btnasper_mark;
        private System.Windows.Forms.Button btnasper_selall;
        public System.Windows.Forms.DataGridView asperfilterdatagrid;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Timer timer2;
    }
}