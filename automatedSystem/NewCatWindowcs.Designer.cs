namespace automatedSystem
{
    partial class NewCatWindowcs
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NewCatWindowcs));
            this.ContentsPanel = new System.Windows.Forms.Panel();
            this.liveDGV = new System.Windows.Forms.DataGridView();
            this.Catagory = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.headerPanel = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.backgroundPanel = new AForge.Controls.PictureBox();
            this.saveBtn = new System.Windows.Forms.Button();
            this.ResetBtn = new System.Windows.Forms.Button();
            this.ButtonPanel = new System.Windows.Forms.Panel();
            this.ContentsPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.liveDGV)).BeginInit();
            this.headerPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.backgroundPanel)).BeginInit();
            this.ButtonPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // ContentsPanel
            // 
            this.ContentsPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.ContentsPanel.Controls.Add(this.liveDGV);
            this.ContentsPanel.Location = new System.Drawing.Point(4, 57);
            this.ContentsPanel.Name = "ContentsPanel";
            this.ContentsPanel.Size = new System.Drawing.Size(422, 417);
            this.ContentsPanel.TabIndex = 87;
            // 
            // liveDGV
            // 
            this.liveDGV.AllowUserToOrderColumns = true;
            this.liveDGV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.liveDGV.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.liveDGV.BackgroundColor = System.Drawing.Color.Gray;
            this.liveDGV.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.liveDGV.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SunkenVertical;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.liveDGV.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.liveDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.liveDGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Catagory});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.LightGray;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.liveDGV.DefaultCellStyle = dataGridViewCellStyle2;
            this.liveDGV.GridColor = System.Drawing.Color.Gray;
            this.liveDGV.Location = new System.Drawing.Point(12, 14);
            this.liveDGV.Name = "liveDGV";
            this.liveDGV.Size = new System.Drawing.Size(400, 397);
            this.liveDGV.TabIndex = 1;
            this.liveDGV.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.liveDGV_CellValueChanged);
            // 
            // Catagory
            // 
            this.Catagory.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Catagory.HeaderText = "Catagory Name";
            this.Catagory.Name = "Catagory";
            this.Catagory.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Catagory.Width = 157;
            // 
            // headerPanel
            // 
            this.headerPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.headerPanel.Controls.Add(this.button1);
            this.headerPanel.Location = new System.Drawing.Point(4, 6);
            this.headerPanel.Name = "headerPanel";
            this.headerPanel.Size = new System.Drawing.Size(422, 45);
            this.headerPanel.TabIndex = 89;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.button1.BackgroundImage = global::automatedSystem.Properties.Resources.exit_button_icon_clipart_1;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Comic Sans MS", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(381, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(38, 35);
            this.button1.TabIndex = 88;
            this.button1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // backgroundPanel
            // 
            this.backgroundPanel.Image = null;
            this.backgroundPanel.Location = new System.Drawing.Point(-1, -2);
            this.backgroundPanel.Name = "backgroundPanel";
            this.backgroundPanel.Size = new System.Drawing.Size(433, 565);
            this.backgroundPanel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.backgroundPanel.TabIndex = 86;
            this.backgroundPanel.TabStop = false;
            // 
            // saveBtn
            // 
            this.saveBtn.BackColor = System.Drawing.Color.Gray;
            this.saveBtn.FlatAppearance.BorderSize = 0;
            this.saveBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.saveBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.saveBtn.Location = new System.Drawing.Point(226, 9);
            this.saveBtn.Margin = new System.Windows.Forms.Padding(2);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(173, 54);
            this.saveBtn.TabIndex = 182;
            this.saveBtn.Text = "Save";
            this.saveBtn.UseVisualStyleBackColor = false;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // ResetBtn
            // 
            this.ResetBtn.BackColor = System.Drawing.Color.Gray;
            this.ResetBtn.FlatAppearance.BorderSize = 0;
            this.ResetBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ResetBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.ResetBtn.Location = new System.Drawing.Point(32, 9);
            this.ResetBtn.Margin = new System.Windows.Forms.Padding(2);
            this.ResetBtn.Name = "ResetBtn";
            this.ResetBtn.Size = new System.Drawing.Size(173, 54);
            this.ResetBtn.TabIndex = 182;
            this.ResetBtn.Text = "Reset";
            this.ResetBtn.UseVisualStyleBackColor = false;
            this.ResetBtn.Click += new System.EventHandler(this.ResetBtn_Click);
            // 
            // ButtonPanel
            // 
            this.ButtonPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.ButtonPanel.Controls.Add(this.saveBtn);
            this.ButtonPanel.Controls.Add(this.ResetBtn);
            this.ButtonPanel.Location = new System.Drawing.Point(3, 482);
            this.ButtonPanel.Name = "ButtonPanel";
            this.ButtonPanel.Size = new System.Drawing.Size(422, 71);
            this.ButtonPanel.TabIndex = 90;
            // 
            // NewCatWindowcs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.ClientSize = new System.Drawing.Size(430, 558);
            this.Controls.Add(this.headerPanel);
            this.Controls.Add(this.ContentsPanel);
            this.Controls.Add(this.ButtonPanel);
            this.Controls.Add(this.backgroundPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "NewCatWindowcs";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Change Catagory";
            this.Load += new System.EventHandler(this.NewCatWindowcs_Load);
            this.ContentsPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.liveDGV)).EndInit();
            this.headerPanel.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.backgroundPanel)).EndInit();
            this.ButtonPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private AForge.Controls.PictureBox backgroundPanel;
        private System.Windows.Forms.Panel ContentsPanel;
        private System.Windows.Forms.Panel headerPanel;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView liveDGV;
        private System.Windows.Forms.DataGridViewTextBoxColumn Catagory;
        private System.Windows.Forms.Button saveBtn;
        private System.Windows.Forms.Button ResetBtn;
        private System.Windows.Forms.Panel ButtonPanel;
    }
}