namespace automatedSystem
{
    partial class FullScreenForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FullScreenForm));
            this.FullScreenPictureBox = new AForge.Controls.PictureBox();
            this.cameraManualLB = new System.Windows.Forms.ListBox();
            this.liveDGV = new System.Windows.Forms.DataGridView();
            this.angleCol = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.angleInput = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Catagory = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.closeBtn3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.FullScreenPictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.liveDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // FullScreenPictureBox
            // 
            this.FullScreenPictureBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FullScreenPictureBox.Image = null;
            this.FullScreenPictureBox.Location = new System.Drawing.Point(0, 0);
            this.FullScreenPictureBox.Name = "FullScreenPictureBox";
            this.FullScreenPictureBox.Size = new System.Drawing.Size(800, 450);
            this.FullScreenPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.FullScreenPictureBox.TabIndex = 0;
            this.FullScreenPictureBox.TabStop = false;
            this.FullScreenPictureBox.Click += new System.EventHandler(this.FullScreenPictureBox_Click);
            this.FullScreenPictureBox.Paint += new System.Windows.Forms.PaintEventHandler(this.FullScreenPictureBox_Paint);
            this.FullScreenPictureBox.MouseDown += new System.Windows.Forms.MouseEventHandler(this.FullScreenPictureBox_MouseDown);
            this.FullScreenPictureBox.MouseEnter += new System.EventHandler(this.FullScreenPictureBox_MouseEnter);
            this.FullScreenPictureBox.MouseLeave += new System.EventHandler(this.FullScreenPictureBox_MouseLeave);
            // 
            // cameraManualLB
            // 
            this.cameraManualLB.BackColor = System.Drawing.Color.Gray;
            this.cameraManualLB.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.cameraManualLB.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cameraManualLB.FormattingEnabled = true;
            this.cameraManualLB.ItemHeight = 23;
            this.cameraManualLB.Location = new System.Drawing.Point(229, 214);
            this.cameraManualLB.Margin = new System.Windows.Forms.Padding(2);
            this.cameraManualLB.Name = "cameraManualLB";
            this.cameraManualLB.Size = new System.Drawing.Size(342, 23);
            this.cameraManualLB.TabIndex = 89;
            this.cameraManualLB.Visible = false;
            // 
            // liveDGV
            // 
            this.liveDGV.AllowUserToAddRows = false;
            this.liveDGV.AllowUserToDeleteRows = false;
            this.liveDGV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.liveDGV.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.liveDGV.BackgroundColor = System.Drawing.Color.Gray;
            this.liveDGV.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.liveDGV.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SunkenVertical;
            this.liveDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.liveDGV.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.angleCol,
            this.angleInput,
            this.Catagory});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.liveDGV.DefaultCellStyle = dataGridViewCellStyle1;
            this.liveDGV.GridColor = System.Drawing.Color.Gray;
            this.liveDGV.Location = new System.Drawing.Point(313, 93);
            this.liveDGV.Name = "liveDGV";
            this.liveDGV.Size = new System.Drawing.Size(543, 581);
            this.liveDGV.TabIndex = 91;
            // 
            // angleCol
            // 
            this.angleCol.HeaderText = "Angle";
            this.angleCol.Name = "angleCol";
            this.angleCol.ReadOnly = true;
            this.angleCol.Width = 59;
            // 
            // angleInput
            // 
            this.angleInput.HeaderText = " ";
            this.angleInput.Name = "angleInput";
            this.angleInput.ReadOnly = true;
            this.angleInput.Width = 35;
            // 
            // Catagory
            // 
            this.Catagory.HeaderText = "Catagory";
            this.Catagory.Items.AddRange(new object[] {
            "Angle Of Head Connection (1)",
            "Angle Of Head Connection (2)",
            "Angle Of Midpiece"});
            this.Catagory.Name = "Catagory";
            this.Catagory.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Catagory.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.Catagory.Width = 74;
            // 
            // closeBtn3
            // 
            this.closeBtn3.BackColor = System.Drawing.Color.Transparent;
            this.closeBtn3.BackgroundImage = global::automatedSystem.Properties.Resources.exit_button_icon_clipart_1;
            this.closeBtn3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.closeBtn3.FlatAppearance.BorderSize = 0;
            this.closeBtn3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeBtn3.Font = new System.Drawing.Font("Comic Sans MS", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.closeBtn3.Location = new System.Drawing.Point(753, 0);
            this.closeBtn3.Name = "closeBtn3";
            this.closeBtn3.Size = new System.Drawing.Size(47, 41);
            this.closeBtn3.TabIndex = 93;
            this.closeBtn3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.closeBtn3.UseVisualStyleBackColor = false;
            this.closeBtn3.Click += new System.EventHandler(this.CloseBtn3_Click);
            // 
            // FullScreenForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.closeBtn3);
            this.Controls.Add(this.FullScreenPictureBox);
            this.Controls.Add(this.cameraManualLB);
            this.Controls.Add(this.liveDGV);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FullScreenForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "fullScreenForm";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Activated += new System.EventHandler(this.FullScreenForm_Activated);
            this.Deactivate += new System.EventHandler(this.FullScreenForm_Deactivate);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FullScreenForm_FormClosed);
            this.Load += new System.EventHandler(this.FullScreenForm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FullScreenForm_KeyDown);
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FullScreenForm_KeyPress);
            ((System.ComponentModel.ISupportInitialize)(this.FullScreenPictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.liveDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AForge.Controls.PictureBox FullScreenPictureBox;
        private System.Windows.Forms.ListBox cameraManualLB;
        private System.Windows.Forms.DataGridView liveDGV;
        private System.Windows.Forms.DataGridViewTextBoxColumn angleCol;
        private System.Windows.Forms.DataGridViewTextBoxColumn angleInput;
        private System.Windows.Forms.DataGridViewComboBoxColumn Catagory;
        private System.Windows.Forms.Button closeBtn3;
    }
}