namespace automatedSystem
{
    partial class changeCam
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(changeCam));
            this.CameraListBox = new System.Windows.Forms.ListBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.backgoundPanel = new System.Windows.Forms.Panel();
            this.saveBtn = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CameraListBox
            // 
            this.CameraListBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.CameraListBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CameraListBox.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.CameraListBox.ForeColor = System.Drawing.SystemColors.Window;
            this.CameraListBox.FormattingEnabled = true;
            this.CameraListBox.ItemHeight = 23;
            this.CameraListBox.Location = new System.Drawing.Point(16, 72);
            this.CameraListBox.Name = "CameraListBox";
            this.CameraListBox.Size = new System.Drawing.Size(325, 299);
            this.CameraListBox.TabIndex = 0;
            this.CameraListBox.SelectedIndexChanged += new System.EventHandler(this.CameraListBox_SelectedIndexChanged);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.saveBtn);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.CameraListBox);
            this.panel1.Location = new System.Drawing.Point(205, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(362, 426);
            this.panel1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.label1.ForeColor = System.Drawing.SystemColors.Window;
            this.label1.Location = new System.Drawing.Point(3, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(166, 23);
            this.label1.TabIndex = 1;
            this.label1.Text = "Select a Camera :";
            // 
            // backgoundPanel
            // 
            this.backgoundPanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.backgoundPanel.Location = new System.Drawing.Point(197, 5);
            this.backgoundPanel.Name = "backgoundPanel";
            this.backgoundPanel.Size = new System.Drawing.Size(377, 441);
            this.backgoundPanel.TabIndex = 2;
            // 
            // saveBtn
            // 
            this.saveBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.saveBtn.ForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.saveBtn.Location = new System.Drawing.Point(249, 377);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(92, 32);
            this.saveBtn.TabIndex = 4;
            this.saveBtn.Text = "Save";
            this.saveBtn.UseVisualStyleBackColor = true;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.button2.BackgroundImage = global::automatedSystem.Properties.Resources.exit_button_icon_clipart_1;
            this.button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("Comic Sans MS", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(324, 1);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(38, 35);
            this.button2.TabIndex = 89;
            this.button2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // changeCam
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SeaGreen;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.backgoundPanel);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "changeCam";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Change Camera";
            this.TransparencyKey = System.Drawing.Color.SeaGreen;
            this.Load += new System.EventHandler(this.changeCam_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox CameraListBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel backgoundPanel;
        private System.Windows.Forms.Button saveBtn;
        private System.Windows.Forms.Button button2;
    }
}