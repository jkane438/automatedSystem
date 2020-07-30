namespace automatedSystem
{
    partial class changeStringFormcs
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(changeStringFormcs));
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.saveBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.stringInput = new System.Windows.Forms.MaskedTextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.backgroundPanel = new AForge.Controls.PictureBox();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.backgroundPanel)).BeginInit();
            this.SuspendLayout();
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.panel2.Controls.Add(this.button2);
            this.panel2.Location = new System.Drawing.Point(1, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(422, 45);
            this.panel2.TabIndex = 92;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.panel1.Controls.Add(this.saveBtn);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.stringInput);
            this.panel1.Location = new System.Drawing.Point(2, 51);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(422, 434);
            this.panel1.TabIndex = 91;
            // 
            // saveBtn
            // 
            this.saveBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.saveBtn.Location = new System.Drawing.Point(300, 167);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(92, 32);
            this.saveBtn.TabIndex = 2;
            this.saveBtn.Text = "Save";
            this.saveBtn.UseVisualStyleBackColor = true;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.label1.ForeColor = System.Drawing.SystemColors.Window;
            this.label1.Location = new System.Drawing.Point(6, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(343, 23);
            this.label1.TabIndex = 1;
            this.label1.Text = "Please Enter New Default Name Here:";
            // 
            // stringInput
            // 
            this.stringInput.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.stringInput.Location = new System.Drawing.Point(10, 122);
            this.stringInput.Name = "stringInput";
            this.stringInput.Size = new System.Drawing.Size(382, 30);
            this.stringInput.TabIndex = 0;
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
            this.button2.Location = new System.Drawing.Point(381, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(38, 35);
            this.button2.TabIndex = 88;
            this.button2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // backgroundPanel
            // 
            this.backgroundPanel.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.backgroundPanel.Image = null;
            this.backgroundPanel.Location = new System.Drawing.Point(-1, 0);
            this.backgroundPanel.Name = "backgroundPanel";
            this.backgroundPanel.Size = new System.Drawing.Size(433, 493);
            this.backgroundPanel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.backgroundPanel.TabIndex = 89;
            this.backgroundPanel.TabStop = false;
            // 
            // changeStringFormcs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Pink;
            this.ClientSize = new System.Drawing.Size(427, 491);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.backgroundPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "changeStringFormcs";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Change Default Name";
            this.Load += new System.EventHandler(this.changeStringFormcs_Load);
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.backgroundPanel)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private AForge.Controls.PictureBox backgroundPanel;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.MaskedTextBox stringInput;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button saveBtn;
    }
}