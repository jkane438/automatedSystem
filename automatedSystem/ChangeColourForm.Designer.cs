namespace automatedSystem
{
    partial class ChangeColourForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChangeColourForm));
#pragma warning disable CS0436 // Type conflicts with imported type
            this.colorWheel1 = new Cyotek.Windows.Forms.ColorWheel();
#pragma warning restore CS0436 // Type conflicts with imported type
            this.ExamplePanel = new System.Windows.Forms.Panel();
            this.backgroundPanel = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.panel5 = new System.Windows.Forms.Panel();
            this.saveBtn = new System.Windows.Forms.Button();
            this.resetBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.backgroundPanel.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // colorWheel1
            // 
            this.colorWheel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.colorWheel1.Color = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.colorWheel1.Location = new System.Drawing.Point(0, 0);
            this.colorWheel1.Name = "colorWheel1";
            this.colorWheel1.Size = new System.Drawing.Size(336, 310);
            this.colorWheel1.TabIndex = 0;
            this.colorWheel1.ColorChanged += new System.EventHandler(this.colorWheel1_ColorChanged);
            // 
            // ExamplePanel
            // 
            this.ExamplePanel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.ExamplePanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.ExamplePanel.Location = new System.Drawing.Point(342, 0);
            this.ExamplePanel.Name = "ExamplePanel";
            this.ExamplePanel.Size = new System.Drawing.Size(367, 310);
            this.ExamplePanel.TabIndex = 1;
            // 
            // backgroundPanel
            // 
            this.backgroundPanel.BackColor = System.Drawing.Color.GreenYellow;
            this.backgroundPanel.Controls.Add(this.panel5);
            this.backgroundPanel.Controls.Add(this.panel3);
            this.backgroundPanel.Controls.Add(this.panel4);
            this.backgroundPanel.Controls.Add(this.panel2);
            this.backgroundPanel.Location = new System.Drawing.Point(3, -2);
            this.backgroundPanel.Name = "backgroundPanel";
            this.backgroundPanel.Size = new System.Drawing.Size(715, 564);
            this.backgroundPanel.TabIndex = 2;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.panel2.Controls.Add(this.ExamplePanel);
            this.panel2.Controls.Add(this.colorWheel1);
            this.panel2.Location = new System.Drawing.Point(3, 154);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(709, 310);
            this.panel2.TabIndex = 1;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.panel3.Controls.Add(this.label1);
            this.panel3.Location = new System.Drawing.Point(3, 54);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(709, 94);
            this.panel3.TabIndex = 2;
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.panel4.Controls.Add(this.button2);
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(715, 48);
            this.panel4.TabIndex = 93;
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
            this.button2.Location = new System.Drawing.Point(674, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(38, 35);
            this.button2.TabIndex = 88;
            this.button2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // panel5
            // 
            this.panel5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(63)))), ((int)(((byte)(63)))), ((int)(((byte)(63)))));
            this.panel5.Controls.Add(this.resetBtn);
            this.panel5.Controls.Add(this.saveBtn);
            this.panel5.Location = new System.Drawing.Point(3, 470);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(709, 91);
            this.panel5.TabIndex = 3;
            // 
            // saveBtn
            // 
            this.saveBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.saveBtn.ForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.saveBtn.Location = new System.Drawing.Point(599, 27);
            this.saveBtn.Name = "saveBtn";
            this.saveBtn.Size = new System.Drawing.Size(92, 32);
            this.saveBtn.TabIndex = 3;
            this.saveBtn.Text = "Save";
            this.saveBtn.UseVisualStyleBackColor = true;
            this.saveBtn.Click += new System.EventHandler(this.saveBtn_Click);
            // 
            // resetBtn
            // 
            this.resetBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.resetBtn.ForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.resetBtn.Location = new System.Drawing.Point(501, 27);
            this.resetBtn.Name = "resetBtn";
            this.resetBtn.Size = new System.Drawing.Size(92, 32);
            this.resetBtn.TabIndex = 4;
            this.resetBtn.Text = "Reset";
            this.resetBtn.UseVisualStyleBackColor = true;
            this.resetBtn.Click += new System.EventHandler(this.resetBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.label1.Location = new System.Drawing.Point(6, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(614, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please Choose the Deired Background Colour You Would Like Below:";
            // 
            // ChangeColourForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Pink;
            this.ClientSize = new System.Drawing.Size(720, 563);
            this.Controls.Add(this.backgroundPanel);
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ChangeColourForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Change Colour Scheme";
            this.backgroundPanel.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

#pragma warning disable CS0436 // Type conflicts with imported type
        private Cyotek.Windows.Forms.ColorWheel colorWheel1;
#pragma warning restore CS0436 // Type conflicts with imported type
        private System.Windows.Forms.Panel ExamplePanel;
        private System.Windows.Forms.Panel backgroundPanel;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button resetBtn;
        private System.Windows.Forms.Button saveBtn;
        private System.Windows.Forms.Label label1;
    }
}