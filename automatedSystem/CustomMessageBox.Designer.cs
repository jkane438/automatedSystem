namespace automatedSystem
{
    partial class CustomMessageBox
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CustomMessageBox));
            this.excelNameTxt = new System.Windows.Forms.RichTextBox();
            this.messageLable = new System.Windows.Forms.Label();
            this.optionBtn = new System.Windows.Forms.Button();
            this.errP = new System.Windows.Forms.ErrorProvider(this.components);
            this.optionBtn0 = new System.Windows.Forms.Button();
            this.messLbl = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.winName = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.NoBtn = new System.Windows.Forms.Button();
            this.YesBtn = new System.Windows.Forms.Button();
            this.CloseBtn1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.errP)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // excelNameTxt
            // 
            this.excelNameTxt.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.excelNameTxt.Location = new System.Drawing.Point(21, 92);
            this.excelNameTxt.MaximumSize = new System.Drawing.Size(607, 29);
            this.excelNameTxt.Name = "excelNameTxt";
            this.excelNameTxt.Size = new System.Drawing.Size(410, 29);
            this.excelNameTxt.TabIndex = 4;
            this.excelNameTxt.Text = "Cell";
            this.excelNameTxt.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ExcelNameTxt_KeyDown);
            // 
            // messageLable
            // 
            this.messageLable.AutoSize = true;
            this.messageLable.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.messageLable.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.messageLable.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.messageLable.Location = new System.Drawing.Point(13, 44);
            this.messageLable.Name = "messageLable";
            this.messageLable.Size = new System.Drawing.Size(236, 23);
            this.messageLable.TabIndex = 5;
            this.messageLable.Text = "Save Excel Document As:";
            // 
            // optionBtn
            // 
            this.optionBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.optionBtn.Location = new System.Drawing.Point(444, 92);
            this.optionBtn.Name = "optionBtn";
            this.optionBtn.Size = new System.Drawing.Size(101, 29);
            this.optionBtn.TabIndex = 6;
            this.optionBtn.Text = "Save";
            this.optionBtn.UseVisualStyleBackColor = true;
            this.optionBtn.Click += new System.EventHandler(this.SaveBtn_Click);
            // 
            // errP
            // 
            this.errP.ContainerControl = this;
            // 
            // optionBtn0
            // 
            this.optionBtn0.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.optionBtn0.Location = new System.Drawing.Point(218, 117);
            this.optionBtn0.Name = "optionBtn0";
            this.optionBtn0.Size = new System.Drawing.Size(101, 29);
            this.optionBtn0.TabIndex = 7;
            this.optionBtn0.Text = "Ok";
            this.optionBtn0.UseVisualStyleBackColor = true;
            this.optionBtn0.Visible = false;
            this.optionBtn0.Click += new System.EventHandler(this.OptionBtn0_Click);
            // 
            // messLbl
            // 
            this.messLbl.AutoSize = true;
            this.messLbl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.messLbl.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.messLbl.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.messLbl.Location = new System.Drawing.Point(17, 61);
            this.messLbl.Name = "messLbl";
            this.messLbl.Size = new System.Drawing.Size(20, 23);
            this.messLbl.TabIndex = 8;
            this.messLbl.Text = "  ";
            this.messLbl.Visible = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Gray;
            this.panel1.Controls.Add(this.winName);
            this.panel1.Controls.Add(this.CloseBtn1);
            this.panel1.Location = new System.Drawing.Point(4, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(549, 35);
            this.panel1.TabIndex = 9;
            // 
            // winName
            // 
            this.winName.AutoSize = true;
            this.winName.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.winName.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.winName.Location = new System.Drawing.Point(7, 5);
            this.winName.Name = "winName";
            this.winName.Size = new System.Drawing.Size(136, 23);
            this.winName.TabIndex = 10;
            this.winName.Text = "Window Name";
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.panel2.Controls.Add(this.NoBtn);
            this.panel2.Controls.Add(this.YesBtn);
            this.panel2.Location = new System.Drawing.Point(4, 41);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(549, 126);
            this.panel2.TabIndex = 10;
            // 
            // NoBtn
            // 
            this.NoBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.NoBtn.Location = new System.Drawing.Point(361, 76);
            this.NoBtn.Name = "NoBtn";
            this.NoBtn.Size = new System.Drawing.Size(101, 29);
            this.NoBtn.TabIndex = 7;
            this.NoBtn.Text = "No";
            this.NoBtn.UseVisualStyleBackColor = true;
            this.NoBtn.Visible = false;
            this.NoBtn.Click += new System.EventHandler(this.NoBtn_Click);
            // 
            // YesBtn
            // 
            this.YesBtn.Font = new System.Drawing.Font("Arial Black", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
            this.YesBtn.Location = new System.Drawing.Point(52, 76);
            this.YesBtn.Name = "YesBtn";
            this.YesBtn.Size = new System.Drawing.Size(101, 29);
            this.YesBtn.TabIndex = 7;
            this.YesBtn.Text = "Yes";
            this.YesBtn.UseVisualStyleBackColor = true;
            this.YesBtn.Visible = false;
            this.YesBtn.Click += new System.EventHandler(this.YesBtn_Click);
            // 
            // CloseBtn1
            // 
            this.CloseBtn1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.CloseBtn1.BackColor = System.Drawing.Color.Gray;
            this.CloseBtn1.BackgroundImage = global::automatedSystem.Properties.Resources.exit_button_icon_clipart_1;
            this.CloseBtn1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.CloseBtn1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CloseBtn1.FlatAppearance.BorderSize = 0;
            this.CloseBtn1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.CloseBtn1.Font = new System.Drawing.Font("Comic Sans MS", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CloseBtn1.Location = new System.Drawing.Point(509, 1);
            this.CloseBtn1.Name = "CloseBtn1";
            this.CloseBtn1.Size = new System.Drawing.Size(34, 31);
            this.CloseBtn1.TabIndex = 2;
            this.CloseBtn1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            this.CloseBtn1.UseVisualStyleBackColor = false;
            this.CloseBtn1.Click += new System.EventHandler(this.CloseBtn1_Click);
            // 
            // CustomMessageBox
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CloseBtn1;
            this.ClientSize = new System.Drawing.Size(557, 172);
            this.Controls.Add(this.messLbl);
            this.Controls.Add(this.optionBtn0);
            this.Controls.Add(this.optionBtn);
            this.Controls.Add(this.messageLable);
            this.Controls.Add(this.excelNameTxt);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "CustomMessageBox";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CustomMessageBox";
            this.Load += new System.EventHandler(this.CustomMessageBox_Load);
            ((System.ComponentModel.ISupportInitialize)(this.errP)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button CloseBtn1;
        private System.Windows.Forms.RichTextBox excelNameTxt;
        private System.Windows.Forms.Label messageLable;
        private System.Windows.Forms.Button optionBtn;
        private System.Windows.Forms.ErrorProvider errP;
        private System.Windows.Forms.Button optionBtn0;
        private System.Windows.Forms.Label messLbl;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label winName;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button NoBtn;
        private System.Windows.Forms.Button YesBtn;
    }
}