using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace automatedSystem
{
    public partial class CustomMessageBox : Form
    {
        public bool input;
        public Color backgroundColor { get; set; }
        public string ReturnValue1 { get; set; }
        public bool IfClosed { get; set; }
        
        public bool yesno { get; set; }

        public CustomMessageBox(string message, string buttonText, bool showSaveBtn, string windowName, bool enableYesNo)
        {
            InitializeComponent();
            this.BackColor = backgroundColor;
            input = showSaveBtn;
            System.Media.SystemSounds.Exclamation.Play();
            winName.Text = windowName;
            this.ReturnValue1 = "Cell 1";
            

            if(input == false && enableYesNo == false)
            {
                excelNameTxt.Visible = false;
                optionBtn0.Visible = true;
                messLbl.Visible = true;
                messLbl.Text = message;
                messageLable.Visible = false;
                optionBtn0.Text = buttonText;
                optionBtn.Visible = false;
                YesBtn.Visible = false;
                NoBtn.Visible = false;
            }
            else if(enableYesNo)
            {
                excelNameTxt.Visible = false;
                optionBtn0.Visible = false;
                messLbl.Visible = true;
                messLbl.Text = message;
                messageLable.Visible = false;
                optionBtn0.Text = buttonText;
                optionBtn.Visible = false;
                YesBtn.Visible = true;
                NoBtn.Visible = true;
            }
            else
            {
                optionBtn.Text = buttonText;
                messageLable.Text = message;
            }
            
        }

        private void CloseBtn1_Click(object sender, EventArgs e)
        {
            if (input)
            {
                if (MessageBox.Show("Are You Sure You Want To Close Without Setting The Value?", "Close Button", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    this.IfClosed = false;
                    this.Close();
                }
            }
            else
            {
                this.IfClosed = false;
                this.Close();
            }

        }

        private void CustomMessageBox_Load(object sender, EventArgs e)
        {
            this.Focus();
            
        }

        private void SaveBtn_Click(object sender, EventArgs e)
        {
            errP.Clear();
            if (MyCustomValidation.ValidLetterNumberWhitespace(excelNameTxt.Text) == false)
            { errP.SetError(excelNameTxt, "Please Enter Valid Value,\n Only Allows Letters, Numbers and Whitespace Allowed"); return; }

            else
            {
                this.ReturnValue1 = MyCustomValidation.RemoveWhiteSpace(excelNameTxt.Text);
                this.IfClosed = true;
                Close();
            }
        }

        private void OptionBtn0_Click(object sender, EventArgs e)
        {
            this.IfClosed = true;
            
            Close();
        }

        private void ExcelNameTxt_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                optionBtn.PerformClick();
            }
        }

        private void YesBtn_Click(object sender, EventArgs e)
        {
            this.yesno = true;
            this.Close();
        }

        private void NoBtn_Click(object sender, EventArgs e)
        {
            this.yesno = false;
            this.Close();

        }
    }
}
