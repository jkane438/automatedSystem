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
    public partial class ChangeColourForm : Form
    {
        public Color backGColor { get; set; }
        public ChangeColourForm(Color backgroundColour)
        {
            InitializeComponent();
            backGColor = backgroundColour;
            BackColor = backGColor;
            colorWheel1.Color = BackColor;
            backgroundPanel.BackColor = colorWheel1.Color;
            ExamplePanel.BackColor = colorWheel1.Color;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            this.backGColor = colorWheel1.Color;
            Close();
        }

        private void colorWheel1_ColorChanged(object sender, EventArgs e)
        {
            backgroundPanel.BackColor = colorWheel1.Color;
            ExamplePanel.BackColor = colorWheel1.Color;
        }

        private void resetBtn_Click(object sender, EventArgs e)
        {
            colorWheel1.Color = Color.YellowGreen;
        }
    }
}
