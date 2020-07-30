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
    public partial class changeStringFormcs : Form
    {

        public string input { get; set; }
        public string output { get; set; }
        public Color backgroundColour { get; set; }

        public changeStringFormcs( string previousString)
        {
            InitializeComponent();
            input = previousString;
            stringInput.Text = input;
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            Close();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            this.output = stringInput.Text;
            Close();
        }

        private void changeStringFormcs_Load(object sender, EventArgs e)
        {
            backgroundPanel.BackColor = backgroundColour;
        }
    }
}
