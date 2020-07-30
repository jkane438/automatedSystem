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
    public partial class NewCatWindowcs : Form
    {
        public DataGridViewComboBoxColumn Dgv { get; set; }
        public Color backgroundColour { get; set; }
        public bool preClosed { get; set; }
        
        public NewCatWindowcs(DataGridViewComboBoxColumn dgvcb)
        {
            InitializeComponent();
            Dgv = dgvcb;
            backgroundPanel.BackColor = backgroundColour;
        }

        private void NewCatWindowcs_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < Dgv.Items.Count; i++)
            {
                liveDGV.Rows.Add();
                liveDGV.Rows[i].Cells["Catagory"].Value = Dgv.Items[i].ToString();

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            preClosed = false;
            Close();
        }

     

        private void liveDGV_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (liveDGV.Rows.Count == 11)
                liveDGV.AllowUserToAddRows = false;
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            DataGridViewComboBoxColumn tempdgv = new DataGridViewComboBoxColumn();
            for (int i = 0; i < liveDGV.Rows.Count - 1; i++)
            {
                tempdgv.Items.Add(liveDGV.Rows[i].Cells[0].Value.ToString());

            }
            Dgv = tempdgv;
            Close();

        }

        private void ResetBtn_Click(object sender, EventArgs e)
        {
            liveDGV.Rows.Clear();

            liveDGV.Rows.Add("Angle Of Head Connection (1)");
            liveDGV.Rows.Add("Angle Of Head Connection (2)");
            liveDGV.Rows.Add("Angle Of Midpiece");
        }
    }
}
