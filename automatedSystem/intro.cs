using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace automatedSystem
{
    public partial class intro : Form
    {

        public intro()
        {
            InitializeComponent();
           
            myTimer2.Tick += new System.EventHandler(MyTimer2_Tick);
            
        }

        public Timer myTimer2 = new Timer();

        private void MyTimer2_Tick(object sender, System.EventArgs e)
        {

            myTimer2.Stop();
            myTimer2.Dispose();
            Method01();
        }

        private void Method01()
        {
          
            AutoPAM frmsecond = new AutoPAM();
            frmsecond.Show();
            this.Hide();
            frmsecond.Closed += (s, args) => this.Close();
        }

        private void Intro_Load(object sender, EventArgs e)
        {
            myTimer2.Interval = 3000;
            myTimer2.Start();
           
        }

 






     
      






    }
}
