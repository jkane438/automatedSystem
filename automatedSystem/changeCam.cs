using AForge.Video.DirectShow;
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
    public partial class changeCam : Form
    {
        public Color backgroundColour { get; set; }
        public bool selected { get; set; }
        public changeCam()
        {
            InitializeComponent();
            selected = true;
            backgoundPanel.BackColor = backgroundColour;
        }

        public int selectedIndex { get; set; }
        private FilterInfoCollection colCamera; private VideoCaptureDevice CaptureDevice;
        private void changeCam_Load(object sender, EventArgs e)
        {
            colCamera = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            foreach (FilterInfo Device in colCamera)
            {
                CameraListBox.Items.Add(Device.Name);// CameraListBox.SelectedIndex = 0;
                CaptureDevice = new VideoCaptureDevice();
            }


        }

        private void CameraListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //selectedIndex = CameraListBox.SelectedIndex;
            //this.Close();
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            if (CameraListBox.SelectedItem == null)
                Close();
            else
            {
                selectedIndex = CameraListBox.SelectedIndex;
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            selected = false;
            Close();
        }
    }
}
          