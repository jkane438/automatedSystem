using AForge.Video;
using AForge.Video.DirectShow;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace automatedSystem
{
    public partial class FullScreenForm : Form
    {
       
        public static bool enableDraw = false;
        public static bool enableAngle = false;
        public static bool enableDrawLive = false;
        public static bool enableAngleLive = true;
        double A = 0, B = 0, C = 0;
        int angleNumLive = 1, penCountLive = 0;
        List<Point> points = new List<Point>();
        List<Point> pointsA = new List<Point>();
        List<PointF> pointsF = new List<PointF>();
        List<Point> pointslive = new List<Point>();
        List<Point> pointsAlive = new List<Point>();
        //public Timer myTimer = new Timer();
        private FilterInfoCollection colCamera; private VideoCaptureDevice CaptureDevice;
        public FullScreenForm()
        {
            InitializeComponent();
           
            //47, 41

            FillInDropDownCamera();
           // myTimer.Tick += new System.EventHandler(MyTimer_Tick);

            CaptureDevice = new VideoCaptureDevice(colCamera[cameraManualLB.SelectedIndex].MonikerString);
            CaptureDevice.NewFrame += new NewFrameEventHandler(FinalFrame_NewFrame);
            CaptureDevice.Start();

            foreach (var scrn in Screen.AllScreens)
            {
                if (scrn.Bounds.Contains(this.closeBtn3.Location))
                {
                    this.closeBtn3.Location = new Point(scrn.Bounds.Right - this.closeBtn3.Width, scrn.Bounds.Top);
                    return;
                }
            }



        }
        //private void MyTimer_Tick(object sender, System.EventArgs e)
        //{
        //    myTimer.Stop();

        //    if(CaptureDevice.IsRunning == false)
        //    {
        //        CaptureDevice.Start();
        //    }


        //}
        void FinalFrame_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            Bitmap video1 = (Bitmap)eventArgs.Frame.Clone();
            Bitmap currentImage = video1;
            currentImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            FullScreenPictureBox.Image = currentImage;
        }
        void CloseCamera()
        {
            if (CaptureDevice != null)
            {
                if (CaptureDevice.IsRunning == true)
                {
                    CaptureDevice.Stop(); FullScreenPictureBox.Image = new Bitmap(640, 480);
                    FullScreenPictureBox.Image = new Bitmap(640, 480);
                    //   picVideoProcessed.Image = new Bitmap(640, 480);
                }
            }
        }
        void FillInDropDownCamera()
        {
            colCamera = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            foreach (FilterInfo Device in colCamera)
            {
                cameraManualLB.Items.Add(Device.Name); cameraManualLB.SelectedIndex = 0;
                CaptureDevice = new VideoCaptureDevice();
            }
        }

       

        private void FullScreenForm_Load(object sender, EventArgs e)
        {


            if (Dgv.Rows.Count != 0)
            {
                for (int i = 0; i < Dgv.Rows.Count; i++)
                {

                    liveDGV.Rows.Add();
                    liveDGV.Rows[i].Cells["angleCol"].Value = Dgv.Rows[i].Cells["angleCol"].Value;
                    liveDGV.Rows[i].Cells["angleInput"].Value = Dgv.Rows[i].Cells["angleInput"].Value;
                    liveDGV.Rows[i].Cells["Catagory"].Value = Dgv.Rows[i].Cells["Catagory"].Value;
                    angleNumLive++;
                    catCount++;
                }

            }

            //myTimer.Interval = 3000;

          
            //myTimer.Start();



        }
        //public void SetNewListBoxContent()
        //{
        //    //try
        //    //{

        //    //    inputCam.Items.Clear();
        //    //    FilterInfoCollection videoDevices = myCamHandler.RefreshCameraList();
        //    //    foreach (FilterInfo device in videoDevices)
        //    //    {
        //    //        inputCam.Items.Add(device.Name);
        //    //    }
        //    //    inputCam.SelectedIndex = 0; //make dafault to first cam
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    MessageBox.Show(ex.ToString(), "Error");

        //    //}
        //    try
        //    {

        //            cameraManualLB.Items.Clear();
        //            AForge.Video.DirectShow.FilterInfoCollection videoDevices = myCamHandler1.RefreshCameraList();
        //            foreach (FilterInfo device in videoDevices)
        //            {
        //                cameraManualLB.Items.Add(device.Name);
        //            }

        //            if (cameraManualLB.Items.Count != 0)
        //            {
        //                cameraManualLB.SelectedIndex = 0;
        //            }
        //            //make dafault to first cam
                
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.ToString(), "Error");

        //    }
        //}
        //public void LoadCam1()
        //{

        //    //myCamHandler1 = new CameraHandler(FullScreenPictureBox);

        //    SetNewListBoxContent();
        //    if (myCamHandler1.SetVideoSourceByIndex(cameraManualLB.SelectedIndex))
        //    {
        //        myCamHandler1.StartCapture();
        //    }
        //    myCamHandler1.StartCapture();

        //}


        //private CameraHandler myCamHandler1 = null;
        private void FullScreenForm_KeyDown(object sender, KeyEventArgs e)
        {
            //if(e.KeyCode == Keys.Escape)
            //{
            //    myCamHandler1.StopCapture();
            //    this.Close();
            //}
           
          
        }

        private void FullScreenForm_Deactivate(object sender, EventArgs e)
        {
            if (CaptureDevice.IsRunning)
                CloseCamera();

        }


        int tempCount = 0;
        private void FullScreenForm_Activated(object sender, EventArgs e)
        {
            //
            //  if(FullScreenPictureBox.Image != null)
            //myCamHandler1.StartCapture();

            if (tempCount == 0)
            {
                tempCount++;
            }
            else
            {
                if (CaptureDevice.IsRunning == false)
                    CaptureDevice.Start();
            }


        }

        private void FullScreenForm_KeyPress(object sender, KeyPressEventArgs e)
        {
        }

        private void CloseBtn3_Click(object sender, EventArgs e)
        {
            this.Dgv = liveDGV;

            if (CaptureDevice.IsRunning)
            {
                CloseCamera();
            }
            Close();

        }



        private void FullScreenPictureBox_Click(object sender, EventArgs e)
        {

        }

        public DataGridView Dgv { get; set; }

        private void FullScreenPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }

        private void CloseBtn3_Click_1(object sender, EventArgs e)
        {
                      
            

        }

        private void FullScreenForm_FormClosed(object sender, FormClosedEventArgs e)
        {

            CloseCamera();

        }

        private void FullScreenPictureBox_Paint(object sender, PaintEventArgs e)
        {
            penCountLive = pointslive.Count;
            if (enableDrawLive)
            {
                if (pointslive.Count > 1) e.Graphics.DrawLines(Pens.Red, pointslive.ToArray());

            }
            else if (enableAngleLive)
            {
                if (pointslive.Count > 1)
                {
                    e.Graphics.DrawLines(Pens.Red, pointslive.ToArray());
                }

            }
        }

        private void FullScreenPictureBox_MouseEnter(object sender, EventArgs e)
        {
         
            //MessageBox.Show("HERE");


            //if (enableDrawLive)
            //{
            //    if (File.Exists(Application.StartupPath + @"\Resources\newPencil.ico"))
            //        this.Cursor = new Cursor(Application.StartupPath + @"\Resources\newPencil.ico");
            //    else
            //    {
            //        this.Cursor = new Cursor(Application.StartupPath + @"\newPencil.ico");
            //    }
            //}
            if (enableAngleLive)
            {
                if (File.Exists(Application.StartupPath + @"\Resources\favicon.ico"))
                    this.Cursor = new Cursor(Application.StartupPath + @"\Resources\favicon.ico");
                else
                {
                    this.Cursor = new Cursor(Application.StartupPath + @"\favicon.ico");
                }


            }

        }
      
    
        
        private void FullScreenPictureBox_MouseDown(object sender, MouseEventArgs e)
        {

           int count = 1;
            MouseEventArgs me = (MouseEventArgs)e;

          


            try
            {
                if (pointslive.Count >= 3 && enableAngleLive)
                {

                    pointslive.Clear();
                    pointsAlive.Clear();
                    FullScreenPictureBox.Invalidate();
                    Show_LengthLive();
                    pointsAlive.Add(e.Location);
                    pointslive.Add(e.Location);
                    FullScreenPictureBox.Refresh();
                    return;
                }

                if (me.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    // ClearBtnLiveM.PerformClick();

                    FullScreenPictureBox.Refresh();
                    return;
                }

                else if (enableDrawLive)
                {

                    pointslive.Add(e.Location);

                    Show_LengthLive();
                    FullScreenPictureBox.Refresh();
                }
                else if (enableAngleLive)
                {

                    if (pointslive.Count == 0)
                    {
                        pointsAlive.Add(e.Location);


                    }
                    if (pointslive.Count % 2 == 0)
                    {
                        pointsAlive.Add(e.Location);

                    }

                    pointslive.Add(e.Location);
                    FullScreenPictureBox.Refresh();
                    Show_LengthLive();
                    count++;

                }
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CaptureDevice.Start();
        }

        void Show_LengthLive()
        {

            if (enableDrawLive)
            {
            }
            if (enableAngleLive)
            {

                int count;
                if (!(pointslive.Count > 1)) return;

                double len = 0;
                for (int i = 1; i < pointslive.Count; i++)
                {
                    len = Math.Sqrt((pointslive[i - 1].X - pointslive[i].X) * (pointslive[i - 1].X - pointslive[i].X)
                                + (pointslive[i - 1].Y - pointslive[i].Y) * (pointslive[i - 1].Y - pointslive[i].Y));
                }
                if (pointslive.Count <= 2)
                {
                    A = len;
                    count = 0;
                }
                else
                {
                    B = len;
                    count = 1;
                }
                if (count >= 1)
                {
                    for (int i = 1; i < pointsAlive.Count; i++)
                    {
                        len = Math.Sqrt((pointsAlive[i - 1].X - pointsAlive[i].X) * (pointsAlive[i - 1].X - pointsAlive[i].X)
                                    + (pointsAlive[i - 1].Y - pointsAlive[i].Y) * (pointsAlive[i - 1].Y - pointsAlive[i].Y));
                    }
                    C = len;

                    double x, y, z;
                    x = A * A;
                    y = B * B;
                    z = C * C;
                    float alpha = (float)Math.Acos((y + z - x) /
                                          (2 * B * C));
                    float betta = (float)Math.Acos((x + z - y) /
                                           (2 * A * C));
                    float gamma = (float)Math.Acos((x + y - z) /
                                           (2 * A * B));
                    alpha = (float)(alpha * 180 / Math.PI);
                    betta = (float)(betta * 180 / Math.PI);
                    gamma = (float)(gamma * 180 / Math.PI);
                    bool check = Double.IsNaN(gamma);


                    if (check)
                    {
                        liveDGV.Rows.Add();
                        liveDGV.Rows[angleNumLive - 1].Cells["angleCol"].Value = "Angle " + angleNumLive + " = ";
                        liveDGV.Rows[angleNumLive - 1].Cells["angleInput"].Value = "0" + "°";
                        liveDGV.FirstDisplayedScrollingRowIndex = liveDGV.RowCount - 1;

                    }
                    else
                    {
                        liveDGV.Rows.Add();
                        liveDGV.Rows[angleNumLive - 1].Cells["angleCol"].Value = "Angle " + angleNumLive + " = ";
                        liveDGV.Rows[angleNumLive - 1].Cells["angleInput"].Value = gamma.ToString() + "°";
                        liveDGV.FirstDisplayedScrollingRowIndex = liveDGV.RowCount - 1;

                    }
                    ShowCategory(0);


                    count = 0;

                    angleNumLive++;
                }


            }
           
        }
        public int catCount = 0;

        public void ShowCategory(int number)
        {
            if (catCount >= 3)
            {
                catCount = 0;
            }

            if (catCount == 0)
            {

                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = "Angle Of Head Connection (1)";

            }
            else if (catCount == 1)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = "Angle Of Head Connection (2)";


            }
            else if (catCount == 2)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = "Angle Of Midpiece";

            }
            catCount++;
        }






























    }
}
