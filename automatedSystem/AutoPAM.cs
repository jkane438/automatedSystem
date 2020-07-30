using AForge.Video;
using AForge.Video.DirectShow;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;

using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Color = System.Drawing.Color;
using Point = System.Drawing.Point;
using Sheets = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using Workbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;
using Worksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace automatedSystem
{
    
    public partial class AutoPAM : Form
    {
        public static bool enableDraw = false;
        public static bool enableAngle = false;
        public static bool enableDrawLive = false;
        public static bool enableAngleLive = false;
        double a = 0, b = 0, c = 0, A = 0, B = 0, C = 0;
        int angleNum = 1, angleNumLive = 1, penCount = 0, penCountLive = 0;
        List<Point> points = new List<Point>();
        List<Point> pointsA = new List<Point>();
        private Image imgOriginal;
        List<PointF> pointsF = new List<PointF>();
        List<Point> pointslive = new List<Point>();
        List<Point> pointsAlive = new List<Point>();
        
       // List<PointF> pointsFlive = new List<PointF>();
        private bool Dragging;
        private int xPos;
        private int yPos;
        public int catCount = 0;
        public static string captureFile = "";
        public static int county = 0;
        public  Timer myTimer = new Timer();
        
        public AutoPAM()
        {
            InitializeComponent();
            stillImageTab1.SelectedIndex = 2;
            
            this.imagePreviewManual.MouseWheel += PictureBox1_MouseWheel;
            this.CameraLiveManualPB.MouseWheel += PictureBox1Live_MouseWheel;
            myTimer.Tick += new System.EventHandler(MyTimer_Tick);
            LoadColour();
            LoadCatagories();
            tempColor = backColourLive.BackColor;
            FillInDropDownCamera();
            if (cameraManualLB.Items.Count != 0)
            {
                if (cameraManualLB.Items.Count > 0)
                    cameraManualLB.SelectedIndex = 0;


                CaptureDevice = new VideoCaptureDevice(colCamera[cameraManualLB.SelectedIndex].MonikerString);
                CaptureDevice.NewFrame += new NewFrameEventHandler(FinalFrame_NewFrame);
                CaptureDevice.Start();
            }
            else
            { CustomMessageBox cmb = new CustomMessageBox("No Cameras Have Been Found", "Ok", false, "Camera Not Found", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();

             }

        }

        public void LoadColour()
        {

            if (File.Exists(Application.StartupPath + @"\Resources\ColourScheme\colour.txt"))
            {
                backColourLive.BackColor = Color.FromArgb(Convert.ToInt32(File.ReadAllText(Application.StartupPath + @"\Resources\ColourScheme\colour.txt")));
                BackcolourStill.BackColor = backColourLive.BackColor;
            }
            else
            {
                backColourLive.BackColor = Color.YellowGreen;
                BackcolourStill.BackColor = backColourLive.BackColor;
            }
            tempColor = backColourLive.BackColor;
        }

        public void LoadCatagories()
        {

           // Catagory.Items.Clear();

            string cat2;

            if (File.Exists(Application.StartupPath + @"\Resources\DefaultCatagories\Cat.txt"))
            {
                 cat2 = (File.ReadAllText(Application.StartupPath + @"\Resources\DefaultCatagories\Cat.txt"));

            }
            else
            {
                return;
            }
            
            string[] split = cat2.Split('\n'); 
           

            for(int i = 0; i< split.Length-1; i++)
            {
                Catagory.Items.Add(split[i].ToString());
            }
            


        }


        private Process[] processes;

     
        private void MyTimer_Tick(object sender, System.EventArgs e)
        {

            savePanel.Visible = false;


            myTimer.Stop();
        }
        private void PictureBox1_MouseWheel(object sender, MouseEventArgs e)
        {

            if (imagePreviewManual.Image != null)
            {
                if ((zoomTracker.Value + (e.Delta / 10) / 2) > 0 && (zoomTracker.Value + (e.Delta / 10) / 2) < 60)
                    zoomTracker.Value += (e.Delta / 10) / 2;

            }

        }

        private void PictureBox1Live_MouseWheel(object sender, MouseEventArgs e)
        {
          
        }

      

        private void Form1_Load(object sender, EventArgs e)
        {
           
           
        }


      

        public void CloseExcel()
        {
            
            ListViewItem itemAdd;
            ListView1.Items.Clear();
            processes = Process.GetProcessesByName("Excel");
            
            foreach (Process proc in processes)
            {
                itemAdd = ListView1.Items.Add(proc.MainWindowTitle);
                itemAdd.SubItems.Add(proc.Id.ToString());
                Process tempProc = Process.GetProcessById(proc.Id);
                tempProc.CloseMainWindow();
            }

        }

        public static Color tempColor;

        public static string CreateNewExcel(string filepath)
        {
            string bb = "";
            CustomMessageBox cmb = new CustomMessageBox("Save Excel Document As:", "Save", true, "Save", false);
            cmb.backgroundColor = tempColor;
            DialogResult dialogResult = cmb.ShowDialog();
            if (cmb.IfClosed == false)
            { return "0"; }

            bb = cmb.ReturnValue1;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            if (File.Exists(filepath + @"\" + bb + ".xlsx"))
            {
                bool valid = false;
                int tempCount = 0;
                do
                {
                    string tempName = filepath + @"\" + bb + (tempCount + 1).ToString() + ".xlsx";
                    if (File.Exists(tempName))
                    {
                        tempCount++;
                    }
                    else
                        valid = true;

                } while (!valid);
             filepath += @"\" + bb + (tempCount + 1).ToString() + ".xlsx";

            }
            else
                filepath += @"\" + bb + ".xlsx";

            //xlWorkSheet.Shapes.AddPicture(@"D:\NCBC\AutomatedSystemNCBC\automatedSystem\automatedSystem\Resources\ncbclogo.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 200, 200, 128, 128);


            county++;
            xlWorkBook.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;
            return filepath;




        }


      

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            CloseCamera();
           
            
        }

      private CameraHandler myCamHandler1;

        private void MainWindowTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (stillImageTab1.SelectedIndex)
            {
                //case 0:
                //    {
                //        if (myCamHandler != null)
                //        {
                //            myCamHandler.StopCapture();
                //        }
                //        if (myCamHandler1 != null)
                //        {
                //            myCamHandler1.StopCapture();
                //        }



                //        SetNewListBoxContent();
                //        if (myCamHandler.SetVideoSourceByIndex(inputCam.SelectedIndex))
                //        {
                //            myCamHandler.StartCapture();
                //        }

                //        previewStillLive.Image = null;

                //        stillImageTab1.SelectedIndex = 0;
                //        AutomaticliveTab.Refresh();
                //        break;
                //    }
                case 0:
                    if (CaptureDevice.IsRunning)
                    {
                        CloseCamera();
                    }
                    break;

                case 1:
                    if (CaptureDevice.IsRunning)
                    {
                        CloseCamera();
                    }
                    break;
                    
                case 2:
                    {
                        if (CaptureDevice.IsRunning == false)
                        {
                            CaptureDevice.Start();
                        }
                        break;
                    }
                case 3:
                    if (CaptureDevice.IsRunning)
                    {
                        CloseCamera();
                    }
                    break;

            }
        }

        private void ImportImageBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileOpen = new OpenFileDialog
            {
                Title = "Open Image file",
                Filter = "JPG Files (*.jpg)| *.jpg"
            };
            imagePreviewManual.Enabled = false;
            if (fileOpen.ShowDialog() == DialogResult.OK)
            {
                Bitmap bmp = new Bitmap(Image.FromFile(fileOpen.FileName), 900, 500);

                imagePreviewManual.Image = bmp;
                ImgOriginal = imagePreviewManual.Image;
                imagePreviewManual.Enabled = true;
            }
            fileOpen.Dispose();
            imagePreviewManual.Refresh();
        }
        Image Zoom(Image img, Size size)
        {

            Bitmap bmp = new Bitmap(img, img.Width + (img.Width * size.Width / 100), img.Height + (img.Height * size.Height / 100));
            Graphics g = Graphics.FromImage(bmp);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            return bmp;
        }




        void Show_LengthLive()
        {

            if (enableDrawLive)
            {

                if(LengthDGV.Rows.Count < 1)
                {
                    LengthDGV.Rows.Add();
                }
                    LengthDGV.Rows[0].Cells["lengthCol"].Value = (pointslive.Count) + " Point(s), No Segments. "; 
                    
                    if (!(pointslive.Count > 1)) return;

                    double len = 0;
                    for (int i = 1; i < pointslive.Count; i++)
                    {
                        len += Math.Sqrt((pointslive[i - 1].X - pointslive[i].X) * (pointslive[i - 1].X - pointslive[i].X)
                                    + (pointslive[i - 1].Y - pointslive[i].Y) * (pointslive[i - 1].Y - pointslive[i].Y));
                    }
                    if (tenMagnifierMLive.Checked)
                        LengthDGV.Rows[0].Cells["lengthInput"].Value = Math.Round(len / 10, 4) + "µm";
                    else if (FourtyMagnifierMLive.Checked)
                        LengthDGV.Rows[0].Cells["lengthInput"].Value = Math.Round(len / 40, 4) + "µm";
                    else if (HundredMagnifierMLive.Checked)
                        LengthDGV.Rows[0].Cells["lengthInput"].Value = Math.Round(len / 100, 4) + "µm";
                
              
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
                            liveDGV.Rows[angleNumLive-1].Cells["angleCol"].Value = "Angle " + angleNumLive + " = ";
                            liveDGV.Rows[angleNumLive-1].Cells["angleInput"].Value = "0" + "°";
                            liveDGV.FirstDisplayedScrollingRowIndex = liveDGV.RowCount - 1;

                    }
                        else
                        {
                            liveDGV.Rows.Add();
                            liveDGV.Rows[angleNumLive-1].Cells["angleCol"].Value = "Angle " + angleNumLive + " = ";
                            liveDGV.Rows[angleNumLive-1].Cells["angleInput"].Value = gamma.ToString() + "°";
                            liveDGV.FirstDisplayedScrollingRowIndex = liveDGV.RowCount - 1;

                    }
                        ShowCategory(0);
                    
                  
                    count = 0;
                   
                    angleNumLive++;
                }


            }
            ManualLiveImage.Refresh();
        }


        void Show_Length()
        {

            if (enableDraw)
            {
                if (stillLengthDGV.Rows.Count < 1)
                {
                    stillLengthDGV.Rows.Add();
                }
                stillLengthDGV.Rows[0].Cells["LengthStillCol"].Value = (points.Count) + " Point(s), No Segments. ";
             




                if (!(points.Count > 1)) return;

                double len = 0;
                for (int i = 1; i < points.Count; i++)
                {
                    len += Math.Sqrt((points[i - 1].X - points[i].X) * (points[i - 1].X - points[i].X)
                                + (points[i - 1].Y - points[i].Y) * (points[i - 1].Y - points[i].Y));
                }
                if (tenMultiStillRB.Checked)
                    stillLengthDGV.Rows[0].Cells["lengthStillInputCol"].Value = Math.Round((len / 10), 4) + "µm";
                else if (fourtyMultiStillRB.Checked)
                    stillLengthDGV.Rows[0].Cells["lengthStillInputCol"].Value = Math.Round((len / 40), 4) + "µm";
                else if (hundredMultiStillRB.Checked)
                    stillLengthDGV.Rows[0].Cells["lengthStillInputCol"].Value = Math.Round((len / 100), 4) + "µm";


            }

            if (enableAngle)
            {
              
                
                
                int count;
                if (!(points.Count > 1)) return;

                double len = 0;
                for (int i = 1; i < points.Count; i++)
                {
                    len = Math.Sqrt((points[i - 1].X - points[i].X) * (points[i - 1].X - points[i].X)
                                + (points[i - 1].Y - points[i].Y) * (points[i - 1].Y - points[i].Y));
                }
                if (points.Count <= 2)
                {
                    a = len;
                    count = 0;
                }
                else
                {
                    b = len;
                    count = 1;
                }
                if (count >= 1)
                {
                    for (int i = 1; i < pointsA.Count; i++)
                    {
                        len = Math.Sqrt((pointsA[i - 1].X - pointsA[i].X) * (pointsA[i - 1].X - pointsA[i].X)
                                    + (pointsA[i - 1].Y - pointsA[i].Y) * (pointsA[i - 1].Y - pointsA[i].Y));
                    }
                    c = len;

                    double x, y, z;
                    x = a * a;
                    y = b * b;
                    z = c * c;
                    float alpha = (float)Math.Acos((y + z - x) /
                                          (2 * b * c));
                    float betta = (float)Math.Acos((x + z - y) /
                                           (2 * a * c));
                    float gamma = (float)Math.Acos((x + y - z) /
                                           (2 * a * b));
                    alpha = (float)(alpha * 180 / Math.PI);
                    betta = (float)(betta * 180 / Math.PI);
                    gamma = (float)(gamma * 180 / Math.PI);


                    bool check = Double.IsNaN(gamma);

                    
                        if (check)
                        {

                            stillAngleDGV.Rows.Add();
                            stillAngleDGV.Rows[angleNum - 1].Cells["angleStillCol"].Value = "Angle " + angleNum + " = ";
                            stillAngleDGV.Rows[angleNum - 1].Cells["AngleInputStillCol"].Value = "0" + "°";
                            stillAngleDGV.FirstDisplayedScrollingRowIndex = stillAngleDGV.RowCount - 1;

                    }
                        else
                        {

                            stillAngleDGV.Rows.Add();
                            stillAngleDGV.Rows[angleNum - 1].Cells["angleStillCol"].Value = "Angle " + angleNum + " = ";
                            stillAngleDGV.Rows[angleNum - 1].Cells["AngleInputStillCol"].Value = gamma.ToString() + "°";
                            stillAngleDGV.FirstDisplayedScrollingRowIndex = stillAngleDGV.RowCount - 1;

                    }
                    
                    ShowCategoryStill(0);
                    
                    count = 0;
                    angleNum++;
                }


            }
            imagePreviewManual.Refresh();
        }



        private void ImagePreview_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {

                int count = 1;
                MouseEventArgs me = (MouseEventArgs)e;
                if (enableAngle == false && enableDraw == false)
                {
                    if (File.Exists(Application.StartupPath + @"\Resources\unnamed.ico"))
                        this.Cursor = new Cursor(Application.StartupPath + @"\Resources\unnamed.ico");
                    else
                        this.Cursor = new Cursor(Application.StartupPath + @"\unnamed.ico");
                    

                    if (e.Button == MouseButtons.Left)
                    {
                        Dragging = true;
                        xPos = e.X;
                        yPos = e.Y;
                    }
                }

                if (points.Count >= 3 && enableAngle)
                {

                    points.Clear();
                    pointsA.Clear();
                    imagePreviewManual.Invalidate();
                    Show_Length();
                    pointsA.Add(e.Location);
                    points.Add(e.Location);
                    imagePreviewManual.Refresh();
                    return;
                }
  

                if (me.Button == MouseButtons.Right)
                {
                   // ClearBtn.PerformClick();
                    imagePreviewManual.Refresh();
                    return;
                }

                else if (enableDraw)
                {

                    points.Add(e.Location);

                    Show_Length();
                    imagePreviewManual.Refresh();
                }
                else if (enableAngle)
                {

                    if (points.Count == 0)
                    {
                        pointsA.Add(e.Location);


                    }
                    if (points.Count % 2 == 0)
                    {
                        pointsA.Add(e.Location);

                    }

                    points.Add(e.Location);
                    imagePreviewManual.Refresh();
                    Show_Length();
                    count++;

                }

                
                imagePreviewManual.Refresh();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }


        }


        public Image ImgOriginal { get => imgOriginal; set => imgOriginal = value; }


        private void ImagePreview_Paint(object sender, PaintEventArgs e)
        {
            try
            {

                penCount = points.Count;
                if (enableDraw)
                {
                    if (points.Count > 1) e.Graphics.DrawLines(Pens.Red, points.ToArray());

                    
                }
                else if (enableAngle)
                {
                    if (points.Count > 1)
                    {
                        e.Graphics.DrawLines(Pens.Red, points.ToArray());

                      
                    }

                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }

        }

        private void ClearBtn_Clcik(object sender, EventArgs e)
        {

            points.Clear();
            pointsA.Clear();
            imagePreviewManual.Invalidate();


            stillAngleDGV.Rows.Clear();
            stillLengthDGV.Rows.Clear();
            angleNum = 1;

        }



     

        private void ZoomTracker_Scroll(object sender, EventArgs e)
        {
            if (zoomTracker.Value > 0)
            {
                if (imagePreviewManual.Image != null)
                {
                    imagePreviewManual.Image = Zoom(ImgOriginal, new Size(zoomTracker.Value, zoomTracker.Value));
                }
            }
        }

        private void ZoomTracker_ValueChanged(object sender, EventArgs e)
        {
            if (zoomTracker.Value > 0)
            {
                if (imagePreviewManual.Image != null)
                {
                    imagePreviewManual.Image = Zoom(ImgOriginal, new Size(zoomTracker.Value, zoomTracker.Value));
                }
            }
        }


        private void CameraLiveManualPB_MouseDown(object sender, MouseEventArgs e)
        {
            int count = 1;
            MouseEventArgs me = (MouseEventArgs)e;

            try
            {
                if (pointslive.Count >= 3 && enableAngleLive)
                {

                    pointslive.Clear();
                    pointsAlive.Clear();
                    CameraLiveManualPB.Invalidate();
                    Show_LengthLive();
                    pointsAlive.Add(e.Location);
                    pointslive.Add(e.Location);
                    CameraLiveManualPB.Refresh();
                    return;
                }

                if (me.Button == System.Windows.Forms.MouseButtons.Right)
                {
                   // ClearBtnLiveM.PerformClick();

                    CameraLiveManualPB.Refresh();
                    return;
                }

                else if (enableDrawLive)
                {

                    pointslive.Add(e.Location);

                    Show_LengthLive();
                    CameraLiveManualPB.Refresh();
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
                    CameraLiveManualPB.Refresh();
                    Show_LengthLive();
                    count++;

                }
                ManualLiveImage.Refresh();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }


        private void CameraLiveManualPB_Paint(object sender, PaintEventArgs e)
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


        private void ClearBtnLiveM_Click(object sender, EventArgs e)
        {

            //pointslive.Clear();
            //pointsAlive.Clear();
            //CameraLiveManualPB.Invalidate();
            //Show_LengthLive();
            //catCount = 0;

            liveDGV.Rows.Clear();
            angleNumLive = 1;
            LengthDGV.Rows.Clear();

            pointslive.Clear();
            CameraLiveManualPB.Invalidate();
            catCount = 0;
            ManualLiveImage.Refresh();
            //LineBtnLiveM.PerformClick();

            //for(int i = 0; i< liveDGV.Rows.Count*2; i++)
            //  {
            //      liveDGV.Rows.RemoveAt(i);
            //  }

            //  angleNumLive = 1;


        }


        private void CameraLiveManualPB_EnabledChanged(object sender, EventArgs e)
        {

        }


        

        private void ImagePreviewManual_MouseUp(object sender, MouseEventArgs e)
        {
            Dragging = false;
            if (enableAngle == false && enableDraw == false)
            {
                if (File.Exists(Application.StartupPath + @"\Resources\openCursor.ico"))
                    this.Cursor = new Cursor(Application.StartupPath + @"\Resources\openCursor.ico");
                else
                    this.Cursor = new Cursor(Application.StartupPath + @"\openCursor.ico");
            }
        }

     

        private void ManualTab_MouseHover(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

      

        private void ManualTab_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void ImagePreviewManual_MouseEnter(object sender, EventArgs e)
        {
            if (enableAngle == false && enableDraw == false)
            {
                if (File.Exists(Application.StartupPath + @"\Resources\openCursor.ico"))
                    this.Cursor = new Cursor(Application.StartupPath + @"\Resources\openCursor.ico");
                else
                    this.Cursor = new Cursor(Application.StartupPath + @"\openCursor.ico");
            }
            if (enableDraw)
            {

                if (File.Exists(Application.StartupPath + @"\Resources\newPencil.ico"))
                    this.Cursor = new Cursor(Application.StartupPath + @"\Resources\newPencil.ico");
                else
                    this.Cursor = new Cursor(Application.StartupPath + @"\newPencil.ico");
            }
            if (enableAngle)
            {

                if (File.Exists(Application.StartupPath + @"\Resources\favicon.ico"))
                    this.Cursor = new Cursor(Application.StartupPath + @"\Resources\favicon.ico");
                else
                    this.Cursor = new Cursor(Application.StartupPath + @"\favicon.ico");
            }
        }

        private void CaptureBtnLive_Click(object sender, EventArgs e)
        {
            //previewStillLive.Image = myCamHandler.GetSnapshot();
        }



        private void GrabBtn_Click(object sender, EventArgs e)
        {
            enableAngle = false;
            enableDraw = false;
            GrabBtn.BackColor = Color.LightBlue;
            angleBtn.BackColor = Color.Gray;
            LineBtnM.BackColor = Color.Gray;
            imagePreviewManual.Enabled = true;
        }


        private void Button1_Click_1(object sender, EventArgs e)
        {
            if (liveDGV.Rows.Count != 0)
            {
                
                liveDGV.Rows.RemoveAt(liveDGV.Rows.Count - 1);
                angleNumLive--;
                if (catCount == 0)
                    catCount = 3;
                else
                    catCount--;
            }

            pointsA.Clear();
            points.Clear();
            pointsF.Clear();
            pointslive.Clear();
            CameraLiveManualPB.Invalidate();
            return;

        }


        private void AngleUndoBtn_Click(object sender, EventArgs e)
        {

            if (stillAngleDGV.Rows.Count != 0)
            {

                stillAngleDGV.Rows.RemoveAt(stillAngleDGV.Rows.Count - 1);
                angleNum--;
              
                if (catcount == 0)
                    catcount = 3;
                else
                    catcount--;
            }

            pointsA.Clear();
            points.Clear();
            pointsF.Clear();
            pointslive.Clear();
            imagePreviewManual.Invalidate();
            return;




        }

        private void CameraLiveManualPB_MouseEnter(object sender, EventArgs e)
        {
            if (enableDrawLive)
            {
                if(File.Exists(Application.StartupPath + @"\Resources\newPencil.ico"))
                this.Cursor = new Cursor(Application.StartupPath + @"\Resources\newPencil.ico");
                else
                {
                    this.Cursor = new Cursor(Application.StartupPath + @"\newPencil.ico");
                }
            }
            if(enableAngleLive)
            {
                if (File.Exists(Application.StartupPath + @"\Resources\favicon.ico"))
                    this.Cursor = new Cursor(Application.StartupPath + @"\Resources\favicon.ico");
                else
                {
                    this.Cursor = new Cursor(Application.StartupPath + @"\favicon.ico");
                }
               

            }

        }

        private void CameraLiveManualPB_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

        private void ImagePreviewManual_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Arrow;
        }

       

        private void ImagePreviewManual_MouseMove(object sender, MouseEventArgs e)
        {
            System.Windows.Forms.Control c = sender as System.Windows.Forms.Control;
            if (Dragging && c != null)
            {
                int maxX = imagePreviewManual.Size.Width * -1 + ZoomMovePanel.Size.Width;
                int maxY = imagePreviewManual.Size.Height * -1 + ZoomMovePanel.Size.Height;

                int newposLeft = e.X + c.Left - xPos;
                int newposTop = e.Y + c.Top - yPos;

                if (newposTop > 0)
                {
                    newposTop = 0;
                }
                if (newposLeft > 0)
                {
                    newposLeft = 0;
                }
                if (newposLeft < maxX)
                {
                    newposLeft = maxX;
                }
                if (newposTop < maxY)
                {
                    newposTop = maxY;
                }
                c.Top = newposTop;
                c.Left = newposLeft;

            }
        }
        public static Image img;
        private void CalculateBtnLiveM_Click(object sender, EventArgs e)
        {
            img = CameraLiveManualPB.Image;
            this.Cursor = Cursors.WaitCursor;
            //ReadXml();
            CloseExcel();

            int num = 1;
            
            if(liveDGV.Rows.Count !=0)
            {
                num++;
            }
            

            if (num == 1)
            {
                CustomMessageBox cmb1 = new CustomMessageBox("Please Use The Angle Tool To Calculate Data", "Ok", false, "Angle Tools", false);
                cmb1.backgroundColor = tempColor;
                DialogResult dialogResult1 = cmb1.ShowDialog();
              
                this.Cursor = Cursors.Default;
                return;
            }

            string filePath = "", startupPath = " ";
            //Find Last Saved Filepath
            if (File.Exists(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
            {
                startupPath = File.ReadAllText(Application.StartupPath + @"\Resources\lastSave\filepath.txt");

               startupPath = startupPath.Remove(startupPath.Length - 2);
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    
                    fbd.SelectedPath = startupPath;
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        filePath = fbd.SelectedPath.ToString();
                    }
                    else
                    { this.Cursor = Cursors.Default; return; }
                }

                 using (System.IO.StreamWriter file = new StreamWriter(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
                {
                    file.WriteLine(filePath);
                }
                 

            }
            else //Create New Last Saved Filepath

            {

                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.SelectedPath = startupPath;
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        filePath = fbd.SelectedPath.ToString();
                    }
                    else
                    { this.Cursor = Cursors.Default; return; }
                }
                if (File.Exists(Application.StartupPath + @"\Resources\lastSave"))
                { }
                else
                {
                    Directory.CreateDirectory(Application.StartupPath + @"\Resources\lastSave");

                }
                using (System.IO.StreamWriter file = new StreamWriter(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
                {
                    file.WriteLine(filePath);
                }

            }


            string filename = CreateNewExcel(filePath);

            if (filename == "0")
            { this.Cursor = Cursors.Default; return; }

        
            dgvAngle = liveDGV;
            dgvLength = LengthDGV;

            CreateSpreadsheetWorkbook(filename);
            InsertWorksheet(filename);
            

            Image image = CameraLiveManualPB.Image;

            InsertImageToNewSheet(filename, image);
           

            this.Cursor = Cursors.Default;

            savedLable.Text = "Excel Saved";
            myTimer.Interval = 3000;

            savePanel.Visible = true;
            myTimer.Start();
            //CustomMessageBox cmb = new CustomMessageBox("Excel Document Created.", "Ok", false, "Excel");
            //    var dialogResult = cmb.ShowDialog();

        }


        public static DataGridView dgvAngle;
        public static DataGridView dgvLength;
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            Rectangle destRect = new Rectangle(0, 0, width, height);
            Bitmap destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (Graphics graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (ImageAttributes wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        private void LineBtnLiveM_Click(object sender, EventArgs e)
        {
            if (tenMagnifierMLive.Checked == false && FourtyMagnifierMLive.Checked == false && HundredMagnifierMLive.Checked == false)
            {
                CustomMessageBox cmb = new CustomMessageBox("Please Select A Magnification of the Microscope", "Ok", false, "Length", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();
               
                return;
            }

            if (enableDrawLive)
            {
                LineBtnLiveM.BackColor = Color.Gray;
                enableDrawLive = false;
                CameraLiveManualPB.Enabled = false;
                pointslive.Clear();
                CameraLiveManualPB.Invalidate();
              
                enableAngleLive = false;
                AngleBtnLiveM.BackColor = Color.Gray;

            }
            else if (enableDrawLive == false)
            {

                LineBtnLiveM.BackColor = Color.LightBlue;
                enableDrawLive = true; CameraLiveManualPB.Enabled = true;
                pointslive.Clear();
                CameraLiveManualPB.Invalidate();
       
                enableAngleLive = false;
                AngleBtnLiveM.BackColor = Color.Gray;
            }
            ManualLiveImage.Refresh();
        }

        private void AngleBtnLiveM_Click(object sender, EventArgs e)
        {
            if (enableAngleLive)
            {
               
                AngleBtnLiveM.BackColor = Color.Gray;
                CameraLiveManualPB.Enabled = false;
                enableAngleLive = false;
                enableDrawLive = false;
                LineBtnLiveM.BackColor = Color.Gray;

            }
            else if (enableAngleLive == false)
            {
                AngleBtnLiveM.BackColor = Color.LightBlue;
               // AngleBtnLiveM.BackColor = Color.Gray;
                CameraLiveManualPB.Invalidate();
                
                CameraLiveManualPB.Enabled = true;
                enableAngleLive = true;
                enableDrawLive = false;
                LineBtnLiveM.BackColor = Color.Gray;

            }
            ManualLiveImage.Refresh();
        }

        private void ImagePreviewManual_EnabledChanged(object sender, EventArgs e)
        {
            //resets the zoom trackbar to 0 when image is imported 
            zoomTracker.Value = zoomTracker.Minimum;

        }

        private void StartBtn_Click(object sender, EventArgs e)
        {

            if (CaptureDevice.IsRunning)
            {
                CloseCamera();
            }
            else
            {
                CaptureDevice.Start();
            }
        }

        private void Button28_Click(object sender, EventArgs e)
        {
            ClosePress();
        }



        private void AutoLiveBtn4_Click(object sender, EventArgs e)
        {
            Nav(1);
        }

        private void AutoStillLiveBtn4_Click(object sender, EventArgs e)
        {
            Nav(2);
        }

        private void ManualLiveImageBtn4_Click(object sender, EventArgs e)
        {
            Nav(3);
        }

        private void ManualStillImageBtn4_Click(object sender, EventArgs e)
        {
            Nav(4);
        }

        public void Nav(int num)
        {
            switch (num)
            {
                case 1: stillImageTab1.SelectedIndex = 0; break;
                case 2: stillImageTab1.SelectedIndex = 1; break;
                case 3: stillImageTab1.SelectedIndex = 2; break;
                case 4: stillImageTab1.SelectedIndex = 3; break;
            }
        }

        private void AutoLiveBtn3_Click(object sender, EventArgs e)
        {
            Nav(1);
        }

        private void AutoStillLiveBtn3_Click(object sender, EventArgs e)
        {
            Nav(2);
        }

        private void ManualLiveImageBtn3_Click(object sender, EventArgs e)
        {
            Nav(3);
        }

        private void ManualStillImageBtn3_Click(object sender, EventArgs e)
        {
            Nav(4);
        }

        private void AutoLiveBtn2_Click(object sender, EventArgs e)
        {
            Nav(1);
        }

        private void AutoStillLiveBtn2_Click(object sender, EventArgs e)
        {
            Nav(2);
        }

        private void ManualLiveImageBtn2_Click(object sender, EventArgs e)
        {
            Nav(3);
        }

        private void ManualStillImageBtn2_Click(object sender, EventArgs e)
        {
            Nav(4);
        }

        private void AutoLiveBtn1_Click(object sender, EventArgs e)
        {
            Nav(1);
        }

        private void AutoStillLiveBtn1_Click(object sender, EventArgs e)
        {
            Nav(2);
        }

        private void ManualLiveImageBtn1_Click(object sender, EventArgs e)
        {
            Nav(3);
        }

        private void ManualStillImageBtn1_Click(object sender, EventArgs e)
        {
            Nav(4);
        }
        public void ClosePress()
        {
            if (CaptureDevice.IsRunning)
            {
                CloseCamera();
            }
            
            Close();
        }

        private void CloseBtn2_Click(object sender, EventArgs e)
        {
            ClosePress();
        }

        private void CloseBtn3_Click(object sender, EventArgs e)
        {
            //myCamHandler1.StopCapture();
            
            ClosePress();
        }

        private void CloseBtn4_Click(object sender, EventArgs e)
        {
            ClosePress();
        }

        private void CloseBtn1_Click(object sender, EventArgs e)
        {
            ClosePress();
        }

        private void Button1_Click_2(object sender, EventArgs e)
        {
          //  myCamHandler.StopCapture();

           // myCamHandler1.StopCapture();
            this.WindowState = FormWindowState.Minimized;
        }

        private void Minimise4_Click(object sender, EventArgs e)
        {
           // myCamHandler.StopCapture();
           // myCamHandler1.StopCapture();
            this.WindowState = FormWindowState.Minimized;
        }

        private void Minimise1_Click(object sender, EventArgs e)
        {
            
          //  myCamHandler.StopCapture();
          //  myCamHandler1.StopCapture();
            this.WindowState = FormWindowState.Minimized;
        }

        private void Minimise2_Click(object sender, EventArgs e)
        {
          ////  myCamHandler.StopCapture();
          //  myCamHandler1.StopCapture();
            this.WindowState = FormWindowState.Minimized;
        }
        
        private void Label4_Click(object sender, EventArgs e)
        {
            myCamHandler1 = new CameraHandler();
            MessageBox.Show(myCamHandler1.Method01());
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            if (CaptureDevice.IsRunning)
            {
                CaptureDevice.Stop();
            }
            else
                CaptureDevice.Start(); 
            // myCamHandler.StopCapture();
          //  if (myCamHandler1.IsRunning())
          //  {
          //      myCamHandler1.StopCapture();
          //      return;
          //  }
          //  //myCamHandler1 = null;
          ////  myCamHandler1 = new CameraHandler(CameraLiveManualPB);
          //  try
          //  {
          //      //  myCamHandler = new CameraHandler(pbcameraOut);

          //      SetNewListBoxContent();
          //      LoadCam1();
          //  }
          //  catch (Exception ex)
          //  {
          //      MessageBox.Show(ex.ToString());
          //  }
        }
        private FilterInfoCollection colCamera; public VideoCaptureDevice CaptureDevice;




        void FinalFrame_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {

            Bitmap video1 = (Bitmap)eventArgs.Frame.Clone();
            Bitmap currentImage = video1;
   
            currentImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
            currentImage.RotateFlip(RotateFlipType.Rotate180FlipXY);

           

            CameraLiveManualPB.Image = currentImage;
        }
        void CloseCamera()
        {
            if (cameraManualLB.Items.Count != 0)
            {
                if (CaptureDevice != null)
                {
                    if (CaptureDevice.IsRunning == true)
                    {
                        CaptureDevice.Stop(); CameraLiveManualPB.Image = new Bitmap(640, 480);
                        CameraLiveManualPB.Image = new Bitmap(640, 480);
                        //   picVideoProcessed.Image = new Bitmap(640, 480);
                    }
                }
            }
        }

        void FillInDropDownCamera()
        {
           
                colCamera = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                foreach (FilterInfo Device in colCamera)
                {
                    cameraManualLB.Items.Add(Device.Name); 
                    CaptureDevice = new VideoCaptureDevice();
                }

           
        }
        public int tempNumber = 0;
        public void ShowCategory(int number)
        {
            if (catCount >= Catagory.Items.Count)
            {
                catCount = 0;
            }
            
            if (catCount == 0)
            {

                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[0];

            }
            else if (catCount == 1)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[1];


            }
            else if (catCount == 2)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[2];

            }
            if (catCount == 3)
            {

                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[3];

            }
            else if (catCount == 4)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[4];


            }
            else if (catCount == 5)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[5];

            }
            if (catCount == 6)
            {

                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[6];

            }
            else if (catCount == 7)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[7];


            }
            else if (catCount == 8)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[8];

            }
            if (catCount == 9)
            {

                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[9];

            }
            else if (catCount == 10)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[10];


            }
            else if (catCount == 11)
            {
                liveDGV.Rows[angleNumLive - 1].Cells["Catagory"].Value = Catagory.Items[11];

            }
            catCount ++;
        }

        public void ShowCategoryStill(int number)
        {

            if (catcount >= 3)
            {
                catcount = 0;
            }

            if (catcount == 0)
            {
                stillAngleDGV.Rows[angleNum - 1].Cells["CatStillCol"].Value = "Angle Of Head Connection (1)";
            }
            else if (catcount == 1)
            {
                stillAngleDGV.Rows[angleNum - 1].Cells["CatStillCol"].Value = "Angle Of Head Connection (2)";
            }
            else if (catcount == 2)
            {
                stillAngleDGV.Rows[angleNum - 1].Cells["CatStillCol"].Value = "Angle Of Midpiece";
            }
            catcount++;


        }

        public int catcount = 0;

  


        private void ExportStillManualBtn_Click(object sender, EventArgs e)
        {

            this.Cursor = Cursors.WaitCursor;
            //ReadXml();
            CloseExcel();

            int num = stillAngleDGV.Rows.Count;
            

            if (num == 1)
            {
                CustomMessageBox cmb = new CustomMessageBox("Please Use The Angle Tool To Calculate Data", "Ok", false, "Angle", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();
                
                this.Cursor = Cursors.Default;
                return;
            }
            else if (imagePreviewManual.Image == null)
            {
                CustomMessageBox cmb = new CustomMessageBox("Please Import an Image", "Ok", false, "Angle", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();

                this.Cursor = Cursors.Default;
                return;

            }

            string filePath = "", startupPath = " ";
            //Last saved filepath
            if (File.Exists(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
            {
                startupPath = File.ReadAllText(Application.StartupPath + @"\Resources\lastSave\filepath.txt");

                startupPath = startupPath.Remove(startupPath.Length - 2);
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {

                    fbd.SelectedPath = startupPath;
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        filePath = fbd.SelectedPath.ToString();
                    }
                    else
                    { this.Cursor = Cursors.Default; return; }
                }

                using (System.IO.StreamWriter file = new StreamWriter(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
                {
                    file.WriteLine(filePath);
                }


            }
            else
            {

                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.SelectedPath = startupPath;
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        filePath = fbd.SelectedPath.ToString();
                    }
                    else
                    { this.Cursor = Cursors.Default; return; }
                }
                if (File.Exists(Application.StartupPath + @"\Resources\lastSave"))
                { }
                else
                {
                    Directory.CreateDirectory(Application.StartupPath + @"\Resources\lastSave");

                }
                using (System.IO.StreamWriter file = new StreamWriter(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
                {
                    file.WriteLine(filePath);
                }

            }

            string filename = CreateNewExcel(filePath);

            if (filename == "0")
                return;

            dgvAngle = stillAngleDGV;
            dgvLength = stillLengthDGV;

            CreateSpreadsheetWorkbook(filename);
            InsertWorksheetStill(filename);
            InsertImageToNewSheet(filename, imagePreviewManual.Image);
            this.Cursor = Cursors.Default;
            CustomMessageBox cmb1 = new CustomMessageBox("Excel Document Created.", "Ok", false, "Excel", false);
            cmb1.backgroundColor = tempColor;
            DialogResult dialogResult1 = cmb1.ShowDialog();
        

        }

      
        public static Image tempImg;
        private void ExportSheetManualLiveBtn_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //ReadXml();
            CloseExcel();

           
            if (liveDGV.Rows.Count < 1)
            {

            
            CustomMessageBox cmb = new CustomMessageBox("Please Use The Angle Tool To Calculate Data.", "Ok", false, "Angle", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();
            
                this.Cursor = Cursors.Default;
                return;
            }
            string filePath = "";

            OpenFileDialog open = new OpenFileDialog
            {
                // Document filters  
                Filter = "Excel Files(*.xlsx; *.xlsm; *.xltx)|*.xlsx; *.xlsm; *.xltx"
            };
            if (open.ShowDialog() == DialogResult.OK)
            {
                filePath = open.FileName;
            }

            dgvAngle = liveDGV;
            dgvLength = LengthDGV;


            tempImg = CameraLiveManualPB.Image;

          

            if (InsertText(filePath))
            {
                InsertImageToNewSheet(filePath, tempImg);
                savedLable.Text = "Excel Added";
                myTimer.Interval = 2000;

                savePanel.Visible = true;
                myTimer.Start();
                //CustomMessageBox cmb1 = new CustomMessageBox("Excel Document Created.", "Ok", false, "Excel");
                //var dialogResult1 = cmb1.ShowDialog();
                this.Cursor = Cursors.Default;

            }
            else
            {
                CustomMessageBox cmb1 = new CustomMessageBox("Documment Not Created.", "Ok", false, "Excel", false);
                cmb1.backgroundColor = tempColor;
                DialogResult dialogResult1 = cmb1.ShowDialog();
                this.Cursor = Cursors.Default;
            }

            



        }

        private void Button1_Click_4(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //ReadXml();
            CloseExcel();

            int num = stillAngleDGV.Rows.Count;

           

            if (num == 1)
            {
                CustomMessageBox cmb = new CustomMessageBox("Please Use The Angle Tool To Calculate Data.", "Ok", false, "Angle", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();

                this.Cursor = Cursors.Default;
                return;
            }
            else if (imagePreviewManual.Image == null)
            {
                CustomMessageBox cmb = new CustomMessageBox("Please Import an Image", "Ok", false, "Angle", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();

                this.Cursor = Cursors.Default;
                return;
            

            }

            string filePath = "";

            OpenFileDialog open = new OpenFileDialog
            {
                // image filters  
                Filter = "Excel Files(*.xlsx; *.xlsm; *.xltx)|*.xlsx; *.xlsm; *.xltx"
            };
            if (open.ShowDialog() == DialogResult.OK)
            {
                filePath = open.FileName;
            }
            dgvAngle = stillAngleDGV;
            dgvLength = stillLengthDGV;


            if (InsertTextStill(filePath))
            {
                InsertImageToNewSheet(filePath, imagePreviewManual.Image);
                CustomMessageBox cmb1 = new CustomMessageBox("Excel Document Created.", "Ok", false, "Excel", false)
                {
                    backgroundColor = tempColor
                };
                DialogResult dialogResult1 = cmb1.ShowDialog();
                this.Cursor = Cursors.Default;
            }
            else
            {
                CustomMessageBox cmb1 = new CustomMessageBox("Documment Not Created.", "Ok", false, "Excel", false);
                cmb1.backgroundColor = tempColor;
                DialogResult dialogResult1 = cmb1.ShowDialog();
                this.Cursor = Cursors.Default;
            }
           

        }

        private void LineBtnM_Click(object sender, EventArgs e)
        {

            if (tenMultiStillRB.Checked == false && fourtyMultiStillRB.Checked == false && hundredMultiStillRB.Checked == false)
            {
                CustomMessageBox cmb = new CustomMessageBox("Please Select A Magnification of the Microscope.", "Ok", false, "Length", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();

                
                return;
            }

            if (enableDraw)
            {
                LineBtnM.BackColor = Color.Gray;
                enableDraw = false;
                imagePreviewManual.Enabled = false;
                points.Clear();
                imagePreviewManual.Invalidate();
               
                enableAngle = false;
                angleBtn.BackColor = Color.Gray;
                GrabBtn.BackColor = Color.Gray;

            }
            else if (enableDraw == false)
            {
                LineBtnM.BackColor = Color.LightBlue;
                enableDraw = true; imagePreviewManual.Enabled = true;
                points.Clear();
                imagePreviewManual.Invalidate();
               
                enableAngle = false;
                angleBtn.BackColor = Color.Gray;
                GrabBtn.BackColor = Color.Gray;
            }
            imagePreviewManual.Refresh();
        }

        private void AngleBtn_Click(object sender, EventArgs e)
        {
            if (enableAngle)
            {
                angleBtn.BackColor = Color.Gray;
                imagePreviewManual.Enabled = false;
                enableAngle = false;
                enableDraw = false;
                LineBtnM.BackColor = Color.Gray;
                GrabBtn.BackColor = Color.Gray;

            }
            else if (enableAngle == false)
            {
                angleBtn.BackColor = Color.LightBlue;
                points.Clear();
                pointsA.Clear();
                imagePreviewManual.Invalidate();
               
                imagePreviewManual.Enabled = true;
                enableAngle = true;
                enableDraw = false;
                LineBtnM.BackColor = Color.Gray;
                GrabBtn.BackColor = Color.Gray;


            }
            imagePreviewManual.Refresh();
        }
        public static int imageDGVCount = 0, tempCounter = 0, fileNum = 0;
        public string captureString = "0001";
        private void CaptureManualLive_Click(object sender, EventArgs e)
        {
            //ImagePreviewLbl.Visible = true;
            //previewManualLivePB.Visible = true;
            //previewManualLivePB.Image = myCamHandler1.GetSnapshot();
            if(imageDGVCount == 0)
            {
                CaptureDGV.Rows.Add();
                CaptureDGV.Rows[imageDGVCount].Cells[0].Value = captureString;
                CaptureDGV.Rows[imageDGVCount].Cells[1].Value = CameraLiveManualPB.Image;

                CaptureDGV.Rows[imageDGVCount].Height = 300;
                CaptureDGV.FirstDisplayedScrollingRowIndex = CaptureDGV.RowCount - 1;
                imageDGVCount++;
            }
            else
            {
                CaptureDGV.Rows.Add();
                CaptureDGV.Rows[imageDGVCount].Cells[0].Value = captureString + "("+ (tempCounter+1).ToString() +")";
                CaptureDGV.Rows[imageDGVCount].Cells[1].Value = CameraLiveManualPB.Image;
                tempCounter++;
                CaptureDGV.Rows[imageDGVCount].Height = 300;
                CaptureDGV.FirstDisplayedScrollingRowIndex = CaptureDGV.RowCount - 1;
                imageDGVCount++;
            }
                

            
           
                }
        
        private void SaveCaptureBtn_Click(object sender, EventArgs e)
        {
            
            if (CaptureDGV.Rows.Count == 0)
            {
                CustomMessageBox cmb = new CustomMessageBox("Please Capture an Image", "Ok", false, "Capture", false);
                cmb.backgroundColor = tempColor;
                DialogResult dialogResult = cmb.ShowDialog();
                return;
            }

            FolderBrowserDialog folderDialog = new FolderBrowserDialog
            {
                Description = "Select Path to Save Images",
                ShowNewFolderButton = true
            };

            string filePath = "", startupPath = " ";

            //Find Last Saved Filepath
            if (File.Exists(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
            {
                startupPath = File.ReadAllText(Application.StartupPath + @"\Resources\lastSave\filepath.txt");

                startupPath = startupPath.Remove(startupPath.Length - 2);
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {

                    fbd.SelectedPath = startupPath;
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        filePath = fbd.SelectedPath.ToString();
                    }
                    else
                    { this.Cursor = Cursors.Default; return; }
                }

                using (System.IO.StreamWriter file1 = new StreamWriter(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
                {
                    file1.WriteLine(filePath);
                }

            }
            else //Create New Last Saved Filepath
            {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.SelectedPath = startupPath;
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        filePath = fbd.SelectedPath.ToString();
                    }
                    else
                    { this.Cursor = Cursors.Default; return; }
                }
                if (File.Exists(Application.StartupPath + @"\Resources\lastSave"))
                { }
                else
                {
                    Directory.CreateDirectory(Application.StartupPath + @"\Resources\lastSave");

                }
                using (System.IO.StreamWriter file1 = new StreamWriter(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
                {
                    file1.WriteLine(filePath);
                }

            }
            
            //if (folderDialog.ShowDialog() == DialogResult.Abort || folderDialog.ShowDialog() == DialogResult.Cancel)
            //{
            //    return;
            //}

            string file = filePath;

            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = "JPeg Image|*.jpg|Excel|.xlsx",
                Title = "Save Image File"
            };

            saveFileDialog1.FileName = "Image";
           

            //if (saveFileDialog1.ShowDialog()==DialogResult.Cancel)
            //{
            //    return;
            //}

            string temp = file + @"\";
           // temp = temp.Remove(temp.Length - 4);
      

            for (int i = 0; i < CaptureDGV.Rows.Count; i++)
            {
                if (File.Exists(temp + CaptureDGV.Rows[i].Cells[0].Value + ".jpg"))
                {
                    bool valid = false;
                    int fileNum = 0;
                    do
                    {
                        string tempName = temp + CaptureDGV.Rows[i].Cells[0].Value + (" " + fileNum + 1).ToString() + ".jpg";
                        if (File.Exists(tempName))
                        {
                            fileNum++;
                        }
                        else
                        {
                            valid = true;
                            saveFileDialog1.FileName = tempName;
                        }

                    } while (!valid);
                }
                else
                    saveFileDialog1.FileName = temp + CaptureDGV.Rows[i].Cells[0].Value + ".jpg";

              
                if (File.Exists(temp + CaptureDGV.Rows[i].Cells[0].Value + ".jpg"))
                {
                   
                    FileStream fs = (FileStream)saveFileDialog1.OpenFile();
                    Bitmap img = (Bitmap)CaptureDGV.Rows[i].Cells[1].Value;

                    img.Save(fs, ImageFormat.Jpeg);
                    fs.Close();

                }
                else
                {
                    
                    FileStream fs = (FileStream)saveFileDialog1.OpenFile();
                    Bitmap img = (Bitmap)CaptureDGV.Rows[i].Cells[1].Value;

                    img.Save(fs, ImageFormat.Jpeg);
                    fs.Close();
                }
                
                savedLable.Text = "Image Saved";
                myTimer.Interval = 2000;

                savePanel.Visible = true;
                myTimer.Start();
            }


            //if (saveFileDialog1.FileName != ""  )
            //{
            //   FileStream fs = (FileStream)saveFileDialog1.OpenFile();
            //    previewManualLivePB.Image.Save(fs,
            //    ImageFormat.Jpeg);
            //    fs.Close();
            //    savedLable.Text = "Image Saved";
            //    myTimer.Interval = 3000;
                
            //    savePanel.Visible = true;
            //    myTimer.Start();



            //}


            //    previewManualLivePB.Image.Save(@"D:\NCBC\AutomatedSystemNCBC\BACKUPS\28042020\image1.jpeg");
            //MessageBox.Show("saved");

        }

        public static void CreateSpreadsheetWorkbook(string filepath)
        {

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.

            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);
         
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            workbookpart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();

        }

        // Given a document name, inserts a new worksheet.
        public static void InsertWorksheet(string docName)
        {
            int num = dgvAngle.Rows.Count;

            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Add a blank WorksheetPart.
                
                WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                
                Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                // Get the sheetData cells
              
                string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

                


                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Give the new worksheet a name.
                string sheetName = GetSheetNameCell(sheetId);
                

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                

                SharedStringTablePart shareStringPart;
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();

                

                Cell cell1 = InsertCellInWorksheet("A", 1, newWorksheetPart);
                cell1.CellValue = new CellValue(InsertSharedStringItem("Angles ", shareStringPart).ToString());
                cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                

                Cell cell2 = InsertCellInWorksheet("C", 1, newWorksheetPart);
                cell2.CellValue = new CellValue(InsertSharedStringItem("Length", shareStringPart).ToString());
                cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                

                Cell cell3 = InsertCellInWorksheet("E", 1, newWorksheetPart);
                cell3.CellValue = new CellValue(InsertSharedStringItem("Angle Categories", shareStringPart).ToString());
                cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                Cell cell4 = InsertCellInWorksheet("H", 1, newWorksheetPart);
                cell4.CellValue = new CellValue(InsertSharedStringItem("Average Angle", shareStringPart).ToString());
                cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                Cell cell5 = InsertCellInWorksheet("J", 1, newWorksheetPart);
                cell5.CellValue = new CellValue(InsertSharedStringItem("Angle Of Head - Standard Deviation", shareStringPart).ToString());
                cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                


                for(int i = 0; i<dgvLength.Rows.Count; i++)
                {
                    Cell cell = InsertCellInWorksheet("C", Convert.ToUInt32(2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem(dgvLength.Rows[i].Cells["lengthInput"].Value.ToString(), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }


                for (int i = 0; i < dgvAngle.Rows.Count; i++)
                {
                    Cell cell = InsertCellInWorksheet("A", Convert.ToUInt32(i + 2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["angleCol"].Value.ToString(), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                for (int i = 0; i < dgvAngle.Rows.Count; i++)
                {
                    
                    Cell cell = InsertCellInWorksheet("B", Convert.ToUInt32(i + 2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem((dgvAngle.Rows[i].Cells["angleInput"].Value).ToString().Remove((dgvAngle.Rows[i].Cells["angleInput"].Value).ToString().Length - 1), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                for (int i = 0; i < dgvAngle.Rows.Count; i++)
                {
                    Cell cell = InsertCellInWorksheet("E", Convert.ToUInt32(i + 2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["Catagory"].Value.ToString(), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }


                double average = 0, variance = 0, deviation = 0;
                //Come Back
                average = CalcAverage();
                variance = CalcVariance(average);
                deviation = Math.Sqrt(variance);

                Cell cell6 = InsertCellInWorksheet("H", 2, newWorksheetPart);
                cell6.CellValue = new CellValue(InsertSharedStringItem(average.ToString(), shareStringPart).ToString());
                cell6.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                Cell cell7 = InsertCellInWorksheet("J", 2, newWorksheetPart);
                cell7.CellValue = new CellValue(InsertSharedStringItem(deviation.ToString(), shareStringPart).ToString());
                cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);



                // Save the new worksheet.


                newWorksheetPart.Worksheet.Save();
                spreadSheet.Close();
            }


        }
        public static void InsertWorksheetStill(string docName)
        {
            int num = dgvAngle.Rows.Count;

         

            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Add a blank WorksheetPart.
                
                WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                
                Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                // Get the sheetData cells
              
                string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

                


                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Give the new worksheet a name.
                string sheetName = GetSheetNameCell(sheetId);
                

                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);
                

                SharedStringTablePart shareStringPart;
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();

                

                Cell cell1 = InsertCellInWorksheet("A", 1, newWorksheetPart);
                cell1.CellValue = new CellValue(InsertSharedStringItem("Angles ", shareStringPart).ToString());
                cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                

                Cell cell2 = InsertCellInWorksheet("C", 1, newWorksheetPart);
                cell2.CellValue = new CellValue(InsertSharedStringItem("Length", shareStringPart).ToString());
                cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                

                Cell cell3 = InsertCellInWorksheet("E", 1, newWorksheetPart);
                cell3.CellValue = new CellValue(InsertSharedStringItem("Angle Categories", shareStringPart).ToString());
                cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                Cell cell4 = InsertCellInWorksheet("H", 1, newWorksheetPart);
                cell4.CellValue = new CellValue(InsertSharedStringItem("Average Angle", shareStringPart).ToString());
                cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                Cell cell5 = InsertCellInWorksheet("J", 1, newWorksheetPart);
                cell5.CellValue = new CellValue(InsertSharedStringItem("Angle Of Head - Standard Deviation", shareStringPart).ToString());
                cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                


                for(int i = 0; i<dgvLength.Rows.Count; i++)
                {
                    Cell cell = InsertCellInWorksheet("C", Convert.ToUInt32(2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem(dgvLength.Rows[i].Cells["lengthStillInputCol"].Value.ToString(), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }


                for (int i = 0; i < dgvAngle.Rows.Count; i++)
                {
                    Cell cell = InsertCellInWorksheet("A", Convert.ToUInt32(i + 2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["angleStillCol"].Value.ToString(), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                for (int i = 0; i < dgvAngle.Rows.Count; i++)
                {
                    Cell cell = InsertCellInWorksheet("B", Convert.ToUInt32(i + 2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem((dgvAngle.Rows[i].Cells["AngleInputStillCol"].Value).ToString().Remove((dgvAngle.Rows[i].Cells["AngleInputStillCol"].Value).ToString().Length - 1), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                for (int i = 0; i < dgvAngle.Rows.Count; i++)
                {
                    Cell cell = InsertCellInWorksheet("E", Convert.ToUInt32(i + 2), newWorksheetPart);
                    cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["CatStillCol"].Value.ToString(), shareStringPart).ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }


                double average = 0, variance = 0, deviation = 0;
                //Come Back
                average = CalcAverageStill();
                variance = CalcVarianceStill(average);
                deviation = Math.Sqrt(variance);

                Cell cell6 = InsertCellInWorksheet("H", 2, newWorksheetPart);
                cell6.CellValue = new CellValue(InsertSharedStringItem(average.ToString(), shareStringPart).ToString());
                cell6.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                Cell cell7 = InsertCellInWorksheet("J", 2, newWorksheetPart);
                cell7.CellValue = new CellValue(InsertSharedStringItem(deviation.ToString(), shareStringPart).ToString());
                cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);



                // Save the new worksheet.


                newWorksheetPart.Worksheet.Save();
                spreadSheet.Close();
            }


        }

      


     

     

        public static void InsertImageToNewSheet(string docName, Image image)
        {

            string sheetName;
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Add a blank WorksheetPart.

                WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());


                Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                // Get the sheetData cells

                string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);




                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Give the new worksheet a name.
                
              sheetName = GetSheetNameImage(sheetId);
     



                // Append the new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.Append(sheet);


                SharedStringTablePart shareStringPart;
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();

                Cell cell7 = InsertCellInWorksheet("H", 2, newWorksheetPart);
                cell7.CellValue = new CellValue(InsertSharedStringItem( "Capture Taken Of The Microscope View" , shareStringPart).ToString());
                cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                newWorksheetPart.Worksheet.Save();
                spreadSheet.Close();
            }
            
            //local document created to store temporary images
            if (File.Exists(Application.StartupPath + @"\Resources\TempImages"))
            { }
            else
            {
                Directory.CreateDirectory(Application.StartupPath + @"\Resources\TempImages");

            }
            
            //Converted to bitmap to resize
            Bitmap newBit = ResizeImage(image, 400, 250);

            image = (Image)newBit;
            //Reverted to Image to save in local document 
            image.Save(Application.StartupPath + @"\Resources\TempImages\tempImage.jpeg");

            try
            {
                //adds image to the excel document given filename and sheet name
                ExcelTools.AddImage(false, docName, sheetName,
                                    Application.StartupPath + @"\Resources\TempImages\tempImage.jpeg", "Sperm Cell 1",
                                    1 /* column */, 1 /* row */);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
            image.Dispose();

            //delete the temporary image
            //File.Delete(Application.StartupPath + @"\Resources\TempImages\tempImage.jpeg");



            
        }

        private static string GetSheetNameImage(uint sheetId)
        {
            string name = "";
           switch(sheetId)
            {
#pragma warning disable CS0162 // Unreachable code detected
                case 2: return name = "Microscope Image Capture " + (sheetId - 1).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 4: return name = "Microscope Image Capture " + (sheetId - 2).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 6: return name = "Microscope Image Capture " + (sheetId - 3).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 8: return name = "Microscope Image Capture " + (sheetId - 4).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 10: return name = "Microscope Image Capture " + (sheetId - 5).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 12: return name = "Microscope Image Capture " + (sheetId - 6).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 14: return name = "Microscope Image Capture " + (sheetId - 7).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 16: return name = "Microscope Image Capture " + (sheetId - 8).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 18: return name = "Microscope Image Capture " + (sheetId - 9).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 20: return name = "Microscope Image Capture " + (sheetId - 10).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 22: return name = "Microscope Image Capture " + (sheetId - 11).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 24: return name = "Microscope Image Capture " + (sheetId - 12).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 26: return name = "Microscope Image Capture " + (sheetId - 13).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 28: return name = "Microscope Image Capture " + (sheetId - 14).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 30: return name = "Microscope Image Capture " + (sheetId - 15).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 32: return name = "Microscope Image Capture " + (sheetId - 16).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 34: return name = "Microscope Image Capture " + (sheetId - 17).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 36: return name = "Microscope Image Capture " + (sheetId - 18).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 38: return name = "Microscope Image Capture " + (sheetId - 19).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 40: return name = "Microscope Image Capture " + (sheetId - 20).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 42: return name = "Microscope Image Capture " + (sheetId - 21).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 44: return name = "Microscope Image Capture " + (sheetId - 22).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 46: return name = "Microscope Image Capture " + (sheetId - 23).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 48: return name = "Microscope Image Capture " + (sheetId - 24).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 50: return name = "Microscope Image Capture " + (sheetId - 25).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                default: return name = "Microscope Image Capture " + (sheetId).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected

            }
        }
        private static string GetSheetNameCell(uint sheetId)
        {
            string name = "";
           switch(sheetId)
            {

#pragma warning disable CS0162 // Unreachable code detected
                case 1: return name = "Sperm Cell " + (sheetId ).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 3: return name = "Sperm Cell " + (sheetId - 1).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 5: return name = "Sperm Cell " + (sheetId - 2).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 7: return name = "Sperm Cell " + (sheetId - 3).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 9: return name = "Sperm Cell " + (sheetId - 4).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 11: return name = "Sperm Cell " + (sheetId - 5).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 13: return name = "Sperm Cell " + (sheetId - 6).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 15: return name = "Sperm Cell " + (sheetId - 7).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 17: return name = "Sperm Cell " + (sheetId - 8).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 19: return name = "Sperm Cell " + (sheetId - 9).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 21: return name = "Sperm Cell " + (sheetId - 10).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 23: return name = "Sperm Cell " + (sheetId - 11).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 25: return name = "Sperm Cell " + (sheetId - 12).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 27: return name = "Sperm Cell " + (sheetId - 13).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 29: return name = "Sperm Cell " + (sheetId - 14).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 31: return name = "Sperm Cell " + (sheetId - 15).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 33: return name = "Sperm Cell " + (sheetId - 16).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 35: return name = "Sperm Cell " + (sheetId - 17).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 37: return name = "Sperm Cell " + (sheetId - 18).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 39: return name = "Sperm Cell " + (sheetId - 19).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 41: return name = "Sperm Cell " + (sheetId - 20).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 43: return name = "Sperm Cell " + (sheetId - 21).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 45: return name = "Sperm Cell " + (sheetId - 22).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 47: return name = "Sperm Cell " + (sheetId - 23).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                case 49: return name = "Sperm Cell " + (sheetId - 24).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected
#pragma warning disable CS0162 // Unreachable code detected
                default: return name = "Sperm Cell " + (sheetId).ToString(); break;
#pragma warning restore CS0162 // Unreachable code detected

            }
        }

        public static double CalcVariance(double average)
        {
            double variance =0;
            double angle = 0;
            int tempCount = 0;
            int num = dgvAngle.Rows.Count;

            for (int i = 0; i<num; i++)
            {
                if((dgvAngle.Rows[i].Cells["Catagory"].Value).ToString() == "Angle Of Midpiece")
                {
                }
                else
                {
                    angle = Convert.ToDouble((dgvAngle.Rows[i].Cells["angleInput"].Value).ToString().Remove((dgvAngle.Rows[i].Cells["angleInput"].Value).ToString().Length - 1));
                    variance += (angle - average) * (angle - average);
                    tempCount++;
                }
               
            }

           

            return variance/tempCount;
        }
        int tempCount = 0;
        

 
        private void NCBCApplication_Activated(object sender, EventArgs e)
        {
            if (cameraManualLB.Items.Count != 0)
            {
                if (tempCount == 0)
                {
                    tempCount++;
                }
                else
                {
                    CaptureDevice.Start();

                }
            }
        }
        

        private void UndoImageBtn_Click(object sender, EventArgs e)
        {
            if (CaptureDGV.Rows.Count > 0)
            {
                CaptureDGV.Rows.Remove(CaptureDGV.Rows[CaptureDGV.Rows.Count-1]);
                imageDGVCount--;
                tempCounter--;
            }
            else
            {
                tempCounter = 0;
            }
        }

        private void ClearImageBtn_Click(object sender, EventArgs e)
        {
           
            CaptureDGV.Rows.Clear();
            imageDGVCount = 0;
            tempCounter = 0;
        }

        private void AutoPAM_KeyDown(object sender, KeyEventArgs e)
        {
          //  if (stillImageTab1.SelectedIndex == 2)
          //  {
          //      if (e.Control == true && e.KeyData == Keys.S)
          //      {
          //          ExportSheetManualLiveBtn.PerformClick();
          //      }
          //  }
          //else if (stillImageTab1.SelectedIndex == 3)
          //  {
          //      if (e.Control == true && e.KeyData == Keys.S)
          //      {
          //          ExportSheetManualStillBtn.PerformClick();
          //      }
          //  }
        }

      

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            closeBtn3.PerformClick();
        }

        private void FullScreenToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if(CaptureDevice.IsRunning)
                CloseCamera();
            FullScreenForm cmb = new FullScreenForm
            {
                Dgv = liveDGV
            };

            cmb.ShowDialog();

            if (cmb.Dgv.Rows.Count != 0)
            {
                for(int i = 0; i<cmb.Dgv.Rows.Count; i++)
                {
                    
                    liveDGV.Rows.Add();
                    liveDGV.Rows[angleNumLive-1].Cells["angleCol"].Value = cmb.Dgv.Rows[i].Cells["angleCol"].Value;
                    liveDGV.Rows[angleNumLive-1].Cells["angleInput"].Value = cmb.Dgv.Rows[i].Cells["angleInput"].Value;
                    liveDGV.Rows[angleNumLive-1].Cells["Catagory"].Value = cmb.Dgv.Rows[i].Cells["Catagory"].Value;
                    angleNumLive++;
                    catCount++;
                }

                liveDGV.FirstDisplayedScrollingRowIndex = liveDGV.RowCount - 1;
                cmb.Dispose();
               
               
            }
        }

        private void changeCameraToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (CaptureDevice.IsRunning)
                CloseCamera();
            changeCam cmb = new changeCam();
            cmb.backgroundColour = backColourLive.BackColor;
            cmb.ShowDialog();

            if (cmb.selected == false)
                return;


            if (CaptureDevice.IsRunning)
                CloseCamera();

            if (cameraManualLB.Items.Count >0)
            {
                cameraManualLB.Items.Clear();
            }

            FillInDropDownCamera();
            cameraManualLB.SelectedIndex = cmb.selectedIndex;
            this.CaptureDevice.Source = (colCamera[cameraManualLB.SelectedIndex].MonikerString);
            this.CaptureDevice.NewFrame += new NewFrameEventHandler(FinalFrame_NewFrame);
            this.CaptureDevice.Start();
           
            


        }

        private void changeDefaultCatagoryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewCatWindowcs newCat = new NewCatWindowcs(Catagory);

             newCat.BackColor = backColourLive.BackColor;
             newCat.ShowDialog();

            if (newCat.preClosed == false)
                return;

            try
            {
                for (int i = 0; i < liveDGV.Rows.Count; i++)
                {
                    liveDGV.Rows[i].Cells[2].Value = null;
                }

                Catagory.Items.Clear();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Personal Exception");
            }

          
            try
            {
                for (int i = 0; i < newCat.Dgv.Items.Count; i++)
                {
                    Catagory.Items.Add(newCat.Dgv.Items[i].ToString());
                }
                //  MessageBox.Show(Catagory.Items[0].ToString());
            }catch(Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Personal Exception");
            }
            catCount = 0;
            try
            {
                for (int i = 0; i < liveDGV.Rows.Count; i++)
                {
                    if (catCount == Catagory.Items.Count)
                    {
                        catCount = 0;
                        liveDGV.Rows[i].Cells["Catagory"].Value = Catagory.Items[catCount];
                    }
                    else
                        liveDGV.Rows[i].Cells["Catagory"].Value = Catagory.Items[catCount];

                    catCount++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Personal Exception");
            }

            if (File.Exists(Application.StartupPath + @"\Resources\DefaultCatagories"))
            { }
            else
            {
                Directory.CreateDirectory(Application.StartupPath + @"\Resources\DefaultCatagories");

            }
            using (System.IO.StreamWriter file = new StreamWriter(Application.StartupPath + @"\Resources\DefaultCatagories\Cat.txt"))
            {
                for (int i = 0; i < Catagory.Items.Count; i++)
                {
                    file.WriteLine(Catagory.Items[i].ToString());
                }

                //string big = "Bing was his name o \n B I N G O \n B I N G O ";

                //if (big.Contains("\n"))
                //    MessageBox.Show("YES " + big);
                //else
                //    MessageBox.Show("NO " + big);
            }

        }

        private void changeDefaultCaptureToolStripMenuItem_Click(object sender, EventArgs e)
        {

            changeStringFormcs cmb = new changeStringFormcs(captureString);
            cmb.backgroundColour = backColourLive.BackColor;
            cmb.ShowDialog();

            captureString = cmb.output;

           
        }

        private void changeColourSchemeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeColourForm cmb = new ChangeColourForm(backColourLive.BackColor);
            
            cmb.ShowDialog();

            backColourLive.BackColor = cmb.backGColor;
            BackcolourStill.BackColor = cmb.backGColor;


            if (File.Exists(Application.StartupPath + @"\Resources\ColourScheme"))
            { }
            else
            {
                Directory.CreateDirectory(Application.StartupPath + @"\Resources\ColourScheme");

            }
            using (System.IO.StreamWriter file = new StreamWriter(Application.StartupPath + @"\Resources\ColourScheme\colour.txt"))
            {
                file.WriteLine(cmb.backGColor.ToArgb());
              
            }



        }

        private void changeDefaultSaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string filePath = "", startupPath = "";
            if (File.Exists(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
            {
                startupPath = File.ReadAllText(Application.StartupPath + @"\Resources\lastSave\filepath.txt");

                startupPath = startupPath.Remove(startupPath.Length - 2);

            }
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {

                    fbd.SelectedPath = startupPath;
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        filePath = fbd.SelectedPath.ToString();
                    }
                    else
                    { this.Cursor = Cursors.Default; return; }
                }
                if (File.Exists(Application.StartupPath + @"\Resources\lastSave"))
                { }
                else
                {
                    Directory.CreateDirectory(Application.StartupPath + @"\Resources\lastSave");

                }
                using (System.IO.StreamWriter file = new StreamWriter(Application.StartupPath + @"\Resources\lastSave\filepath.txt"))
                {
                    file.WriteLine(filePath);
                }
            
        }

     

        private void CaptureDGV_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if(CaptureDGV.Rows.Count <= 1)
            { }
            else if((CaptureDGV.Rows[imageDGVCount - 1].Cells[0].Value).ToString().Contains(')'))
            { }
            else
            {
                captureString = (CaptureDGV.Rows[imageDGVCount - 1].Cells[0].Value).ToString();
                tempCounter = 0;
            }
        }

        private void NCBCApplication_Deactivate(object sender, EventArgs e)
        {
            if (cameraManualLB.Items.Count != 0)
                    CloseCamera();
        }

        public static double CalcVarianceStill(double average)
        {
            double variance = 0;
            double angle = 0;
            int tempCount = 0;
            int num = dgvAngle.Rows.Count;

            for (int i = 0; i < num; i++)
            {
                if ((dgvAngle.Rows[i].Cells["CatStillCol"].Value).ToString() == "Angle Of Midpiece")
                {
                }
                else
                {
                    angle = Convert.ToDouble((dgvAngle.Rows[i].Cells["AngleInputStillCol"].Value).ToString().Remove((dgvAngle.Rows[i].Cells["AngleInputStillCol"].Value).ToString().Length - 1));
                    variance += (angle - average) * (angle - average);
                    tempCount++;
                }

            }



            return variance / tempCount;
        }
        public static double CalcAverage()
        {
            double average = 0;

            int num = dgvAngle.Rows.Count;
            int tempcount = 0;
            for(int i = 0; i< dgvAngle.Rows.Count; i++)
            {
                if ((dgvAngle.Rows[i].Cells["Catagory"].Value).ToString() == "Angle Of Midpiece")
                {
                    
                }
                else
                {
                    average += Convert.ToDouble((dgvAngle.Rows[i].Cells["angleInput"].Value).ToString().Remove((dgvAngle.Rows[i].Cells["angleInput"].Value).ToString().Length - 1));
                    tempcount++;
                }
            }
            average = average / tempcount;
            return average;


        }
        public static double CalcAverageStill()
        {
            double average = 0;

            int num = dgvAngle.Rows.Count;
            int tempcount = 0;
            for(int i = 0; i< dgvAngle.Rows.Count; i++)
            {
                if ((dgvAngle.Rows[i].Cells["CatStillCol"].Value).ToString() == "Angle Of Midpiece")
                {
                    
                }
                else
                {
                    average += Convert.ToDouble((dgvAngle.Rows[i].Cells["AngleInputStillCol"].Value).ToString().Remove((dgvAngle.Rows[i].Cells["AngleInputStillCol"].Value).ToString().Length - 1));
                    tempcount++;
                }
            }
            average = average / tempcount;
            return average;


        }

        //public void ReadXml()
        //{
        //    var filePath = @"D:\NCBC\AutomatedSystemNCBC\excel\Book1.xlsx";
        //    using (var document = SpreadsheetDocument.Open(filePath, false))
        //    {
        //        var workbookPart = document.WorkbookPart;
        //        var workbook = workbookPart.Workbook;

        //        var sheets = workbook.Descendants<Sheet>();
        //        foreach (var sheet in sheets)
        //        {
        //            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        //            var sharedStringPart = workbookPart.SharedStringTablePart;
        //            var values = sharedStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();

        //            var cells = worksheetPart.Worksheet.Descendants<Cell>();
        //            string cellV = "";

        //            foreach (var cell in cells)
        //            {
        //                //Console.WriteLine(cell.CellReference);
        //                // The cells contains a string input that is not a formula
        //                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        //                {
        //                    var index = int.Parse(cell.CellValue.Text);
        //                    var value = values[index].InnerText;

        //                    cellV += value.ToString() + "/";

        //                    //  Console.WriteLine(value);
        //                }
        //                else
        //                {

        //                    // Console.WriteLine(cell.CellValue.Text);

        //                }

        //                if (cell.CellFormula != null)
        //                {
        //                    Console.WriteLine(cell.CellFormula.Text);
        //                }

        //            }


        //            //foreach (Control c in xmlPanel.Controls)
        //            //{
        //            //    if (c is Label)
        //            //    {
        //            //        string[] site = null;
        //            //        char[] splitsite = { '/' };
        //            //        site = cellV.Split(splitsite);

        //            //        for (int i = site.Length - 1; i >= 0; i--)
        //            //        {
        //            //            if (i == site.Length - 1)
        //            //            {
        //            //                ((Label)c).Text = site[i].ToString();
        //            //            }
        //            //            else
        //            //            {
        //            //            }
        //            //        }
        //            //    }
        //            //}
        //        }

        //    }


        //    // Console.ReadLine();

        //    //for (int i = 0; i < 26; i++)
        //    //{
        //    //    InsertText(filePath, Convert.ToChar(i + (int)'A') + " - ");
        //    //}

        //    InsertText(filePath, "BEBE");
        //    MessageBox.Show("Excel Document Created, \nYou Can Find the Document at:\n" + filePath, "Excel Document Added");

        //}




        // Given a document name and text, 
        // inserts a new worksheet and writes  text into the new worksheet.
        public static bool InsertText(string docName)
        {

            try
            {
            
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
                {

                    // Get the SharedStringTablePart. If it does not exist, create a new one.
                    SharedStringTablePart shareStringPart;
                    if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {

                        shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();

                    }


                    // Insert the text into the SharedStringTablePart.
                    // int index = InsertSharedStringItem(text, shareStringPart);

                    // Insert a new worksheet.
                    WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
                    // Insert cell A1 into the new worksheet.




                    Cell cell1 = InsertCellInWorksheet("A", 1, worksheetPart);
                    cell1.CellValue = new CellValue(InsertSharedStringItem("Angles ", shareStringPart).ToString());
                    cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                    Cell cell2 = InsertCellInWorksheet("C", 1, worksheetPart);
                    cell2.CellValue = new CellValue(InsertSharedStringItem("Length", shareStringPart).ToString());
                    cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);



                    Cell cell3 = InsertCellInWorksheet("E", 1, worksheetPart);
                    cell3.CellValue = new CellValue(InsertSharedStringItem("Angle Categories", shareStringPart).ToString());
                    cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                    Cell cell4 = InsertCellInWorksheet("H", 1, worksheetPart);
                    cell4.CellValue = new CellValue(InsertSharedStringItem("Average Angle", shareStringPart).ToString());
                    cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                    Cell cell5 = InsertCellInWorksheet("J", 1, worksheetPart);
                    cell5.CellValue = new CellValue(InsertSharedStringItem("Angle Of Head - Standard Deviation", shareStringPart).ToString());
                    cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                  


                    for (int i = 0; i < dgvLength.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("C", Convert.ToUInt32(2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvLength.Rows[i].Cells["lengthInput"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }


                    for (int i = 0; i < dgvAngle.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("A", Convert.ToUInt32(i + 2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["angleCol"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                    for (int i = 0; i < dgvAngle.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("B", Convert.ToUInt32(i + 2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["angleInput"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                    for (int i = 0; i < dgvAngle.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("E", Convert.ToUInt32(i + 2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["Catagory"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }


                    double average = 0, variance = 0, deviation = 0;

                    average = CalcAverage();
                    variance = CalcVariance(average);
                    deviation = Math.Sqrt(variance);

                    Cell cell6 = InsertCellInWorksheet("H", 2, worksheetPart);
                    cell6.CellValue = new CellValue(InsertSharedStringItem(average.ToString(), shareStringPart).ToString());
                    cell6.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    Cell cell7 = InsertCellInWorksheet("J", 2, worksheetPart);
                    cell7.CellValue = new CellValue(InsertSharedStringItem(deviation.ToString(), shareStringPart).ToString());
                    cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);



                   // Save the new worksheet.


                   // Save the new worksheet.
                   worksheetPart.Worksheet.Save();

                    spreadSheet.Close();

                    return true;
                }
            }catch(Exception)
            {
                
                MessageBox.Show("Please Close Any Existing Excel Documents");
                return false;
            }
        }

        public static bool InsertTextStill(string docName)
        {

            try
            {

                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
                {

                    // Get the SharedStringTablePart. If it does not exist, create a new one.
                    SharedStringTablePart shareStringPart;
                    if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {

                        shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();

                    }


                    // Insert the text into the SharedStringTablePart.
                    // int index = InsertSharedStringItem(text, shareStringPart);

                    // Insert a new worksheet.
                    WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
                    // Insert cell A1 into the new worksheet.




                    Cell cell1 = InsertCellInWorksheet("A", 1, worksheetPart);
                    cell1.CellValue = new CellValue(InsertSharedStringItem("Angles ", shareStringPart).ToString());
                    cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                    Cell cell2 = InsertCellInWorksheet("C", 1, worksheetPart);
                    cell2.CellValue = new CellValue(InsertSharedStringItem("Length", shareStringPart).ToString());
                    cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);



                    Cell cell3 = InsertCellInWorksheet("E", 1, worksheetPart);
                    cell3.CellValue = new CellValue(InsertSharedStringItem("Angle Categories", shareStringPart).ToString());
                    cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                    Cell cell4 = InsertCellInWorksheet("H", 1, worksheetPart);
                    cell4.CellValue = new CellValue(InsertSharedStringItem("Average Angle", shareStringPart).ToString());
                    cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);


                    Cell cell5 = InsertCellInWorksheet("J", 1, worksheetPart);
                    cell5.CellValue = new CellValue(InsertSharedStringItem("Angle Of Head - Standard Deviation", shareStringPart).ToString());
                    cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);




                    for (int i = 0; i < dgvLength.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("C", Convert.ToUInt32(2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvLength.Rows[i].Cells["lengthStillInputCol"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }


                    for (int i = 0; i < dgvAngle.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("A", Convert.ToUInt32(i + 2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["angleStillCol"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                    for (int i = 0; i < dgvAngle.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("B", Convert.ToUInt32(i + 2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["AngleInputStillCol"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                    for (int i = 0; i < dgvAngle.Rows.Count; i++)
                    {
                        Cell cell = InsertCellInWorksheet("E", Convert.ToUInt32(i + 2), worksheetPart);
                        cell.CellValue = new CellValue(InsertSharedStringItem(dgvAngle.Rows[i].Cells["CatStillCol"].Value.ToString(), shareStringPart).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }


                    double average = 0, variance = 0, deviation = 0;

                    average = CalcAverageStill();
                    variance = CalcVarianceStill(average);
                    deviation = Math.Sqrt(variance);

                    Cell cell6 = InsertCellInWorksheet("H", 2, worksheetPart);
                    cell6.CellValue = new CellValue(InsertSharedStringItem(average.ToString(), shareStringPart).ToString());
                    cell6.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    Cell cell7 = InsertCellInWorksheet("J", 2, worksheetPart);
                    cell7.CellValue = new CellValue(InsertSharedStringItem(deviation.ToString(), shareStringPart).ToString());
                    cell7.DataType = new EnumValue<CellValues>(CellValues.SharedString);



                    // Save the new worksheet.


                    // Save the new worksheet.
                    worksheetPart.Worksheet.Save();



                    return true;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Please Close Any Existing Excel Documents");
                return false;
            }
        }


        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName;
            sheetName = GetSheetNameCell(sheetId);

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 


        // Given a WorkbookPart, inserts a new worksheet.

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
           
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

           
            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex};
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);


                worksheet.Save();
                return newCell;
            }
        }

    }
}
