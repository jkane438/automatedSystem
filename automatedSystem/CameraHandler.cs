using AForge.Video;
using AForge.Video.DirectShow;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace automatedSystem
{
    class CameraHandler
    {

        private readonly string mes = "Created By Jordan Kane";
            private FilterInfoCollection videoDevices = null;
            private VideoCaptureDevice videoSource = null;
            private PictureBox pbCamPreview = null;
            private Bitmap currentImage = null;
           




        public void Dispose()
        {
            videoDevices = null;
            videoSource = null;
            pbCamPreview = null;
            currentImage = null;

        }


            // initialize with picturebox - used to load new frames later on
            public CameraHandler()
            {
                //PictureBox pbCamPreview
                //this.pbCamPreview = pbCamPreview;
            }


            public Bitmap GetSnapshot()
            {
                return currentImage;
            }

        public bool IsRunning()
        {
            if (videoSource.IsRunning)
                return true;
            else
                return false;
        }
            // this function stops the image capturing 
            public void StopCapture()
            {
              //  if (!(videoSource == null))
                    if (videoSource.IsRunning)
                    {
                        videoSource.Stop();
                        videoSource.SignalToStop();
                        //videoSource = null;
                        
                    }

            }

        public string Method01()
        {
            return mes;
        }
            // starts the video capturering function
            public bool StartCapture()
            {
           
                try
                {

                    videoSource.Start();
                    return true;

                }
                catch (Exception)
            {
                    return false;
                }
            }

            // sets a VideoSource to capture images
            public bool SetVideoSourceByIndex(int index)
            {
              
                videoSource = null;
                try
                {
                    
                    videoSource = new VideoCaptureDevice(videoDevices[index].MonikerString);
                    videoSource.NewFrame += new NewFrameEventHandler(SetNewFrame);
                    return true;
                }
                catch
                {

                    return false;
                }
            }

            /* refreshes the list of video sources for further processing and 
             * returns a list of all devices, found on the system */
            public FilterInfoCollection RefreshCameraList()
            {
                videoDevices = null;
                try
                {
                    videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                    if (videoDevices.Count == 0)
                        throw new ApplicationException();
                }
                catch (ApplicationException)
                {

                }
                return videoDevices;
            }




        //    public void setBright(int _bright)
        //    {

        //        this.bright = _bright;
        //    try
        //    {
               
               
               
        //    }
        //    catch (ApplicationException)
        //    {

        //    }
        //}   

        //eventhandler for every new frame
        private void SetNewFrame(object sender,  NewFrameEventArgs eventArgs)
            {

            //increases brightness of camera feed
            //float value = bright * 0.01f;
            //float[][] colorMatrixElements = {
            //                                new float[] {1,0,0,0,0},
            //                                new float[] {0,1,0,0,0},
            //                                new float[] {0,0,1,0,0},
            //                                new float[] {0,0,0,1,0},
            //                                new float[] {value,value,value,0,1}
            //                                };
            //ColorMatrix colorMatrix = new ColorMatrix(colorMatrixElements);
            //ImageAttributes imageAttributes = new ImageAttributes();


            //imageAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);



                Image _img = (Bitmap)eventArgs.Frame.Clone(); 
                //PictureBox1.Image
                Graphics _g = default(Graphics);
                Bitmap bm_dest = new Bitmap(Convert.ToInt32(_img.Width), Convert.ToInt32(_img.Height));
                _g = Graphics.FromImage(bm_dest);
                _g.DrawImage(_img, new Rectangle(0, 0, bm_dest.Width + 1, bm_dest.Height + 1), 0, 0, bm_dest.Width + 1, bm_dest.Height + 1, GraphicsUnit.Pixel);
            



            try
                {
                currentImage = (Bitmap)eventArgs.Frame.Clone();
                currentImage = bm_dest;
                currentImage.RotateFlip(RotateFlipType.RotateNoneFlipX);
                //do show image in assigned picturebox
                pbCamPreview.Image = currentImage;
                }
                catch (Exception)
            {

                }
            }

        }
    }

