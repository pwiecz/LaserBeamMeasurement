using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LaserBeamMeasurement
{
    public partial class Form1 : Form
    {

        bool hand_mode = false;

        int mouse_x = 0;
        int mouse_y = 0;

        bool no_image = true;

        double pixsize = 2.2;

        ImageData _imagedata = new ImageData();
        BeamParameters _beamparameters = new BeamParameters();

        private ICamera _camera = null;

        private bool _captureInProgress;
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;

        static ICamera CreateCamera()
        {
            UEyeCamera uEyeCamera = new UEyeCamera();
            if (uEyeCamera.Init())
            {
                return uEyeCamera;
            }
            return new VideoCaptureCamera();
        }

        public Form1()
        {


            InitializeComponent();


            dataGridView1.Rows.Add();
            dataGridView1.Rows.Add();
            dataGridView1.RowHeadersWidth = 78;
            dataGridView1.Rows[0].HeaderCell.Value = "FWHM";
            dataGridView1.Rows[1].HeaderCell.Value = "1/e^2";
            dataGridView1.Rows.Add();
            dataGridView1.Rows[2].HeaderCell.Value = "DIV";

            xWidthChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
            yWidthChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
            xWidthFilteredChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
            yWidthFilteredChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;

            tabControl1.SelectedIndex = 0;

            CvInvoke.UseOpenCL = false;

            _camera = CreateCamera();
            if (_camera != null)
            {
                _camera.OnNewFrame += ProcessFrame;
            }
        }

        private static dynamic CreateElement(DepthType depthType)
        {
            switch (depthType)
            {
                case DepthType.Cv8S: return new sbyte[1];
                case DepthType.Cv8U: return new byte[1];
                case DepthType.Cv16S: return new short[1];
                case DepthType.Cv16U: return new ushort[1];
                case DepthType.Cv32S: return new int[1];
                case DepthType.Cv32F: return new float[1];
                case DepthType.Cv64F: return new double[1];
            }
            return new float[1];
        }
        public void ProcessFrame(object sender, EventArgs arg) // will be simplified
        {
            Mat grayFrame = _camera.GetGrayFrame();

            double thresh_med;
            double thresh_e2;

            System.Drawing.Image OrgImage;
            System.Drawing.Image OrgImage1;

            Image<Rgb, Byte> tothermo = grayFrame.ToImage<Rgb, Byte>(); // original
            Image<Gray, Byte> tothermo1 = grayFrame.ToImage<Gray, Byte>();

            OrgImage = tothermo.ToBitmap();
            OrgImage1 = tothermo.ToBitmap();

            Invoke((MethodInvoker)(
                () =>
                {
                    dynamicGrayPictureBox.Image = OrgImage1;
                    dynamicGrayPictureBox.Refresh();
                    _imagedata.MakeFalse((Bitmap)OrgImage);
                    dynamicThermoPictureBox.Image = OrgImage;
                    dynamicThermoPictureBox.Refresh();
                    double[] minVal;
                    double[] maxVal;
                    System.Drawing.Point[] minLoc;
                    System.Drawing.Point[] maxLoc;
                    grayFrame.MinMax(out minVal, out maxVal, out minLoc, out maxLoc);
                    _imagedata.sizex = OrgImage.Width;
                    _imagedata.sizey = OrgImage.Height;

                    if (hand_mode)
                    {
                        _imagedata.centerx = mouse_x;
                        _imagedata.centery = mouse_y;
                    }
                    else
                    {
                        _imagedata.centerx = maxLoc[0].X;
                        _imagedata.centery = maxLoc[0].Y;
                    }

                    dynamicThermoPictureBox.Refresh();
                    dynamicGrayPictureBox.Refresh();

                    //_imagedata.GraphFill(tothermo1);
                    _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartXstartX, _imagedata.ChartXstartY, _imagedata.ChartXstopX, _imagedata.ChartXstopY, true);
                    _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartYstartX, _imagedata.ChartYstartY, _imagedata.ChartYstopX, _imagedata.ChartYstopY, false);
                    _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartXstartX, _imagedata.ChartXstartY, _imagedata.ChartXstopX, _imagedata.ChartXstopY, true);
                    _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartYstartX, _imagedata.ChartYstartY, _imagedata.ChartYstopX, _imagedata.ChartYstopY, false);

                    // thresh calculate
                    if (hand_mode)
                    {
                        var value = CreateElement(grayFrame.Depth);
                        Marshal.Copy(grayFrame.GetDataPointer(new int[] { _imagedata.centerx, _imagedata.centery }), value, 0, 1);
                        thresh_med = value[0] / 2;
                        thresh_e2 = value[0] / Math.Exp(2);
                    }
                    else
                    {
                        thresh_med = maxVal[0] / 2;
                        thresh_e2 = maxVal[0] / Math.Exp(2);
                    }

                    {
                        Image<Gray, Byte> threshold = new Image<Gray, byte>(grayFrame.Width, grayFrame.Height);
                        CvInvoke.Threshold(grayFrame, threshold, thresh_e2, 100, ThresholdType.Binary);
                        System.Drawing.Rectangle boundRect_e2 = new System.Drawing.Rectangle();
                        CvInvoke.FloodFill(threshold, null, new System.Drawing.Point(_imagedata.centerx, _imagedata.centery), new MCvScalar(1),
                            out boundRect_e2, new MCvScalar(5), new MCvScalar(1), Connectivity.FourConnected, FloodFillType.MaskOnly);
                        _imagedata.boundRect_e2 = boundRect_e2;
                    }
                    {
                        Image<Gray, Byte> threshold = new Image<Gray, byte>(grayFrame.Width, grayFrame.Height);
                        CvInvoke.Threshold(grayFrame, threshold, thresh_med, 100, ThresholdType.Binary);
                        System.Drawing.Rectangle boundRect_med = new System.Drawing.Rectangle();
                        CvInvoke.FloodFill(threshold, null, new System.Drawing.Point(_imagedata.centerx, _imagedata.centery), new MCvScalar(1),
                            out boundRect_med, new MCvScalar(5), new MCvScalar(1), Connectivity.FourConnected, FloodFillType.MaskOnly);
                        _imagedata.boundRect_med = boundRect_med;
                    }

                    _beamparameters.BeamSizeDetect(thresh_med, thresh_e2, _imagedata);
                }));
        }

        public void ProcessStaticFrame(Image ImFromFile, bool filt)
        {
            double thresh_med;
            double thresh_e2;
            // Converting the master image to a bitmap
            Bitmap masterImage = (Bitmap)ImFromFile;

            // Normalizing it to grayscale
            Image<Gray, Byte> grayFrame = masterImage.ToImage<Gray, Byte>();//new Image<Gray, Byte>(masterImage);
            Image<Gray, Byte> grayFramefilt = masterImage.ToImage<Gray, Byte>();//new Image<Gray, Byte>(masterImage);

            double[] minVal;
            double[] maxVal;
            System.Drawing.Point[] minLoc;
            System.Drawing.Point[] maxLoc;
            grayFrame.MinMax(out minVal, out maxVal, out minLoc, out maxLoc);

            if (filt)
                CvInvoke.MedianBlur(grayFrame, grayFrame, 7);

            _imagedata.sizex = grayFrame.Width;
            _imagedata.sizey = grayFrame.Height;

            if (hand_mode)
            {
                _imagedata.centerx = mouse_x;
                _imagedata.centery = mouse_y;
            }
            else
            {
                _imagedata.centerx = maxLoc[0].X;
                _imagedata.centery = maxLoc[0].Y;
            }

            // fill 3D plot data
            int width = 200;
            int height = 200;

            int k = 0;
            int m = 0;
            double rotation_angle = (int)angleNumericUpDown.Value * Math.PI / 180;

            double rotation_sin = Math.Sin(rotation_angle);
            double rotation_cos = Math.Cos(rotation_angle);

            for (int i = -width; i < width; i++)
            {
                for (int j = -height; j < height; j++)
                {
                    int grayFrameX = (int)(i * rotation_sin + j * rotation_cos) + _imagedata.centery;
                    int grayFrameY = (int)(i * rotation_cos - j * rotation_sin) + _imagedata.centerx;
                    if (grayFrameX >= 0 && grayFrameX < grayFrame.Data.GetLength(1) &&
                        grayFrameY >= 0 && grayFrameY < grayFrame.Data.GetLength(0))
                    {
                        _imagedata.Graph3d[k, m] = grayFrame.Data[grayFrameY, grayFrameX, 0];
                    }
                    m++;
                }
                k++; m = 0;
            }

            /*

            for (int i = _imagedata.centerx - width; i < _imagedata.centerx + width; i++)
            {
                for (int j = _imagedata.centery - height; j < _imagedata.centery + height; j++)

                {
                    _imagedata.Graph3d[k, m] = grayFrame.Data[j, i, 0];
                    m++;
                }

                k++; m = 0;

            } */

            _imagedata.GraphFillRotate(grayFrame, _imagedata.ChartXstartX, _imagedata.ChartXstartY, _imagedata.ChartXstopX, _imagedata.ChartXstopY, true);
            _imagedata.GraphFillRotate(grayFrame, _imagedata.ChartYstartX, _imagedata.ChartYstartY, _imagedata.ChartYstopX, _imagedata.ChartYstopY, false);

            _imagedata.zero_level = Convert.ToInt32(zeroLevelTextBox.Text);

            // thresh calculate
            if (hand_mode)
            {
                thresh_med = (grayFrame.Data[_imagedata.centery, _imagedata.centerx, 0] - _imagedata.zero_level) / 2 + _imagedata.zero_level;
                thresh_e2 = (grayFrame.Data[_imagedata.centery, _imagedata.centerx, 0] - _imagedata.zero_level) / Math.Exp(2) + _imagedata.zero_level;
            }
            else
            {
                thresh_med = (maxVal[0] - _imagedata.zero_level) / 2 + _imagedata.zero_level;
                thresh_e2 = (maxVal[0] - _imagedata.zero_level) / Math.Exp(2) + _imagedata.zero_level;
            }
            {
                Image<Gray, Byte> threshold = new Image<Gray, byte>(grayFrame.Width, grayFrame.Height);
                CvInvoke.Threshold(grayFrame, threshold, thresh_e2, 100, ThresholdType.Binary);
                System.Drawing.Rectangle boundRect_e2 = new System.Drawing.Rectangle();
                CvInvoke.FloodFill(threshold, null, new System.Drawing.Point(_imagedata.centerx, _imagedata.centery), new MCvScalar(1),
                    out boundRect_e2, new MCvScalar(5), new MCvScalar(1), Connectivity.FourConnected, FloodFillType.MaskOnly);
                _imagedata.boundRect_e2 = boundRect_e2;
            }
            {
                Image<Gray, Byte> threshold = new Image<Gray, byte>(grayFrame.Width, grayFrame.Height);
                CvInvoke.Threshold(grayFrame, threshold, thresh_med, 100, ThresholdType.Binary);
                System.Drawing.Rectangle boundRect_med = new System.Drawing.Rectangle();
                CvInvoke.FloodFill(threshold, null, new System.Drawing.Point(_imagedata.centerx, _imagedata.centery), new MCvScalar(1),
                    out boundRect_med, new MCvScalar(5), new MCvScalar(1), Connectivity.FourConnected, FloodFillType.MaskOnly);
                _imagedata.boundRect_med = boundRect_med;
            }
            _beamparameters.BeamSizeDetect(thresh_med, thresh_e2, _imagedata);

            for (int i = 0; i < _imagedata.TreshE2X.Length; i++)
            {
                _imagedata.TreshE2X[i] = (int)thresh_e2;
                _imagedata.TreshMedX[i] = (int)thresh_med;
                _imagedata.zero[i] = _imagedata.zero_level;

            }

            if (filt)
            {
                _beamparameters.sizex_med_filter = _beamparameters.sizex_med;
                _beamparameters.sizey_med_filter = _beamparameters.sizey_med;

                label3.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizex_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizex_e2 * pixsize) + " um";
                label4.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizey_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizey_e2 * pixsize) + " um";
                dataGridView1[2, 0].Value = Convert.ToString(_beamparameters.sizex_med * pixsize);
                dataGridView1[3, 0].Value = Convert.ToString(_beamparameters.sizey_med * pixsize);
                dataGridView1[2, 1].Value = Convert.ToString(_beamparameters.sizex_e2 * pixsize);
                dataGridView1[3, 1].Value = Convert.ToString(_beamparameters.sizey_e2 * pixsize);
                dataGridView1[6, 0].Value = Convert.ToString(_beamparameters.boundRect_med.Width * pixsize);
                dataGridView1[7, 0].Value = Convert.ToString(_beamparameters.boundRect_med.Height * pixsize);
                dataGridView1[6, 1].Value = Convert.ToString(_beamparameters.boundRect_e2.Width * pixsize);
                dataGridView1[7, 1].Value = Convert.ToString(_beamparameters.boundRect_e2.Height * pixsize);
                xWidthFilteredChart.Series["x filter"].Points.DataBindY(_imagedata.ChartX);
                yWidthFilteredChart.Series["y filter"].Points.DataBindY(_imagedata.ChartY);
                xWidthFilteredChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                xWidthFilteredChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                yWidthFilteredChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                yWidthFilteredChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                xWidthFilteredChart.Series["zero"].Points.DataBindY(_imagedata.zero);
                yWidthFilteredChart.Series["zero"].Points.DataBindY(_imagedata.zero);

                for (int i = 0; i < 2004; i++)
                {
                    _imagedata.ChartXfilter[i] = _imagedata.ChartX[i];
                    _imagedata.ChartYfilter[i] = _imagedata.ChartY[i];
                }
            }
            else
            {
                label1.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizex_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizex_e2 * pixsize) + " um";
                label2.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizey_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizey_e2 * pixsize) + " um";

                xWidthChart.Series["x"].Points.DataBindY(_imagedata.ChartX);
                yWidthChart.Series["y"].Points.DataBindY(_imagedata.ChartY);
                xWidthChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                xWidthChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                yWidthChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                yWidthChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                xWidthChart.Series["zero"].Points.DataBindY(_imagedata.zero);
                yWidthChart.Series["zero"].Points.DataBindY(_imagedata.zero);

                dataGridView1[0, 0].Value = Convert.ToString(_beamparameters.sizex_med * pixsize);
                dataGridView1[1, 0].Value = Convert.ToString(_beamparameters.sizey_med * pixsize);
                dataGridView1[0, 1].Value = Convert.ToString(_beamparameters.sizex_e2 * pixsize);
                dataGridView1[1, 1].Value = Convert.ToString(_beamparameters.sizey_e2 * pixsize);
                dataGridView1[4, 0].Value = Convert.ToString(_beamparameters.boundRect_med.Width * pixsize);
                dataGridView1[5, 0].Value = Convert.ToString(_beamparameters.boundRect_med.Height * pixsize);
                dataGridView1[4, 1].Value = Convert.ToString(_beamparameters.boundRect_e2.Width * pixsize);
                dataGridView1[5, 1].Value = Convert.ToString(_beamparameters.boundRect_e2.Height * pixsize);
            }
            dataGridView1.Refresh();
        }


        private void startButton_Click(object sender, EventArgs e)
        {
            if (_captureInProgress)
            { //stop the capture
                startButton.Text = "start";
                if (_camera != null)
                {
                    _camera.Stop();
                }
            }
            else
            {
                startButton.Text = "stop";
                if (_camera != null)
                {
                    _camera.Start();
                }
            }
            _captureInProgress = !_captureInProgress;
        }

        public void Mess()
        { MessageBox.Show("please, load image"); }


        private void dynamicThermoPictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            mouse_x = e.X;
            mouse_y = e.Y;
        }


        private void dynamicThermoPictureBox_Paint(object sender, PaintEventArgs e)
        {

            label3.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizex_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizex_e2 * pixsize) + " um";
            label4.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizey_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizey_e2 * pixsize) + " um";
            dataGridView1[2, 0].Value = Convert.ToString(_beamparameters.sizex_med * pixsize);
            dataGridView1[3, 0].Value = Convert.ToString(_beamparameters.sizey_med * pixsize);
            dataGridView1[2, 1].Value = Convert.ToString(_beamparameters.sizex_e2 * pixsize);
            dataGridView1[3, 1].Value = Convert.ToString(_beamparameters.sizey_e2 * pixsize);
            dataGridView1[6, 0].Value = Convert.ToString(_beamparameters.boundRect_med.Width * pixsize);
            dataGridView1[7, 0].Value = Convert.ToString(_beamparameters.boundRect_med.Height * pixsize);
            dataGridView1[6, 1].Value = Convert.ToString(_beamparameters.boundRect_e2.Width * pixsize);
            dataGridView1[7, 1].Value = Convert.ToString(_beamparameters.boundRect_e2.Height * pixsize);
            xWidthFilteredChart.Series["x filter"].Points.DataBindY(_imagedata.ChartX);
            yWidthFilteredChart.Series["y filter"].Points.DataBindY(_imagedata.ChartY);
            xWidthFilteredChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            xWidthFilteredChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            yWidthFilteredChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            yWidthFilteredChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            xWidthFilteredChart.Series["zero"].Points.DataBindY(_imagedata.zero);
            yWidthFilteredChart.Series["zero"].Points.DataBindY(_imagedata.zero);

            label1.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizex_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizex_e2 * pixsize) + " um";
            label2.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizey_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizey_e2 * pixsize) + " um";

            xWidthChart.Series["x"].Points.DataBindY(_imagedata.ChartX);
            yWidthChart.Series["y"].Points.DataBindY(_imagedata.ChartY);
            xWidthChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            xWidthChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            yWidthChart.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            yWidthChart.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            xWidthChart.Series["zero"].Points.DataBindY(_imagedata.zero);
            yWidthChart.Series["zero"].Points.DataBindY(_imagedata.zero);

            xWidthChart.Update();
            yWidthChart.Update();
            xWidthFilteredChart.Update();
            yWidthFilteredChart.Update();

            dataGridView1[0, 0].Value = Convert.ToString(_beamparameters.sizex_med * pixsize);
            dataGridView1[1, 0].Value = Convert.ToString(_beamparameters.sizey_med * pixsize);
            dataGridView1[0, 1].Value = Convert.ToString(_beamparameters.sizex_e2 * pixsize);
            dataGridView1[1, 1].Value = Convert.ToString(_beamparameters.sizey_e2 * pixsize);

            dataGridView1.Refresh();

            Graphics gr = e.Graphics;
            drawmark(gr, Color.Red);
        }

        private void manualModeCheckBox_CheckStateChanged(object sender, EventArgs e)
        {
            if (hand_mode) hand_mode = false;
            else hand_mode = true;
        }

        private void loadImage_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0) return;

            openFileDialog1.Filter = "beam picture|*.jpg;*.png;*.gif;*.bmp| All (*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Image t = Image.FromFile(openFileDialog1.FileName);
                staticGrayPictureBox.Image = t;

                Image tf = (Image)t.Clone();
                _imagedata.MakeFalse((Bitmap)tf);
                staticThermoPictureBox.Image = tf;

                _imagedata.ImageFromFile = t;

                mouse_x = (int)(_imagedata.ImageFromFile.Size.Width / 2);
                mouse_y = (int)(_imagedata.ImageFromFile.Size.Height / 2);
                ProcessStaticFrame(t, true);
                ProcessStaticFrame(t, false);

                dynamicGrayPictureBox.Refresh();
                staticThermoPictureBox.Refresh();
                staticGrayPictureBox.Refresh();

                /*ProcessStaticFrame(t, true);
                ProcessStaticFrame(t, false);

                pictureBox3.Refresh();
                pictureBox2.Refresh();*/

                no_image = false;

            }
        }


        private void dynamicGrayPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Graphics gr = e.Graphics;
            drawmark(gr, Color.Red);
            Pen rect = new Pen(Color.White);
            rect.Width = 2;
            //e.Graphics.DrawRectangle(rect, _imagedata.graphstartx, _imagedata.graphstarty, _imagedata.spotsize<<1, _imagedata.spotsize<<1);
        }


        private void staticGrayPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Graphics gr = e.Graphics;
            drawmark(gr, Color.Red);
            Pen rect = new Pen(Color.White);
            rect.Width = 2;

            //e.Graphics.DrawRectangle(rect, _imagedata.graphstartx, _imagedata.graphstarty, _imagedata.spotsize<<1, _imagedata.spotsize<<1);
        }

        private void staticGrayPictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (no_image) { Mess(); return; }

            mouse_x = e.X;
            mouse_y = e.Y;
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
            staticGrayPictureBox.Refresh(); staticThermoPictureBox.Refresh();
            //ProcessStaticFrame(_imagedata.ImageFromFile, true);
            //ProcessStaticFrame(_imagedata.ImageFromFile, false);

        }

        private void staticThermoPictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (no_image) { Mess(); return; }

            mouse_x = e.X;
            mouse_y = e.Y;
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
            staticThermoPictureBox.Refresh(); staticGrayPictureBox.Refresh();
            //ProcessStaticFrame(_imagedata.ImageFromFile, true);
            //ProcessStaticFrame(_imagedata.ImageFromFile, false);
        }

        private void staticThermoPictureBox_Paint(object sender, PaintEventArgs e)
        {
            Graphics gr = e.Graphics;
            drawmark(gr, Color.White);
        }


        private void staticGrayPictureBox_Click(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            staticGrayPictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
            staticThermoPictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
        }

        private void staticThermoPictureBox_Click(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            staticThermoPictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
            staticGrayPictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
        }

        private void pixelPitchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (pixelPitchTextBox.Text != "")
            {
                pixsize = Convert.ToDouble(pixelPitchTextBox.Text);
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
                //ProcessStaticFrame(_imagedata.ImageFromFile, true);
                //ProcessStaticFrame(_imagedata.ImageFromFile, false);
                staticThermoPictureBox.Refresh(); staticGrayPictureBox.Refresh();
            }
        }

        private void pixelPitchTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (char.IsDigit(e.KeyChar) == true || e.KeyChar == (char)Keys.Back || e.KeyChar == ',') return;
            if (e.KeyChar == (char)Keys.Back) return;
            e.Handled = true;
        }

        private void zeroLevelTextBox_TextChanged(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (zeroLevelTextBox.Text != "")
            {
                _imagedata.zero_level = Convert.ToInt16(zeroLevelTextBox.Text);
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);

                staticThermoPictureBox.Refresh(); staticGrayPictureBox.Refresh();
                //ProcessStaticFrame(_imagedata.ImageFromFile, true);
                //ProcessStaticFrame(_imagedata.ImageFromFile, false);
            }
        }

        private void zeroLevelTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (char.IsDigit(e.KeyChar) == true || e.KeyChar == (char)Keys.Back) return;
            e.Handled = true;
        }

        private void dynamicGrayPictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (no_image) { Mess(); return; }

            mouse_x = e.X;
            mouse_y = e.Y;
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
            staticGrayPictureBox.Refresh(); staticThermoPictureBox.Refresh();
            //ProcessStaticFrame(_imagedata.ImageFromFile, true);
            //ProcessStaticFrame(_imagedata.ImageFromFile, false);
        }

        private void radiusTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (char.IsDigit(e.KeyChar) == true || e.KeyChar == (char)Keys.Back || e.KeyChar == (char)Keys.Enter)
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (radiusTextBox.Text != "")
                    {
                        _imagedata.spotsize = Convert.ToInt16(radiusTextBox.Text);
                        if (_imagedata.spotsize > _imagedata.maxspotsize) _imagedata.spotsize = _imagedata.maxspotsize;
                        xWidthChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
                        yWidthChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
                        xWidthFilteredChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
                        yWidthFilteredChart.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;

                        if (tabControl1.SelectedIndex == 1)
                        {
                            ProcessStaticFrame(_imagedata.ImageFromFile, true);
                            ProcessStaticFrame(_imagedata.ImageFromFile, false);
                            staticGrayPictureBox.Refresh(); staticThermoPictureBox.Refresh();
                            //ProcessStaticFrame(_imagedata.ImageFromFile, true);
                            //ProcessStaticFrame(_imagedata.ImageFromFile, false);
                        }
                    }
                }
                return;
            }
            e.Handled = true;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {   // from static to dynamic
                startButton.Visible = true; groupBox1.Visible = false;
            }
            else
            {
                // from dynamic to static
                if (_captureInProgress)
                {
                    startButton.Text = "start";
                    if (_camera != null)
                    {
                        _camera.Stop();
                        _captureInProgress = !_captureInProgress;
                    }
                    no_image = false;
                }

                if (no_image) { startButton.Visible = false; groupBox1.Visible = true; return; }

                staticGrayPictureBox.Image = dynamicGrayPictureBox.Image;
                staticThermoPictureBox.Image = dynamicThermoPictureBox.Image;
                _imagedata.ImageFromFile = dynamicGrayPictureBox.Image;

                mouse_x = (int)(_imagedata.ImageFromFile.Size.Width / 2);
                mouse_y = (int)(_imagedata.ImageFromFile.Size.Height / 2);

                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);

                staticGrayPictureBox.Refresh();
                staticThermoPictureBox.Refresh();

                startButton.Visible = false; groupBox1.Visible = true;
            }
        }

        private void drawmark(Graphics gr, Color cl)
        {
            Pen axis = new Pen(cl);

            axis.Width = 2;
            axis.Color = Color.Red;

            int size = _imagedata.spotsize / 2;

            _imagedata.ChartXstartX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(angleNumericUpDown.Value + 180) / 180 * Math.PI));
            _imagedata.ChartXstartY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(angleNumericUpDown.Value + 180) / 180 * Math.PI));
            _imagedata.ChartXstopX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(angleNumericUpDown.Value) / 180 * Math.PI));
            _imagedata.ChartXstopY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(angleNumericUpDown.Value) / 180 * Math.PI));
            _imagedata.ChartYstartX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(angleNumericUpDown.Value + 270) / 180 * Math.PI));
            _imagedata.ChartYstartY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(angleNumericUpDown.Value + 270) / 180 * Math.PI));
            _imagedata.ChartYstopX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(angleNumericUpDown.Value + 90) / 180 * Math.PI));
            _imagedata.ChartYstopY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(angleNumericUpDown.Value + 90) / 180 * Math.PI));

            Point p1 = new Point(_imagedata.ChartXstartX, _imagedata.ChartXstartY);
            Point p2 = new Point(_imagedata.ChartXstopX, _imagedata.ChartXstopY);
            Point p3 = new Point(_imagedata.ChartYstartX, _imagedata.ChartYstartY);
            Point p4 = new Point(_imagedata.ChartYstopX, _imagedata.ChartYstopY);

            gr.DrawLine(axis, p1, p2);
            axis.Color = Color.Green;
            gr.DrawLine(axis, p3, p4);
            axis.Color = Color.White;
            gr.DrawEllipse(axis, _imagedata.centerx - _imagedata.spotsize / 2, _imagedata.centery - _imagedata.spotsize / 2, _imagedata.spotsize, _imagedata.spotsize);
            axis.Color = Color.Orange;
            gr.DrawRectangle(axis, _imagedata.boundRect_e2);
        }


        private void angleNumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (angleNumericUpDown.Value == 360)
                angleNumericUpDown.Value = 0;
            if (angleNumericUpDown.Value == -1)
                angleNumericUpDown.Value = 359;

            if (tabControl1.SelectedIndex == 1)
            {
                // if static selected
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
                staticThermoPictureBox.Refresh(); staticGrayPictureBox.Refresh();
                //ProcessStaticFrame(_imagedata.ImageFromFile, true);
                //ProcessStaticFrame(_imagedata.ImageFromFile, false);
            }
        }

        private void exportToExcel_Click(object sender, EventArgs e)
        {
            try
            {
                excelapp = new Excel.Application();
            }
            catch (COMException)
            {
                return;
            }
            excelapp.Visible = true;

            excelapp.SheetsInNewWorkbook = 3;
            excelapp.Workbooks.Add(Type.Missing);

            excelappworkbooks = excelapp.Workbooks;
            excelappworkbook = excelappworkbooks[1];

            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelworksheet.Activate();

            // data x , data y ,xfilter,yfilter - to exel cells

            for (int i = 2; i < _imagedata.spotsize + 4; i++)
            {

                excelcells = (Excel.Range)excelworksheet.Cells[i, 1];
                excelcells.Value2 = _imagedata.ChartX[i];

            }

            for (int i = 2; i < _imagedata.spotsize + 4; i++)
            {

                excelcells = (Excel.Range)excelworksheet.Cells[i, 2];
                excelcells.Value2 = _imagedata.ChartY[i];

            }

            for (int i = 2; i < _imagedata.spotsize + 4; i++)
            {

                excelcells = (Excel.Range)excelworksheet.Cells[i, 4];
                excelcells.Value2 = _imagedata.ChartXfilter[i];

            }

            for (int i = 2; i < _imagedata.spotsize + 4; i++)
            {

                excelcells = (Excel.Range)excelworksheet.Cells[i, 5];
                excelcells.Value2 = _imagedata.ChartYfilter[i];

            }

            excelcells = excelworksheet.get_Range("A1", "A1");

            excelcells.Value2 = " X ";
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("B1", "B1");

            excelcells.Value2 = " Y ";
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("D1", "D1");

            excelcells.Value2 = " X filter ";
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("E1", "E1");

            excelcells.Value2 = " Y filter";
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("W11", "W11");

            excelcells.Value2 = "pixel pitch(um):";
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("Y11", "Y11");

            excelcells.Value2 = pixelPitchTextBox.Text;
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("W13", "W13");

            excelcells.Value2 = "zero level:";
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("Y13", "Y13");

            excelcells.Value2 = zeroLevelTextBox.Text;
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("H3", "H3");

            excelcells.Value2 = label1.Text;
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("P3", "P3");

            excelcells.Value2 = label2.Text;
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("H26", "H26");

            excelcells.Value2 = label3.Text;
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("P26", "P26");

            excelcells.Value2 = label4.Text;
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            // chart export
            Excel.ChartObjects chartsobjrcts =
            (Excel.ChartObjects)excelworksheet.ChartObjects(Type.Missing);
            Excel.ChartObject chartsobjrct = chartsobjrcts.Add(300, 50, 300, 300);
            chartsobjrct.Chart.ChartWizard(excelworksheet.get_Range("A1", "A" + Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);


            Excel.ChartObject chartsobjrct1 = chartsobjrcts.Add(700, 50, 300, 300);
            chartsobjrct1.Chart.ChartWizard(excelworksheet.get_Range("B1", "B" + Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);


            Excel.ChartObject chartsobjrct2 = chartsobjrcts.Add(300, 400, 300, 300);
            chartsobjrct2.Chart.ChartWizard(excelworksheet.get_Range("D1", "D" + Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);

            Excel.ChartObject chartsobjrct3 = chartsobjrcts.Add(700, 400, 300, 300);
            chartsobjrct3.Chart.ChartWizard(excelworksheet.get_Range("E1", "E" + Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);
        }


        private void divergenceCheckBox_CheckStateChanged(object sender, EventArgs e)
        {
            if (divergenceCheckBox.Checked)
            {
                wavelengthLabel.Visible = true; waveLengthListBox.Visible = true;
            }
            else
            {
                wavelengthLabel.Visible = false; waveLengthListBox.Visible = false;
            }
        }

        private void wavelengthListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1[0, 2].Value = Convert.ToString(((_beamparameters.sizex_med * pixsize / 1000) / _imagedata.wavelength[waveLengthListBox.SelectedIndex]) * 180 / Math.PI);
            dataGridView1[1, 2].Value = Convert.ToString(((_beamparameters.sizey_med * pixsize / 1000) / _imagedata.wavelength[waveLengthListBox.SelectedIndex]) * 180 / Math.PI);
            dataGridView1[2, 2].Value = Convert.ToString(((_beamparameters.sizex_med_filter * pixsize / 1000) / _imagedata.wavelength[waveLengthListBox.SelectedIndex]) * 180 / Math.PI);
            dataGridView1[3, 2].Value = Convert.ToString(((_beamparameters.sizey_med_filter * pixsize / 1000) / _imagedata.wavelength[waveLengthListBox.SelectedIndex]) * 180 / Math.PI);

            dataGridView1.Refresh();
        }

        private void _3dButton_Click(object sender, EventArgs e)
        {
            int n = 400, m = 400, i, j, nlev = 20;
            float[,] zmat = new float[n, m];
            float[] xray = new float[n];
            float[] yray = new float[m];
            float[] zlev = new float[nlev];

            double x, y, step;
            double stepx = 400.0 / (n - 1);
            double stepy = 400.0 / (m - 1);

            for (i = 0; i < n; i++)
            {
                x = i * stepx;
                xray[i] = (float)x;
                for (j = 0; j < m; j++)
                {
                    y = j * stepy;
                    yray[j] = (float)y;
                    zmat[i, j] = _imagedata.Graph3d[i, j];
                }
            }

            dislin.scrmod("revers");
            dislin.metafl("cons");
            dislin.setpag("da4p");
            dislin.disini();
            dislin.pagera();
            dislin.hwfont();

            dislin.axspos(200, 2600);
            dislin.axslen(1800, 1800);

            dislin.name("X-axis", "x");
            dislin.name("Y-axis", "y");
            dislin.name("Z-axis", "z");

            dislin.titlin("       ", 1);
            dislin.titlin("       ", 3);

            /* dislin.graf3d (0.0f, 360.0f, 0.0f, 90.0f,
                            0.0f, 360.0f, 0.0f, 90.0f,
                            -5.0f, 5.0f, -5.0f, 5.0f);*/
            dislin.graf3d(0.0f, 400.0f, 0.0f, 90.0f,
                          0.0f, 400.0f, 0.0f, 90.0f,
                          0f, 256f, 0f, 256f);
            dislin.height(50);
            dislin.title();

            dislin.grfini(-1.0f, -1.0f, -1.0f, 1.0f, -1.0f, -1.0f,
                           1.0f, 1.0f, -1.0f);
            dislin.nograf();
            dislin.graf(0.0f, 400.0f, 0.0f, 90.0f, 0.0f, 400.0f, 0.0f, 90.0f);
            step = 4.0 / nlev;
            for (i = 0; i < nlev; i++)
                zlev[i] = (float)(-2.0 + i * step);

            // dislin.conshd (xray, n, yray, n, zmat, zlev, nlev);
            dislin.box2d();
            dislin.reset("nograf");
            dislin.grffin();

            dislin.shdmod("smooth", "surface");
            dislin.surshd(xray, n, yray, m, zmat);
            dislin.disfin();
        }

        private void dynamicTabPage_Click(object sender, EventArgs e)
        {

        }

        private void staticTabPage_Click(object sender, EventArgs e)
        {

        }
    }
}
