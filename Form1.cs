

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;



using Emgu.CV;
using Emgu.CV.UI;
using Emgu.Util;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Excel = Microsoft.Office.Interop.Excel;


namespace LaserBeamMeasurement
{


    public partial class Form1 : Form
    {

        bool hand_mode = false;
        bool d3_mode = false;

        int mouse_x = 0;
        int mouse_y = 0;

        bool no_image = true;
       
        double pixsize = 2.2;

        ImageData _imagedata = new ImageData();
        BeamParameters _beamparameters = new BeamParameters();

        private Capture _capture = null;

        private bool _captureInProgress;
        private Excel.Application excelapp;
        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;
        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcells;

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

            chart1.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
            chart2.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
            chart3.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
            chart4.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;

            tabControl1.SelectedIndex = 1;

            CvInvoke.UseOpenCL = false;

            try
            {
                _capture = new Capture();

                _capture.ImageGrabbed += ProcessFrame;
            

            }
            catch (NullReferenceException excpt)
            {
                MessageBox.Show(excpt.Message);
            }
        }

        public void ProcessFrame(object sender, EventArgs arg) // will be simplified
        {

            //get frame from the Camera

            Mat frame = new Mat();
            Mat grayFrame = new Mat();

            double thresh_med;
            double thresh_e2;

            System.Drawing.Image OrgImage;
            System.Drawing.Image OrgImage1;

            _capture.Retrieve(frame);

            CvInvoke.CvtColor(frame, grayFrame, ColorConversion.Bgr2Gray);
            Image<Rgb, Byte> tothermo = grayFrame.ToImage<Rgb, Byte>(); // original
            Image<Gray, Byte> tothermo1 = grayFrame.ToImage<Gray, Byte>();

            OrgImage = tothermo.ToBitmap();
            OrgImage1 = tothermo.ToBitmap();

            pictureBox4.Image = OrgImage1;

            _imagedata.MakeFalse((Bitmap)OrgImage);
            pictureBox1.Image = OrgImage;

            pictureBox1.Refresh();
            pictureBox4.Refresh();
          


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
            pictureBox1.Refresh();
            pictureBox4.Refresh();

            //_imagedata.GraphFill(tothermo1);
             _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartXstartX, _imagedata.ChartXstartY, _imagedata.ChartXstopX, _imagedata.ChartXstopY, true);
             _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartYstartX, _imagedata.ChartYstartY, _imagedata.ChartYstopX, _imagedata.ChartYstopY, false);
             _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartXstartX, _imagedata.ChartXstartY, _imagedata.ChartXstopX, _imagedata.ChartXstopY, true);
             _imagedata.GraphFillRotate(tothermo1, _imagedata.ChartYstartX, _imagedata.ChartYstartY, _imagedata.ChartYstopX, _imagedata.ChartYstopY, false);

            // thresh calculate
            if (hand_mode)
            {
                thresh_med = grayFrame.Bitmap.GetPixel(_imagedata.centerx, _imagedata.centery).R / 2;
                thresh_e2 = grayFrame.Bitmap.GetPixel(_imagedata.centerx, _imagedata.centery).R / Math.Exp(2);
            }
            else
            {
                thresh_med = maxVal[0] / 2;
                thresh_e2 = maxVal[0] / Math.Exp(2);
            }

            _beamparameters.BeamSizeDetect(thresh_med, thresh_e2, _imagedata);



        }

        public void ProcessStaticFrame(Image ImFromFile, bool filt)
        {
            double thresh_med;
            double thresh_e2;
            // Converting the master image to a bitmap
            Bitmap masterImage = (Bitmap)ImFromFile;

            // Normalizing it to grayscale
            Image<Gray, Byte> grayFrame = new Image<Gray, Byte>(masterImage);
            Image<Gray, Byte> grayFramefilt = new Image<Gray, Byte>(masterImage);

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
            double rotation_angle = (int)numericUpDown1.Value * Math.PI/180;

            double rotation_sin = Math.Sin(rotation_angle);
            double rotation_cos = Math.Cos(rotation_angle);

            for (int i = - width; i < width; i++)
            {
                for (int j = - height; j < height; j++)

                {
                    _imagedata.Graph3d[k, m] = grayFrame.Data[(int)(i * rotation_sin + j * rotation_cos) + _imagedata.centery, (int)(i* rotation_cos - j* rotation_sin) + _imagedata.centerx, 0];   //grayFrame.Bitmap.GetPixel(i, j).R;
                    m++;
                }

                k++; m = 0;

            }


            /*

            for (int i = _imagedata.centerx - width; i < _imagedata.centerx + width; i++)
            {
                for (int j = _imagedata.centery - height; j < _imagedata.centery + height; j++)

                {
                    _imagedata.Graph3d[k, m] = grayFrame.Data[j, i, 0];   //grayFrame.Bitmap.GetPixel(i, j).R;
                    m++;
                }

                k++; m = 0;

            } */



            _imagedata.GraphFillRotate(grayFrame, _imagedata.ChartXstartX, _imagedata.ChartXstartY, _imagedata.ChartXstopX, _imagedata.ChartXstopY,true);
             _imagedata.GraphFillRotate(grayFrame, _imagedata.ChartYstartX, _imagedata.ChartYstartY, _imagedata.ChartYstopX, _imagedata.ChartYstopY,false);

       
              _imagedata.zero_level = Convert.ToInt32(textBox2.Text);
   

            // thresh calculate
            if (hand_mode)
            {
                thresh_med = (grayFrame.Bitmap.GetPixel(_imagedata.centerx, _imagedata.centery).R- _imagedata.zero_level) / 2 + _imagedata.zero_level;
                thresh_e2 = (grayFrame.Bitmap.GetPixel(_imagedata.centerx, _imagedata.centery).R- _imagedata.zero_level) / Math.Exp(2) + _imagedata.zero_level;
            }
            else
            {
                thresh_med = (maxVal[0]- _imagedata.zero_level)/ 2 + _imagedata.zero_level;
                thresh_e2 = (maxVal[0]- _imagedata.zero_level) / Math.Exp(2) + _imagedata.zero_level;
            }
     
            _beamparameters.BeamSizeDetect(thresh_med, thresh_e2, _imagedata);

            for (int i=0;i<_imagedata.TreshE2X.Length;i++)
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
                chart3.Series["x filter"].Points.DataBindY(_imagedata.ChartX);
                chart4.Series["y filter"].Points.DataBindY(_imagedata.ChartY);
                chart3.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                chart3.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                chart4.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                chart4.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                chart3.Series["zero"].Points.DataBindY(_imagedata.zero);
                chart4.Series["zero"].Points.DataBindY(_imagedata.zero);

                for(int i=0;i< 2004;i++)
                {
                    _imagedata.ChartXfilter[i] = _imagedata.ChartX[i];
                    _imagedata.ChartYfilter[i] = _imagedata.ChartY[i];
                }
            }
            else
            {
                label1.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizex_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizex_e2 * pixsize) + " um";
                label2.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizey_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizey_e2 * pixsize) + " um";

                chart1.Series["x"].Points.DataBindY(_imagedata.ChartX);
                chart2.Series["y"].Points.DataBindY(_imagedata.ChartY);
                chart1.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                chart1.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                chart2.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
                chart2.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
                chart1.Series["zero"].Points.DataBindY(_imagedata.zero);
                chart2.Series["zero"].Points.DataBindY(_imagedata.zero);

                dataGridView1[0, 0].Value = Convert.ToString(_beamparameters.sizex_med * pixsize);
                dataGridView1[1, 0].Value = Convert.ToString(_beamparameters.sizey_med * pixsize);
                dataGridView1[0, 1].Value = Convert.ToString(_beamparameters.sizex_e2 * pixsize);
                dataGridView1[1, 1].Value = Convert.ToString(_beamparameters.sizey_e2 * pixsize);
            }


             pictureBox3.Refresh(); pictureBox2.Refresh();

        }
 
        private void button1_Click(object sender, EventArgs e)
        {


            if (_capture != null)
            {
                if (_captureInProgress)
                {  //stop the capture
                    button1.Text = "start";
                    _capture.Pause();

                }
                else
                {
                    //start the capture
                    button1.Text = "stop";
                    _capture.Start();

                }

                _captureInProgress = !_captureInProgress;
            }
        }

        public void Mess()
        { MessageBox.Show("please, load image"); }


        private void pictureBox1_MouseClick(object sender, MouseEventArgs e)
        {
            mouse_x = e.X;
            mouse_y = e.Y;
        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {

            label3.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizex_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizex_e2 * pixsize) + " um";
            label4.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizey_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizey_e2 * pixsize) + " um";
            dataGridView1[2, 0].Value = Convert.ToString(_beamparameters.sizex_med * pixsize);
            dataGridView1[3, 0].Value = Convert.ToString(_beamparameters.sizey_med * pixsize);
            dataGridView1[2, 1].Value = Convert.ToString(_beamparameters.sizex_e2 * pixsize);
            dataGridView1[3, 1].Value = Convert.ToString(_beamparameters.sizey_e2 * pixsize);
            chart3.Series["x filter"].Points.DataBindY(_imagedata.ChartX);
            chart4.Series["y filter"].Points.DataBindY(_imagedata.ChartY);
            chart3.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            chart3.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            chart4.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            chart4.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            chart3.Series["zero"].Points.DataBindY(_imagedata.zero);
            chart4.Series["zero"].Points.DataBindY(_imagedata.zero);

            label1.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizex_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizex_e2 * pixsize) + " um";
            label2.Text = "FWHM:  " + Convert.ToString(_beamparameters.sizey_med * pixsize) + " um           1/e^2:  " + Convert.ToString(_beamparameters.sizey_e2 * pixsize) + " um";

            chart1.Series["x"].Points.DataBindY(_imagedata.ChartX);
            chart2.Series["y"].Points.DataBindY(_imagedata.ChartY);
            chart1.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            chart1.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            chart2.Series["fwhm"].Points.DataBindY(_imagedata.TreshMedX);
            chart2.Series["1/e^2"].Points.DataBindY(_imagedata.TreshE2X);
            chart1.Series["zero"].Points.DataBindY(_imagedata.zero);
            chart2.Series["zero"].Points.DataBindY(_imagedata.zero);

       
            dataGridView1[0, 0].Value = Convert.ToString(_beamparameters.sizex_med * pixsize);
            dataGridView1[1, 0].Value = Convert.ToString(_beamparameters.sizey_med * pixsize);
            dataGridView1[0, 1].Value = Convert.ToString(_beamparameters.sizex_e2 * pixsize);
            dataGridView1[1, 1].Value = Convert.ToString(_beamparameters.sizey_e2 * pixsize);




            Graphics gr = e.Graphics;
            drawmark(gr, Color.Red);

        }


        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (hand_mode) hand_mode = false;
            else hand_mode = true;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0) return;

            openFileDialog1.Filter = "beam picture|*.jpg;*.png;*.gif;*.bmp| All (*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Image t = Image.FromFile(openFileDialog1.FileName);
                pictureBox2.Image = t;

                Image tf = (Image)t.Clone();
                _imagedata.MakeFalse((Bitmap)tf);
                pictureBox3.Image = tf;

                _imagedata.ImageFromFile = t;
             
                mouse_x = (int)(_imagedata.ImageFromFile.Size.Width/2);
                mouse_y = (int)(_imagedata.ImageFromFile.Size.Height/2);
                ProcessStaticFrame(t,true);
                ProcessStaticFrame(t, false);
                          
                pictureBox4.Refresh();
                pictureBox3.Refresh();

                ProcessStaticFrame(t, true);
                ProcessStaticFrame(t, false);

                no_image = false;

            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {   

        }


        private void pictureBox4_Paint(object sender, PaintEventArgs e)
        {

                  
            Graphics gr = e.Graphics;
            drawmark(gr, Color.Red);
            Pen rect = new Pen(Color.White);
            rect.Width = 2;
            //e.Graphics.DrawRectangle(rect, _imagedata.graphstartx, _imagedata.graphstarty, _imagedata.spotsize<<1, _imagedata.spotsize<<1);
        }

        private void pictureBox2_Paint(object sender, PaintEventArgs e)
        {
            Graphics gr = e.Graphics;
            drawmark(gr,Color.Red);
            Pen rect = new Pen(Color.White);
            rect.Width = 2;

            //e.Graphics.DrawRectangle(rect, _imagedata.graphstartx, _imagedata.graphstarty, _imagedata.spotsize<<1, _imagedata.spotsize<<1);
        }

        private void pictureBox2_MouseClick(object sender, MouseEventArgs e)
        {
            if (no_image) { Mess(); return; }
                
            mouse_x = e.X;
            mouse_y = e.Y;
            ProcessStaticFrame(_imagedata.ImageFromFile,true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
            pictureBox2.Refresh(); pictureBox3.Refresh();
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
        }

        private void pictureBox3_MouseClick(object sender, MouseEventArgs e)
        {
            if (no_image) { Mess(); return; }

            mouse_x = e.X;
            mouse_y = e.Y;
            ProcessStaticFrame(_imagedata.ImageFromFile,true);
            ProcessStaticFrame(_imagedata.ImageFromFile,false);
            pictureBox3.Refresh(); pictureBox2.Refresh();
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
        }

        private void pictureBox3_Paint(object sender, PaintEventArgs e)
        {
            Graphics gr = e.Graphics;
            drawmark(gr, Color.White);
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            pictureBox2.SizeMode = PictureBoxSizeMode.AutoSize;
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            pictureBox3.SizeMode = PictureBoxSizeMode.AutoSize;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (textBox1.Text != "")
            {
                pixsize = Convert.ToDouble(textBox1.Text);
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (char.IsDigit(e.KeyChar) == true || e.KeyChar == (char)Keys.Back ||  e.KeyChar == ',') return;
            if (e.KeyChar == (char)Keys.Back) return;
            e.Handled = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (textBox2.Text != "")
            {
                _imagedata.zero_level = Convert.ToInt16(textBox2.Text);
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (char.IsDigit(e.KeyChar) == true || e.KeyChar == (char)Keys.Back) return;
            e.Handled = true;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
        }

        private void pictureBox4_MouseClick(object sender, MouseEventArgs e)
        {
            if (no_image) { Mess(); return; }

            mouse_x = e.X;
            mouse_y = e.Y;
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false); pictureBox2.Refresh(); pictureBox3.Refresh();
            ProcessStaticFrame(_imagedata.ImageFromFile, true);
            ProcessStaticFrame(_imagedata.ImageFromFile, false);
        }

         private void textBox3_TextChanged(object sender, EventArgs e)
        { 
   
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (char.IsDigit(e.KeyChar) == true || e.KeyChar == (char)Keys.Back || e.KeyChar == (char)Keys.Enter)
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (textBox3.Text != "")
                    {
                        _imagedata.spotsize = Convert.ToInt16(textBox3.Text);
                        if (_imagedata.spotsize > _imagedata.maxspotsize) _imagedata.spotsize = _imagedata.maxspotsize;
                        chart1.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
                        chart2.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
                        chart3.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;
                        chart4.ChartAreas[0].AxisX.Maximum = _imagedata.spotsize + 4;


                        if (tabControl1.SelectedIndex == 1)
                        {

                            ProcessStaticFrame(_imagedata.ImageFromFile, true);
                            ProcessStaticFrame(_imagedata.ImageFromFile, false);
                            pictureBox2.Refresh(); pictureBox3.Refresh();
                            ProcessStaticFrame(_imagedata.ImageFromFile, true);
                            ProcessStaticFrame(_imagedata.ImageFromFile, false);
                        }
                    }
                }
                    return;
            }
            e.Handled = true;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0) {   // from static to dinamic
             button1.Visible = true; groupBox1.Visible = false;
              

            }

            else if(tabControl1.SelectedIndex == 2) // 3d mode
            {
              
              
            }

            else {



                if (_captureInProgress)
                { 
                    button1.Text = "start";
                    _capture.Pause();
                    _captureInProgress = !_captureInProgress;
                    no_image = false;

                }

                if (no_image) { button1.Visible = false; groupBox1.Visible = true; return; }

                pictureBox2.Image= pictureBox4.Image;
                pictureBox3.Image = pictureBox1.Image;
                _imagedata.ImageFromFile = pictureBox4.Image;

                mouse_x = (int)(_imagedata.ImageFromFile.Size.Width / 2);
                mouse_y = (int)(_imagedata.ImageFromFile.Size.Height / 2);

                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);

                pictureBox2.Refresh();
                pictureBox3.Refresh();


                button1.Visible = false; groupBox1.Visible = true; }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void drawmark (Graphics gr, Color cl)
        {
            Pen axis = new Pen(cl);
  
            axis.Width = 2;
            axis.Color = Color.Red;

            int size = _imagedata.spotsize / 2;

            _imagedata.ChartXstartX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(numericUpDown1.Value + 180) / 180 * Math.PI));
            _imagedata.ChartXstartY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(numericUpDown1.Value + 180) / 180 * Math.PI));
            _imagedata.ChartXstopX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(numericUpDown1.Value ) / 180 * Math.PI));
            _imagedata.ChartXstopY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(numericUpDown1.Value ) / 180 * Math.PI));
            _imagedata.ChartYstartX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(numericUpDown1.Value + 270) / 180 * Math.PI));
            _imagedata.ChartYstartY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(numericUpDown1.Value + 270) / 180 * Math.PI));
            _imagedata.ChartYstopX = _imagedata.centerx + Convert.ToInt32(size * Math.Cos(Convert.ToDouble(numericUpDown1.Value + 90) / 180 * Math.PI));
            _imagedata.ChartYstopY = _imagedata.centery + Convert.ToInt32(size * Math.Sin(Convert.ToDouble(numericUpDown1.Value + 90) / 180 * Math.PI));

            Point p1 = new Point(_imagedata.ChartXstartX, _imagedata.ChartXstartY);
            Point p2 = new Point(_imagedata.ChartXstopX,_imagedata.ChartXstopY);
            Point p3 = new Point(_imagedata.ChartYstartX, _imagedata.ChartYstartY);
            Point p4 = new Point(_imagedata.ChartYstopX, _imagedata.ChartYstopY);

           
            gr.DrawLine(axis, p1, p2);
            axis.Color = Color.Green;
            gr.DrawLine(axis, p3, p4);
            axis.Color = Color.White;
            gr.DrawEllipse(axis, _imagedata.centerx- _imagedata.spotsize/2, _imagedata.centery- _imagedata.spotsize / 2, _imagedata.spotsize, _imagedata.spotsize);

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (no_image) { Mess(); return; }

            if (numericUpDown1.Value == 360)
                numericUpDown1.Value = 0;
            if (numericUpDown1.Value == -1)
                numericUpDown1.Value = 359;

            if (tabControl1.SelectedIndex == 1)
            {
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
                ProcessStaticFrame(_imagedata.ImageFromFile, true);
                ProcessStaticFrame(_imagedata.ImageFromFile, false);
            }
     
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
    
             excelapp = new Excel.Application();
             excelapp.Visible = true;

             excelapp.SheetsInNewWorkbook = 3;
             excelapp.Workbooks.Add(Type.Missing);
          
            
            excelappworkbooks = excelapp.Workbooks;
            excelappworkbook=excelappworkbooks[1];
     
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelworksheet.Activate();

            // data x , data y ,xfilter,yfilter - to exel cells
 
            for (int i = 2; i < _imagedata.spotsize + 4 ; i++)
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

            excelcells.Value2 = textBox1.Text;
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;

            excelcells = excelworksheet.get_Range("W13", "W13");

            excelcells.Value2 = "zero level:";
            excelcells.Font.Size = 12;
            excelcells.Font.Italic = true;
            excelcells.Font.Bold = true;


            excelcells = excelworksheet.get_Range("Y13", "Y13");

            excelcells.Value2 = textBox2.Text;
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
            chartsobjrct.Chart.ChartWizard(excelworksheet.get_Range("A1", "A"+ Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);


            Excel.ChartObject chartsobjrct1 = chartsobjrcts.Add(700, 50, 300, 300);
            chartsobjrct1.Chart.ChartWizard(excelworksheet.get_Range("B1", "B"+ Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);


            Excel.ChartObject chartsobjrct2 = chartsobjrcts.Add(300, 400, 300, 300);
            chartsobjrct2.Chart.ChartWizard(excelworksheet.get_Range("D1", "D"+Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);

            Excel.ChartObject chartsobjrct3 = chartsobjrcts.Add(700, 400, 300, 300);
            chartsobjrct3.Chart.ChartWizard(excelworksheet.get_Range("E1", "E"+ Convert.ToString(_imagedata.spotsize + 20)),
            Excel.XlChartType.xlLine, 2, Excel.XlRowCol.xlColumns, Type.Missing,
              1, true, " ", Type.Missing);
     


           



        }

      

 
        private void checkBox2_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                label9.Visible = true; listBox1.Visible = true;
            }
            else
            {
                label9.Visible = false; listBox1.Visible = false;
            }
            
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1[0, 2].Value = Convert.ToString(((_beamparameters.sizex_med * pixsize / 1000) / _imagedata.wavelength[listBox1.SelectedIndex])*180/Math.PI);
            dataGridView1[1, 2].Value = Convert.ToString(((_beamparameters.sizey_med * pixsize / 1000) / _imagedata.wavelength[listBox1.SelectedIndex]) * 180 / Math.PI);
            dataGridView1[2, 2].Value = Convert.ToString(((_beamparameters.sizex_med_filter * pixsize / 1000) / _imagedata.wavelength[listBox1.SelectedIndex]) * 180 / Math.PI);
            dataGridView1[3, 2].Value = Convert.ToString(((_beamparameters.sizey_med_filter * pixsize / 1000) / _imagedata.wavelength[listBox1.SelectedIndex]) * 180 / Math.PI);


        }

        private void button2_Click_1(object sender, EventArgs e)
        {


     int n = 400, m = 400, i, j, nlev = 20;
     float [,] zmat = new float [n,m];
     float [] xray = new float [n];
     float [] yray = new float [m];
     float [] zlev = new float [nlev];

     double x, y, step;
     double stepx = 400.0 / (n-1);
     double stepy = 400.0 / (m-1);

     
     for (i = 0; i < n; i++) {
       x = i * stepx;
       xray[i] = (float) x;
       for (j = 0; j < m; j++) {
         y = j * stepy;
         yray[j] = (float) y;
         zmat[i, j] =  _imagedata.Graph3d[i,j];
       }
     }
       
     dislin.scrmod ("revers");
     dislin.metafl ("cons");
     dislin.setpag ("da4p");
     dislin.disini ();
     dislin.pagera ();
     dislin.hwfont ();

     dislin.axspos (200, 2600);
     dislin.axslen (1800, 1800);
         
     dislin.name   ("X-axis", "x");
     dislin.name   ("Y-axis",  "y");
     dislin.name   ("Z-axis",  "z");

     dislin.titlin ("       ", 1);
     dislin.titlin ("       ", 3);

    /* dislin.graf3d (0.0f, 360.0f, 0.0f, 90.0f,
                    0.0f, 360.0f, 0.0f, 90.0f,
                    -5.0f, 5.0f, -5.0f, 5.0f);*/
            dislin.graf3d(0.0f, 400.0f, 0.0f, 90.0f,
                          0.0f, 400.0f, 0.0f, 90.0f,
                          0f, 256f, 0f, 256f);
            dislin.height (50);
     dislin.title  ();
 
     dislin.grfini (-1.0f, -1.0f, -1.0f, 1.0f, -1.0f, -1.0f,
                    1.0f, 1.0f, -1.0f);
     dislin.nograf ();
     dislin.graf (0.0f, 400.0f, 0.0f, 90.0f, 0.0f, 400.0f, 0.0f, 90.0f);
     step = 4.0 / nlev;
     for (i = 0; i < nlev; i++)
       zlev[i] = (float) (-2.0 + i * step); 

    // dislin.conshd (xray, n, yray, n, zmat, zlev, nlev);
     dislin.box2d ();
     dislin.reset ("nograf");
     dislin.grffin ();

     dislin.shdmod  ("smooth", "surface");
     dislin.surshd (xray, n, yray, m, zmat);
     dislin.disfin ();

        }
    }

   
    }








