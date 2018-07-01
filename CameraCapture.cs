//----------------------------------------------------------------------------
//  Copyright (C) 2004-2018 by EMGU Corporation. All rights reserved.       
//----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using System.IO.Ports;
using System.Management;
using System.Collections.Generic;
using System.Linq;

namespace CameraCapture
{
   public partial class CameraCapture : Form
   {
      private VideoCapture _capture = null;
      private bool _captureInProgress;
      private Mat _frame;
      private Mat _grayFrame;
      private Mat _smallGrayFrame;
      private bool detection;
      private bool serial;
  

      public CameraCapture()
      {
         InitializeComponent();
         CvInvoke.UseOpenCL = false;
         try
         {
            _capture = new VideoCapture();
            _capture.ImageGrabbed += ProcessFrame;
         }
         catch (NullReferenceException excpt)
         {
            MessageBox.Show(excpt.Message);
         }
         _frame = new Mat();
         _grayFrame = new Mat();
         _smallGrayFrame = new Mat();
         serialPort1.DataReceived += new SerialDataReceivedEventHandler(this.port_DataReceived);

        }

        private void ProcessFrame(object sender, EventArgs arg)
      {
            
            if (_capture != null && _capture.Ptr != IntPtr.Zero)
            {
                String rect = "";
                _capture.Retrieve(_frame, 0);

                CvInvoke.CvtColor(_frame, _grayFrame, ColorConversion.Bgr2Gray);

                CvInvoke.PyrDown(_grayFrame, _smallGrayFrame);

                long detectionTime;
                List<Rectangle> faces = new List<Rectangle>();
                List<Rectangle> eyes = new List<Rectangle>();

                if (detection) { 
                DetectFace.Detect(
                  _frame, "haarcascade_frontalface_default.xml",
                  faces, eyes,
                  out detectionTime);
                
                foreach (Rectangle face in faces) { 
                    rect = face.ToString();
                    CvInvoke.Rectangle(_grayFrame, face, new Bgr(Color.Red).MCvScalar, 2);
                    CvInvoke.Rectangle(_frame, face, new Bgr(Color.Red).MCvScalar, 2);

                }
                }


                captureImageBox.Image = _frame;
            grayscaleImageBox.Image = _grayFrame;
            AppendTextBox(rect, textBox1);
            }
      }

        public void AppendTextBox(string value, TextBox textbox)
        {
            if (InvokeRequired)
            {
                this.BeginInvoke(new Action<string, TextBox>(AppendTextBox),  value  , textbox);
                return;
            }
            textbox.Text = value;
        }

        private void captureButtonClick(object sender, EventArgs e)
      {
         if (_capture != null)
         {
            if (_captureInProgress)
            {  //stop the capture
               captureButton.Text = "Start Capture";
               _capture.Pause();
            }
            else
            {
               //start the capture
               captureButton.Text = "Stop";
               _capture.Start();
            }

            _captureInProgress = !_captureInProgress;
         }
      }

      private void ReleaseData()
      {
         if (_capture != null)
            _capture.Dispose();
      }

      private void FlipHorizontalButtonClick(object sender, EventArgs e)
      {
         if (_capture != null) _capture.FlipHorizontal = !_capture.FlipHorizontal;
      }

      private void FlipVerticalButtonClick(object sender, EventArgs e)
      {
         if (_capture != null) _capture.FlipVertical = !_capture.FlipVertical;
      }

        private void serialCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            serial = serialCheckbox.Checked;
            if (serial) {
                textBox2.AppendText("Started serial coms..." + Environment.NewLine);

                try
                {
                    serialPort1.PortName = comboBox1.Items[comboBox1.SelectedIndex].ToString().Split(' ')[0];

                    serialPort1.Open();
                    if (serialPort1.IsOpen)
                    {
                        comboBox1.Enabled = false;

                    }

                }
                catch(Exception ex)
                {
                    textBox2.AppendText(ex.StackTrace + Environment.NewLine);


                }
            }
            else
            {
                textBox2.AppendText("Ended serial coms..." + Environment.NewLine);

                try
                {
                    serialPort1.Close();
                    comboBox1.Enabled = true;
                }
                catch (Exception ex)
                {
                    textBox2.AppendText(ex.StackTrace + Environment.NewLine);

                }
            }

            

           
            
            }
        private void detectCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            detection = detectCheckBox.Checked;
            if (serial)
            {
                textBox2.AppendText("Started face tracking..." + Environment.NewLine);
            }
            else
            {
                textBox2.AppendText("Ended face tracking..." + Environment.NewLine);

            }
        }

        private void updateCOMports()
        {

            if (comboBox1.Items.Count == SerialPort.GetPortNames().Length) {
                return;
            }
            comboBox1.Items.Clear();

            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
            {
                var portnames = SerialPort.GetPortNames();
                var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());

                var portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();

                foreach (string s in portList)
                {
                    comboBox1.Items.Add(s);
                }
            }
        }

    

        private void timerCOM_Tick(object sender, EventArgs e)
        {
            updateCOMports();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (serialPort1.IsOpen == true)  
                serialPort1.Close();        

        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            // Show all the incoming data in the port's buffer
            textBox2.AppendText(serialPort1.ReadLine() + Environment.NewLine);
        }
    }
}
