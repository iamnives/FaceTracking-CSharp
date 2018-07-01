namespace CameraCapture
{
    partial class CameraCapture
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            ReleaseData();
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.captureButton = new System.Windows.Forms.Button();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.captureImageBox = new Emgu.CV.UI.ImageBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.detectCheckBox = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.serialCheckbox = new System.Windows.Forms.CheckBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.splitContainer4 = new System.Windows.Forms.SplitContainer();
            this.grayscaleImageBox = new Emgu.CV.UI.ImageBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.timerCOM = new System.Windows.Forms.Timer(this.components);
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.Panel2.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.captureImageBox)).BeginInit();
            this.panel1.SuspendLayout();
            this.splitContainer4.Panel1.SuspendLayout();
            this.splitContainer4.Panel2.SuspendLayout();
            this.splitContainer4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grayscaleImageBox)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.captureButton);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(1095, 691);
            this.splitContainer1.SplitterDistance = 36;
            this.splitContainer1.TabIndex = 0;
            // 
            // captureButton
            // 
            this.captureButton.Location = new System.Drawing.Point(15, 12);
            this.captureButton.Name = "captureButton";
            this.captureButton.Size = new System.Drawing.Size(102, 23);
            this.captureButton.TabIndex = 0;
            this.captureButton.Text = "Start Capture";
            this.captureButton.UseVisualStyleBackColor = true;
            this.captureButton.Click += new System.EventHandler(this.captureButtonClick);
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.splitContainer3);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.splitContainer4);
            this.splitContainer2.Size = new System.Drawing.Size(1095, 651);
            this.splitContainer2.SplitterDistance = 534;
            this.splitContainer2.TabIndex = 0;
            // 
            // splitContainer3
            // 
            this.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer3.Location = new System.Drawing.Point(0, 0);
            this.splitContainer3.Name = "splitContainer3";
            this.splitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.captureImageBox);
            this.splitContainer3.Panel1.Controls.Add(this.panel1);
            // 
            // splitContainer3.Panel2
            // 
            this.splitContainer3.Panel2.Controls.Add(this.detectCheckBox);
            this.splitContainer3.Panel2.Controls.Add(this.label4);
            this.splitContainer3.Panel2.Controls.Add(this.comboBox1);
            this.splitContainer3.Panel2.Controls.Add(this.serialCheckbox);
            this.splitContainer3.Panel2.Controls.Add(this.textBox2);
            this.splitContainer3.Size = new System.Drawing.Size(534, 651);
            this.splitContainer3.SplitterDistance = 313;
            this.splitContainer3.TabIndex = 0;
            // 
            // captureImageBox
            // 
            this.captureImageBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.captureImageBox.Location = new System.Drawing.Point(0, 33);
            this.captureImageBox.Name = "captureImageBox";
            this.captureImageBox.Size = new System.Drawing.Size(534, 280);
            this.captureImageBox.TabIndex = 1;
            this.captureImageBox.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(534, 33);
            this.panel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Captured Image:";
            // 
            // detectCheckBox
            // 
            this.detectCheckBox.AutoSize = true;
            this.detectCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.detectCheckBox.Location = new System.Drawing.Point(135, 25);
            this.detectCheckBox.Name = "detectCheckBox";
            this.detectCheckBox.Size = new System.Drawing.Size(119, 21);
            this.detectCheckBox.TabIndex = 4;
            this.detectCheckBox.Text = "Start detection";
            this.detectCheckBox.UseVisualStyleBackColor = true;
            this.detectCheckBox.CheckedChanged += new System.EventHandler(this.detectCheckBox_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label4.Location = new System.Drawing.Point(12, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 17);
            this.label4.TabIndex = 3;
            this.label4.Text = "Ports:";
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(15, 66);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(239, 21);
            this.comboBox1.TabIndex = 2;
            // 
            // serialCheckbox
            // 
            this.serialCheckbox.AutoSize = true;
            this.serialCheckbox.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.serialCheckbox.Location = new System.Drawing.Point(15, 25);
            this.serialCheckbox.Name = "serialCheckbox";
            this.serialCheckbox.Size = new System.Drawing.Size(114, 21);
            this.serialCheckbox.TabIndex = 1;
            this.serialCheckbox.Text = "Send to serial";
            this.serialCheckbox.UseVisualStyleBackColor = true;
            this.serialCheckbox.CheckedChanged += new System.EventHandler(this.serialCheckbox_CheckedChanged);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(15, 93);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(502, 229);
            this.textBox2.TabIndex = 0;
            // 
            // splitContainer4
            // 
            this.splitContainer4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer4.Location = new System.Drawing.Point(0, 0);
            this.splitContainer4.Name = "splitContainer4";
            this.splitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer4.Panel1
            // 
            this.splitContainer4.Panel1.Controls.Add(this.grayscaleImageBox);
            this.splitContainer4.Panel1.Controls.Add(this.panel2);
            // 
            // splitContainer4.Panel2
            // 
            this.splitContainer4.Panel2.Controls.Add(this.label3);
            this.splitContainer4.Panel2.Controls.Add(this.textBox1);
            this.splitContainer4.Size = new System.Drawing.Size(557, 651);
            this.splitContainer4.SplitterDistance = 313;
            this.splitContainer4.TabIndex = 0;
            // 
            // grayscaleImageBox
            // 
            this.grayscaleImageBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.grayscaleImageBox.Location = new System.Drawing.Point(0, 33);
            this.grayscaleImageBox.Name = "grayscaleImageBox";
            this.grayscaleImageBox.Size = new System.Drawing.Size(557, 280);
            this.grayscaleImageBox.TabIndex = 1;
            this.grayscaleImageBox.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label2);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(557, 33);
            this.panel2.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Grayscale Image:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.label3.Location = new System.Drawing.Point(48, 29);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(43, 17);
            this.label3.TabIndex = 2;
            this.label3.Text = "Face:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(97, 29);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(228, 20);
            this.textBox1.TabIndex = 1;
            // 
            // timerCOM
            // 
            this.timerCOM.Enabled = true;
            this.timerCOM.Interval = 1000;
            this.timerCOM.Tick += new System.EventHandler(this.timerCOM_Tick);
            // 
            // CameraCapture
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1095, 691);
            this.Controls.Add(this.splitContainer1);
            this.Name = "CameraCapture";
            this.Text = "Camera Capture";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.Panel2.ResumeLayout(false);
            this.splitContainer3.Panel2.PerformLayout();
            this.splitContainer3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.captureImageBox)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.splitContainer4.Panel1.ResumeLayout(false);
            this.splitContainer4.Panel2.ResumeLayout(false);
            this.splitContainer4.Panel2.PerformLayout();
            this.splitContainer4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grayscaleImageBox)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Button captureButton;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.SplitContainer splitContainer3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.SplitContainer splitContainer4;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label2;
        private Emgu.CV.UI.ImageBox captureImageBox;
        private Emgu.CV.UI.ImageBox grayscaleImageBox;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.CheckBox serialCheckbox;
        private System.Windows.Forms.CheckBox detectCheckBox;
        private System.Windows.Forms.Timer timerCOM;
        private System.IO.Ports.SerialPort serialPort1;
    }
}

