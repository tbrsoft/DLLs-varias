using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using usb_api;

namespace InterruptTransferTest
{

    public partial class Form1 : Form
    {
        PICAsyncManager pam;
        public Form1()
        {
            InitializeComponent();
            pam = new PICAsyncManager();
            this.Controls.Add(pam);
            pam.EventNewUSBIntData += new NewUSBIntDataEventHandler(AppendNewUSBIntData);
        }        

        private void button1_Click(object sender, EventArgs e)
        {

            if (!pam.IsRunning)
            {                
                pam.Start();                
                textBox1.AppendText("Interrupt monitor thread started...");
            }
            else
            {
                pam.Stop();
                textBox1.AppendText("Interrupt monitor thread stopped...");
            }
        }

        private void AppendNewUSBIntData(NewUSBIntDataEventArgs e)
        {
            textBox1.AppendText(e.newdata[1].ToString()+" ");
        }
    }
}