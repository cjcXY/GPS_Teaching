using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GPSTeachingSys.OtherForms
{
    public partial class QRDH : Form
    {
        public QRDH()
        {
            InitializeComponent();
            timer1.Start();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            Form1 fr1 = Form1.pCurrentWin;
            fr1.webBrowser1.Document.InvokeScript("jixu");
            this.Close();
            fr1.webBrowser1.Document.InvokeScript("jisuanluchengshijian");
        }

        private void QRDH_Load(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (progressBar1.Value<100)
            progressBar1.Value += progressBar1.Step;
            else if (progressBar1.Value >= 100)
            {
                label1.Visible = true;
                button1.Visible = true;
                label2.Visible = false;
                timer1.Stop();
            }
        }
    }
}
