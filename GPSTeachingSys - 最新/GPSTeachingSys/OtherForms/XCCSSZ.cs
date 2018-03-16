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
    public partial class XCCSSZ : Form
    {
        Form1 fr1 = Form1.pCurrentWin;
        public XCCSSZ()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            fr1.webBrowser1.Document.GetElementById("carspeed").InnerText = ""+numericUpDown1.Value ;
            fr1.webBrowser1.Document.GetElementById("yanse").InnerText = colorDialog1.Color.Name;
            fr1.webBrowser1.Document.GetElementById("toumingdu").InnerText = "" + numericUpDown2.Value;
            fr1.webBrowser1.Document.GetElementById("kuandu").InnerText = "" + numericUpDown3.Value;
            fr1.webBrowser1.Document.InvokeScript("xiaochecanshushezhi");
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
        }

        private void XCCSSZ_Load(object sender, EventArgs e)
        {
            numericUpDown1.Value =decimal.Parse(fr1.webBrowser1.Document.GetElementById("carspeed").InnerText);
            numericUpDown2.Value = decimal.Parse(fr1.webBrowser1.Document.GetElementById("toumingdu").InnerText);
            numericUpDown3.Value = decimal.Parse(fr1.webBrowser1.Document.GetElementById("kuandu").InnerText);
        }
    }
}
