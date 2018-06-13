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
            fr1.co = colorDialog1.Color;
            this.Close();
            MessageBox.Show("设置已成功更新", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
           // colorDialog1.ShowDialog();
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {

                button2.BackColor = colorDialog1.Color;
            }
            
        }

        private void XCCSSZ_Load(object sender, EventArgs e)
        {
            numericUpDown1.Value =decimal.Parse(fr1.webBrowser1.Document.GetElementById("carspeed").InnerText);
            numericUpDown2.Value = decimal.Parse(fr1.webBrowser1.Document.GetElementById("toumingdu").InnerText);
            numericUpDown3.Value = decimal.Parse(fr1.webBrowser1.Document.GetElementById("kuandu").InnerText);
            colorDialog1.Color = fr1.co;
            button2.BackColor = colorDialog1.Color;
        }

        private void XCCSSZ_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void XCCSSZ_Click(object sender, EventArgs e)
        {
           
        }


        private void XCCSSZ_MouseMove(object sender, MouseEventArgs e)
        {
        //    if (button2.BackColor != colorDialog1.Color)
         //   {
         //       button2.BackColor = colorDialog1.Color;
            //}
        }

        private void XCCSSZ_MouseLeave(object sender, EventArgs e)
        {

        }

        private void XCCSSZ_MouseCaptureChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
