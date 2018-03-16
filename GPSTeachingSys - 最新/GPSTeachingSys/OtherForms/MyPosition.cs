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
    public partial class MyPosition : Form
    {
        Form1 fr1 = Form1.pCurrentWin;
        public MyPosition()
        {
            InitializeComponent();
        }

        private void MyPosition_Load(object sender, EventArgs e)
        {

            //MessageBox.Show(fr1.webBrowser1.Document.GetElementById("weizhi").InnerText);
            // fr1.webBrowser1.Document.InvokeScript("WoDangQianWZ");
           

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void MyPosition_Shown(object sender, EventArgs e)
        {
            
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                textBox1.Text = fr1.webBrowser1.Document.GetElementById("weizhi1").InnerText;
                textBox2.Text = fr1.webBrowser1.Document.GetElementById("weizhi2").InnerText;
            }
            else
            {
                textBox1.Text = "";
                textBox2.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == ""|| textBox2.Text == "")
            {
                MessageBox.Show("请填写经纬度！");
            }
            else
            {
                fr1.webBrowser1.Document.GetElementById("weizhi1").InnerText = textBox1.Text;
                fr1.webBrowser1.Document.GetElementById("weizhi2").InnerText = textBox2.Text;
                fr1.webBrowser1.Document.InvokeScript("SSDW");
                this.Close();
            }
        }
    }
}
