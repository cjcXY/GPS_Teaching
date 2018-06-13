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
        bool needChick = true;
        public MyPosition()
        {
            InitializeComponent();
        }

        private void MyPosition_Load(object sender, EventArgs e)
        {

            //MessageBox.Show(fr1.webBrowser1.Document.GetElementById("weizhi").InnerText);
            fr1.webBrowser1.Document.InvokeScript("WoDangQianWZ");
           

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
                fr1.webBrowser1.Document.InvokeScript("WoDangQianWZ");
                textBox1.Text = fr1.webBrowser1.Document.GetElementById("weizhi1").InnerText;
                textBox2.Text = fr1.webBrowser1.Document.GetElementById("weizhi2").InnerText;
                needChick = false;
            }
            else
            {
                textBox1.Text = "";
                textBox2.Text = "";
                needChick = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int sd;
            if(!needChick){
                fr1.webBrowser1.Document.GetElementById("weizhi1").InnerText = textBox1.Text;
                fr1.webBrowser1.Document.GetElementById("weizhi2").InnerText = textBox2.Text;
                fr1.webBrowser1.Document.InvokeScript("SSDW");
                this.Close();
            }else{
            if (textBox1.Text == ""|| textBox2.Text == "")
            {
                MessageBox.Show("请填写经纬度！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (!int.TryParse(textBox1.Text, out sd) || !int.TryParse(textBox2.Text, out sd))
            {
                MessageBox.Show("经纬度必须为数字", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (int.Parse(textBox1.Text) < -180 || int.Parse(textBox1.Text) > 180)
            {
                MessageBox.Show("经度的范围必须在[-180,180]内", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (int.Parse(textBox2.Text) < -90 || int.Parse(textBox2.Text) > 90)
            {
                MessageBox.Show("纬度的范围必须在[-90,90]内", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
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

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
