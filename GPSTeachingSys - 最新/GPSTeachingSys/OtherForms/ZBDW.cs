using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GPSTeachingSys;

namespace GPSTeachingSys.OtherForms
{
    public partial class ZBDW : Form
    {
        public ZBDW()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 fr1 = Form1.pCurrentWin;
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("请输入有效地经纬度！");
            }
            else
            {
                fr1.webBrowser1.Document.GetElementById("longittude").InnerText = textBox1.Text;
                fr1.webBrowser1.Document.GetElementById("latitude").InnerText = textBox2.Text;
                fr1.webBrowser1.Document.InvokeScript("theLocation2");
                textBox1.Text = "";
                textBox2.Text = "";
                this.Close();
            }
        }
    }
}
