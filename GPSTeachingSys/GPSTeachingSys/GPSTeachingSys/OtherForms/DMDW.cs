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
    public partial class DMDW : Form
    {
        public DMDW()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 fr1 = Form1.pCurrentWin;
            if (textBox1.Text != "")
            {
                fr1.webBrowser1.Document.GetElementById("cityname").InnerText = textBox1.Text;
                fr1.webBrowser1.Document.InvokeScript("theLocation");
                this.Close();
            }
            else
                MessageBox.Show("请输入有效地地名！");
            textBox1.Text = "";
            
        }
    }
}
