using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace GPSTeachingSys.OtherForms
{
    public partial class LX : Form
    {
        public LX()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 fr1 = Form1.pCurrentWin;
            //   fr1.webBrowser1.Document.InvokeScript("Closefff");
            //  MessageBox.Show(textBox1.Text);
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("请输入有效的地名！");
            }
            else
            {
                fr1.webBrowser1.Document.GetElementById("point1").InnerText = textBox1.Text;
                fr1.webBrowser1.Document.GetElementById("point2").InnerText = textBox2.Text;
                fr1.webBrowser1.Document.InvokeScript("ssss");
                fr1.webBrowser1.Document.InvokeScript("jisuanluchengshijian");
                this.Close();
                QRDH queren = new QRDH();
                queren.Show();
            }
            textBox1.Text = "";
            textBox2.Text = "";
            
        }

        private void LX_FormClosed(object sender, FormClosedEventArgs e)
        {
        }
    }
}
