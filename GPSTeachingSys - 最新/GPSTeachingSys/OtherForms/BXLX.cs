﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GPSTeachingSys.OtherForms
{
    public partial class BXLX : Form
    {
        public BXLX()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 fr1 = Form1.pCurrentWin;
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("请输入有效的地名！");
            }
            else
            {
                fr1.webBrowser1.Document.GetElementById("buxingstart").InnerText = textBox1.Text;
                fr1.webBrowser1.Document.GetElementById("buxingstop").InnerText = textBox2.Text;
                fr1.webBrowser1.Document.InvokeScript("buxing");
                // MessageBox.Show(fr1.webBrowser1.Document.GetElementById("n-result").InnerText);
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }
    }
}
