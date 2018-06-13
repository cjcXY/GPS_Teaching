using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GPSTeachingSys.OtherForms
{
    public partial class SheZhiTuPianShuXing : Form
    {
        Image im;
        public static SheZhiTuPianShuXing pcurrshuxing = null;
        Form1 fr1 = Form1.pCurrentWin;
        public SheZhiTuPianShuXing()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            fr1.webBrowser1.Document.InvokeScript("huoquzuobiaodian");
            fr1.webBrowser1.Document.Click += new HtmlElementEventHandler(Document_MouseDown_Second);
           // this.Cursor.Dispose();
            this.Visible = false;
            //fr1.Cursor.Dispose();
            
        }
        void Document_MouseDown_Second(object sender, HtmlElementEventArgs e)
        {
            
            textBox4.Text = fr1.webBrowser1.Document.GetElementById("zuobiaox").InnerText;
            textBox5.Text = fr1.webBrowser1.Document.GetElementById("zuobiaoy").InnerText;
            fr1.webBrowser1.Document.Click -= Document_MouseDown_Second;
            fr1.webBrowser1.Document.InvokeScript("endhuoquzuobiaodian");
            this.Visible = true;                    
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "图片|*.jpg;*.png;*.gif;*.bmp;";
            open.ShowDialog();
            textBox1.Text = open.FileName;
            if(textBox1.Text!=""){//判断是否真的添加了图片上去
            im=Image.FromFile(open.FileName);
            textBox3.Text =""+ im.Width;
            textBox2.Text =""+ im.Height;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(textBox1.Text!=""&&textBox2.Text!=""&&textBox3.Text!=""&&textBox4.Text!=""&&textBox5.Text!=""){
            fr1.webBrowser1.Document.GetElementById("iconurl").InnerText=textBox1.Text;
            fr1.webBrowser1.Document.GetElementById("iconwidth").InnerText = textBox3.Text;
            fr1.webBrowser1.Document.GetElementById("iconheight").InnerText = textBox2.Text;
            fr1.webBrowser1.Document.GetElementById("zuobiaox").InnerText = textBox4.Text;
            fr1.webBrowser1.Document.GetElementById("zuobiaoy").InnerText = textBox5.Text;
            fr1.webBrowser1.Document.InvokeScript("tianjiatuxiang");
            this.Close();
            }else if(textBox1.Text==""){//判断图片是否正确输入
                MessageBox.Show("图片路径不能为空", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }else if(textBox2.Text==""||textBox3.Text==""){//判断图片的大小是否正确输入
                MessageBox.Show("图片的大小不能为空", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else if(textBox4.Text==""||textBox5.Text==""){//判断图片的坐标值是否正确输入
                MessageBox.Show("图片的坐标值不能为空", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SheZhiTuPianShuXing_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        
    }
}
