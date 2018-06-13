using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace GPSTeachingSys.OtherForms
{
    public partial class KML : Form
    {
        string filePath = null;
        Form1 fr1 = Form1.pCurrentWin;
        public KML()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "KML文件(*.kml)|*.kml;*.xml";
            dlg.InitialDirectory = Path.Combine(Application.StartupPath + @"\data\");
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                filePath = dlg.FileName;
                this.textBox1.Text = filePath;
            }
        }
        private void getFromKml()
        {       
            if (textBox1.Text.Trim()!= "")
            {                  
                fr1.webBrowser1.Document.GetElementById("KML_latlng").InnerText = filePath;             
                fr1.webBrowser1.Document.InvokeScript("readKML");
                fr1.webBrowser1.Document.InvokeScript("addKML");
                MessageBox.Show("数据加载成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            else
            {
                MessageBox.Show("请加载KML文件", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            getFromKml();
           
        }

        private void KML_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
