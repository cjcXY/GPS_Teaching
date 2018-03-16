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
    public partial class GJHFSZ : Form
    {
        int []num=new int[1000];
        int i = 0;
        public GJHFSZ()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 fr1 = Form1.pCurrentWin;
            fr1.webBrowser1.Document.InvokeScript("getpointsnumber");
            int a=int.Parse(fr1.webBrowser1.Document.GetElementById("pointnum").InnerText);
            for (int i = 1; i <= a; i++)
            {
                treeView1.Nodes.Add("点" + i);
            }
            //num=new int[a];
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            contextMenuStrip1.Show(MousePosition);
        }

        private void 删除该点ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeView1.SelectedNode.Remove();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 fr1 = Form1.pCurrentWin;
            foreach (TreeNode tn in treeView1.Nodes)
            {
                var charNo = System.Text.RegularExpressions.Regex.Replace(tn + "", @"[^0-9]+", "");
              //  MessageBox.Show(charNo);
                num[i++] = int.Parse(charNo);
            }
            for (int i = 0; i <= num.Length; i++)
            {
                if (num[i] == 0)
                    break;
               // MessageBox.Show(num[i]+"");
                fr1.webBrowser1.Document.GetElementById("newpoint").InnerText = ""+num[i];
                fr1.webBrowser1.Document.InvokeScript("setpoints");
            }
            fr1.webBrowser1.Document.InvokeScript("setpointsok");
            fr1.webBrowser1.Document.InvokeScript("newpoint");
            this.Close();
            num=new int[1000];
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
    }
}
