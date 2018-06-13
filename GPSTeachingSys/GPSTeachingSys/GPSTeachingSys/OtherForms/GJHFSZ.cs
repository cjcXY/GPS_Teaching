using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GPSTeachingSys.OtherForms
{
    public partial class GJHFSZ : Form
    {
        OleDbConnection conn;
        OleDbCommand da;
        Form1 fr1 = Form1.pCurrentWin;
        int []num=new int[1000];
        int i = 0;
        String mypath2;
        public GJHFSZ()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 fr1 = Form1.pCurrentWin;
            fr1.webBrowser1.Document.InvokeScript("getpointsnumber");
            int a=int.Parse(fr1.webBrowser1.Document.GetElementById("pointnum").InnerText);
            if(a==0){
                MessageBox.Show("您还没有绘制轨迹点", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            treeView1.Nodes.Clear();
            for (int i = 1; i <= a; i++)
            {
                treeView1.Nodes.Add("测试点" + i);
            }
            okook();
          //  fr1.UpdateCS();
            
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            contextMenuStrip1.Show(MousePosition);
        }

        private void 删除该点ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string tex = treeView1.SelectedNode.Text;
            if(tex!=""){
            string result = System.Text.RegularExpressions.Regex.Replace(tex, @"[^0-9]+", "");
            // int a = int.Parse(result);
            string txt1 =
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;

            // int a = ds.Tables[0].Rows.Count;
            string txt2 = "Delete From 测试点 Where ID = " + result;
            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            conn.Close();
            treeView1.SelectedNode.Remove();
            }
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

        private void 定位该点ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            string tex = treeView1.SelectedNode.Text;
            if(tex!=""){
            string result = System.Text.RegularExpressions.Regex.Replace(tex, @"[^0-9]+", "");
            int a = int.Parse(result);
            result = "" + a;
            //MessageBox.Show(result);
            string txt1 =
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt3 = "Select 经度,纬度 From 测试点 Where ID =" + result;
            conn = new OleDbConnection(txt1);
            OleDbDataAdapter daa = new OleDbDataAdapter(txt3, conn);
            DataSet ds = new DataSet("ds");
            daa.Fill(ds);
            string[] pointx = ds.Tables[0].Rows[0][0].ToString().Split('E');
            fr1.webBrowser1.Document.GetElementById("ceshidianx").InnerText = pointx[0];
            string[] pointy = ds.Tables[0].Rows[0][1].ToString().Split('N');
            fr1.webBrowser1.Document.GetElementById("ceshidiany").InnerText = pointy[0];
            fr1.webBrowser1.Document.GetElementById("ceshidianb").InnerText = result;
            fr1.webBrowser1.Document.InvokeScript("SuoFangDaoDian");
            conn.Close();
            }
        }

        private void GJHFSZ_Load(object sender, EventArgs e)
        {
            mypath2 = Path.Combine(Application.StartupPath + @"\data\Database1.mdb");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void okook()
        {
            string txt1_1 =
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2_1 = "Delete From 测试点 ";

            conn = new OleDbConnection(txt1_1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2_1;
            da.Connection = conn;
            da.ExecuteNonQuery();
            conn.Close();
            //fr1.treeView1.Nodes[1].Nodes.Clear();





            string mmm = fr1.webBrowser1.Document.GetElementById("ceshidian").InnerText;
            if (mmm == "" || mmm == null)
            {
                return;
            }
            string[] nnn = mmm.Split(';');
            string txt1 =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            for (int i = 0; i < nnn.Length - 1; i++)
            {
                string txt2 =
           "Insert Into 测试点(ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态) Values('" +
                nnn[i] + "')";
                conn = new OleDbConnection(txt1);
                conn.Open();
                da = new OleDbCommand();
                da.CommandText = txt2;
                da.Connection = conn;
                da.ExecuteNonQuery();
                conn.Close();
                TreeNode tnn = new TreeNode("测试点" + fr1.count);
                fr1.treeView1.Nodes[1].Nodes.Add(tnn);
                fr1.count++;
            }


            string txtConn =
     "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM 测试点";
            OleDbDataAdapter daa = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            daa.Fill(ds, "测试点");
            //dataGridView1.DataSource = ds;
            //dataGridView1.DataMember = "测试点";
        }
    }
}
