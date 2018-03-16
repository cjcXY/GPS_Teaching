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
    public partial class CeShiDian : Form
    {
        Form1 fr1 = Form1.pCurrentWin;
        OleDbConnection conn;
        OleDbCommand da;
        String mypath2;
        public CeShiDian()
        {
            InitializeComponent();
            mypath2 = Path.Combine(Application.StartupPath + @"..\..\..\..\data\Database1.mdb"); 
        }

        private void CeShiDian_Load(object sender, EventArgs e)
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
            for (int i = 0; i < nnn.Length-1; i++)
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
                TreeNode tnn = new TreeNode("测试点"+fr1.count);
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
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "测试点";
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
