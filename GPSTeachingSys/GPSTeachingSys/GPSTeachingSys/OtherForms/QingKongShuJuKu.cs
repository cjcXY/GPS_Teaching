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
    public partial class QingKongShuJuKu : Form
    {
        Form1 fr1 = Form1.pCurrentWin;
        SJGL shujuForm = SJGL.SuJuGuanLi;
        OleDbConnection conn;
        OleDbCommand da;
        String mypath2;
        public QingKongShuJuKu()
        {
            InitializeComponent();
            mypath2 = Path.Combine(Application.StartupPath + @"\data\Database1.mdb");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string txt1 =
           "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2 = "Delete From DW Where ID >0 ";

            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            conn.Close();
            fr1.webBrowser1.Document.InvokeScript("resecttt");
            fr1.RemovePoints();
            
            fr1.treeView1.Nodes[0].Collapse();
            fr1.treeView1.Nodes[0].Nodes.Clear();
            
            fr1.ID = 0;
            shujuForm.dataGridView2.DataSource = null;
            MessageBox.Show("数据已经清空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void QingKongShuJuKu_Load(object sender, EventArgs e)
        {

        }
    }
}
