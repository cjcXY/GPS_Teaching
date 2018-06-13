//using ADOX;
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
    public partial class DLSJ : Form
    {
        Form1 fr1 = Form1.pCurrentWin;
        string filePath = null;
        private string mdbPath;
        public DLSJ()
        {
            InitializeComponent();
            mdbPath = Path.Combine(Application.StartupPath + @"\data\Database1.mdb"); 
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel文件(*.xls)|*.xls;*.xlsx";
            dlg.InitialDirectory = Path.Combine(Application.StartupPath + @"\data\");
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                filePath = dlg.FileName;
                this.textBox1.Text = filePath;
            
            try
            {
                OleDbConnectionStringBuilder connectStringBuilder = new OleDbConnectionStringBuilder();
                connectStringBuilder.DataSource = this.textBox1.Text.Trim();
                connectStringBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
                connectStringBuilder.Add("Extended Properties", "Excel 8.0");
                using (OleDbConnection cn = new OleDbConnection(connectStringBuilder.ConnectionString))
                {
                    DataSet ds = new DataSet();
                    string sql = "Select * from [Sheet1$]";
                    OleDbCommand cmdLiming = new OleDbCommand(sql, cn);
                    cn.Open();
                    using (OleDbDataReader drLiming = cmdLiming.ExecuteReader())
                    {
                        ds.Load(drLiming, LoadOption.OverwriteChanges, new string[] { "Sheet1" });
                        DataTable dt = ds.Tables["Sheet1"];
                        List<string> name = new List<string>();
                        List<Type> type = new List<Type>();
                        if (dt.Rows.Count > 0)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                name.Add(dt.Columns[i].ToString());//获取每一列的列名
                                type.Add(dt.Columns[i].DataType);//获取每一列的数据类型
                            }

                        }
                      
                        dataGridView1.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            }
        }

        private void DLSJ_Load(object sender, EventArgs e)
        {
            if (!File.Exists(mdbPath))
            {
               /* ADOX.Catalog catalog = new Catalog();
                catalog.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mdbPath + ";Jet OLEDB:Engine Type=5");*/
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0 )
            {
                MessageBox.Show("请选择需要导入的Execl文件","提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    OleDbConnectionStringBuilder connectStringBuilder = new OleDbConnectionStringBuilder();
                    connectStringBuilder.DataSource = this.textBox1.Text.Trim();
                    connectStringBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
                    connectStringBuilder.Add("Extended Properties", "Excel 8.0");
                    using (OleDbConnection cn = new OleDbConnection(connectStringBuilder.ConnectionString))
                    {
                        DataSet ds = new DataSet();
                        string sql = "Select * from [Sheet1$]";
                        OleDbCommand cmdLiming = new OleDbCommand(sql, cn);
                        cn.Open();
                        using (OleDbDataReader drLiming = cmdLiming.ExecuteReader())
                        {
                            ds.Load(drLiming, LoadOption.OverwriteChanges, new string[] { "Sheet1" });
                            DataTable dt = ds.Tables["Sheet1"];
                            List<string> name = new List<string>();
                            List<Type> type = new List<Type>();
                            if (dt.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    name.Add(dt.Columns[i].ToString());//获取每一列的列名
                                    type.Add(dt.Columns[i].DataType);//获取每一列的数据类型
                                }
                            }
                            fr1.UpdateID();
                            
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                int a=int.Parse(dt.Rows[i][0].ToString());
                                dt.Rows[i][0] = fr1.ID+a;
                               // MessageBox
                            }
                              //  MessageBox.Show(dt.Rows[1][0].ToString());
                            if (dt.Rows.Count > 0)
                            {
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    string MySql1 = "insert into DW(";
                                    for (int j = 0; j < name.Count - 1; j++)
                                    {
                                        if (!(dt.Rows[i][j] is DBNull))
                                        {
                                            MySql1 = MySql1 + "[" + name[j] + "]" + ",";
                                        }                                   
                                    }                               
                                    if (!(dt.Rows[i][name.Count - 1] is DBNull))
                                    {
                                        MySql1 = MySql1 + "[" + name[name.Count - 1] + "]" + ") values(";
                                    }
                                    else
                                    {
                                        MySql1 = MySql1.Remove(MySql1.Length - 1, 1);
                                        MySql1 = MySql1 + ") values(";
                                    }                        
                                    for (int j = 0; j < name.Count - 1; j++)
                                    {
                                        if (!(dt.Rows[i][j] is DBNull))
                                        {                                  
                                            MySql1 = MySql1 + "'" + dt.Rows[i][j] + "',";

                                        }

                                    }
                                    if (!(dt.Rows[i][name.Count - 1] is DBNull))
                                    {
                                        MySql1 = MySql1 + "'" + dt.Rows[i][name.Count - 1] + "')";
                                    }
                                    else
                                    {
                                        MySql1 = MySql1.Remove(MySql1.Length - 1, 1);
                                        MySql1 = MySql1 + ")";
                                    }
                                    SQLExecute(MySql1);
                                }
                                MessageBox.Show("数据导入成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                
                                fr1.UpdateGPS();
                                fr1.treeView1.Nodes[0].Expand();
                                fr1.RemovePoints();
                                fr1.GetPoints();
                                //fr1.UpdateCS();
                                this.Close();
                            }
                            else
                            {
                                MessageBox.Show("请检查你的Excel中是否存在数据", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        public bool SQLExecute(string sql)
        {
            try
            {
                string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mdbPath;
                OleDbConnection conn = new OleDbConnection(ConnectionString);
                conn.Open();
                OleDbCommand comm = new OleDbCommand();
                comm.Connection = conn;
                comm.CommandText = sql;
                comm.ExecuteNonQuery();
                comm.Connection.Close();
                conn.Close();
                return true;
            }
            catch
            {
                return false;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

       
          
    }
}
