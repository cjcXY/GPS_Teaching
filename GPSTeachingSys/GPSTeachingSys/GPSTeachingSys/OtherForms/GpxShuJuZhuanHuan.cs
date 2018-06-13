using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace GPSTeachingSys.OtherForms
{
    public partial class GpxShuJuZhuanHuan : Form
    {
        int a;
        int b;
        string st1;
        string[] st1_1;
        string[] st2_1;
        Form1 fr1 = Form1.pCurrentWin;
        string excelPath;
        string filePath = null;
        Boolean isok = true;
        public GpxShuJuZhuanHuan()
        {
            InitializeComponent();
            excelPath = Path.Combine(Application.StartupPath + @"\data\Sheet1$.xls");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "GPX文件(*.gpx)|*.gpx;*.xml";
                dlg.InitialDirectory=Path.Combine(Application.StartupPath + @"\data\");
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                filePath = dlg.FileName;
                this.textBox1.Text = filePath;
                isok = true;
            }
            


        }
        private void getFromXml()
        {
            //1.创建Applicaton对象
            Microsoft.Office.Interop.Excel.Application xApp = new

            Microsoft.Office.Interop.Excel.Application();

            //2.得到workbook对象，打开已有的文件
            Microsoft.Office.Interop.Excel.Workbook xBook = xApp.Workbooks.Open(excelPath,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                  Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //3.指定要操作的Sheet
            Microsoft.Office.Interop.Excel.Worksheet xSheet = (Microsoft.Office.Interop.Excel.Worksheet)xBook.Sheets[1];
           // Microsoft.Office.Interop.Excel.Range r = xSheet.Range[xSheet.Cells[1, 1], xSheet.Cells[xSheet.Rows., columnUsed]];
            xSheet.Rows.Clear();
            xSheet.Columns.Clear();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(filePath);
            XmlNodeList nodes;
            XmlNodeList nodes2;
            nodes = xmlDoc.GetElementsByTagName("trkpt");
            nodes2 = xmlDoc.GetElementsByTagName("time");
            int nn=nodes.Count;
            string[,] sttr=new string[nn+1,7];
            sttr[0,0] = "ID";
            sttr[0,1]="坐标点名称";
            sttr[0,2]="坐标点系列";
            sttr[0,3] = "UTM时间";
            sttr[0,4]="纬度";
            sttr[0,5]="经度";
            sttr[0,6] = "GPS定位状态";
            for (int i = 0; i < nn; i++)
            {
                a = i + 1;
              //  b = i + 1;
                sttr[a,0] =""+a;
               // MessageBox.Show(sttr[a, 0]);
                sttr[a,1] = "点" + a;
                sttr[a,2] = "点";
                st1 = nodes2[i].InnerText;
                st1_1 = st1.Split('T');
                // MessageBox.Show(st1_1[0]);
                st2_1 = st1_1[1].Split('Z');
                // MessageBox.Show(st2_1[0]);
                sttr[a,3] = st1_1[0] + " " + st2_1[0];
                XmlAttribute att = nodes[i].Attributes["lat"];
                //  MessageBox.Show(att.Value);
                sttr[a,4] = att.Value;
                XmlAttribute ong = nodes[i].Attributes["lon"];
                sttr[a,5] = ong.Value;
                sttr[a,6] = "差分定位";

          /*      //nodes = xmlDoc.GetElementsByTagName("trkpt");
                xSheet.Cells[1][a] = b;
                xSheet.Cells[2][a] = "点" + b;
                xSheet.Cells[3][a] = "点";        
                st1 = nodes2[i].InnerText;
                st1_1 = st1.Split('T');
                // MessageBox.Show(st1_1[0]);
                st2_1 = st1_1[1].Split('Z');
                // MessageBox.Show(st2_1[0]);
                xSheet.Cells[4][a] = st1_1[0] + " " + st2_1[0];
                XmlAttribute att = nodes[i].Attributes["lat"];　　
              //  MessageBox.Show(att.Value);
                xSheet.Cells[5][a] = att.Value;
                XmlAttribute ong = nodes[i].Attributes["lon"];
                xSheet.Cells[6][a] = ong.Value;
                xSheet.Cells[7][a] = "差分定位";
           * */

            }
      /*      nodes = xmlDoc.GetElementsByTagName("time");
            for (int i = 0; i < nodes.Count;i++ )
            {
                string st1 = nodes[i].InnerText;
                string[] st1_1 = st1.Split('T');
               // MessageBox.Show(st1_1[0]);
                string[] st2_1 = st1_1[1].Split('Z');
               // MessageBox.Show(st2_1[0]);
                xSheet.Cells[4][i + 2] = st1_1[0]+" "+ st2_1[0];
            }*/
          //  string nnnn="G"+(nodes.Count+1);
         //   xSheet.get_Range("A1:G8", Type.Missing).Value2 =(Object[,]) sttr;
         //   xSheet.get_Range("A1", xSheet.Cells[nn + 1, 7]).Value2 = sttr;
            Microsoft.Office.Interop.Excel.Range r = xSheet.Range[xSheet.Cells[1, 1], xSheet.Cells[nn+1, 7]];
            r.Value2 = sttr;
            xBook.Save();
            xSheet = null;
            xBook.Close();
            xBook = null;
            xApp.DisplayAlerts = false;
            xApp.Quit();
            xApp = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (isok)
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("未添加文件", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    getFromXml();
                    ShowExcel();                 
                    isok = false;
                }
            }
            else
            {
                MessageBox.Show("重复操作", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }

        private void ShowExcel()
        {
            try
            {
                OleDbConnectionStringBuilder connectStringBuilder = new OleDbConnectionStringBuilder();
                connectStringBuilder.DataSource = excelPath.Trim();
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
                MessageBox.Show("文件转换成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void GpxShuJuZhuanHuan_Load(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
           
        }
    }
}
