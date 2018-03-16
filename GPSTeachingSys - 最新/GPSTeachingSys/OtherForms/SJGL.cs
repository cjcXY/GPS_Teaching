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

namespace GPSTeachingSys.OtherForms
{
    public partial class SJGL : Form
    {
        Form1 fr1 = Form1.pCurrentWin;
        OleDbConnection conn;
        OleDbCommand da;
        String mypath2 ;
        string txtCommand;
        string savedExcelpath;
        List<string> lt = new List<string>();
        public SJGL()
        {
            InitializeComponent();
            mypath2 = Path.Combine(Application.StartupPath + @"..\..\..\..\data\Database1.mdb");
            savedExcelpath = Path.Combine(Application.StartupPath + @"..\..\..\..\data\savedExcel.xls"); 
        }

        private void SJGL_Load(object sender, EventArgs e)
        {
            
       /*     string txtConn =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
           // dataGridView1.DataSource = ds;
          //  dataGridView1.DataMember = "DW";
            dataGridView2.DataSource = ds;
            dataGridView2.DataMember = "DW";
            //MessageBox.Show(dateTimePicker1.Value + "");
         */
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
      /*      if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "")
            {
                MessageBox.Show("所有项都必须填写！ ");
                return;
            }//注意，下边语句必须保证数据库文件studentI.mdb在文件夹路径C:\\**中
            string txt1 =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2 =
       "Insert Into DW(ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态) Values('";
            txt2 += textBox1.Text + "' , '";
            txt2 += textBox2.Text + "' , '";
            txt2 += textBox3.Text + "' , '";
            txt2 += textBox4.Text + "' , '";
            txt2 += textBox5.Text + "' , '";
            txt2 += textBox6.Text + "' , '";
            txt2 += textBox7.Text + "')";
            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            conn.Close();*/
        }

        private void button2_Click(object sender, EventArgs e)
        {
       /*     if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "" || textBox5.Text == "" || textBox6.Text == "" || textBox7.Text == "")
            {
                MessageBox.Show("所有项都必须填写！ 空值填写null");
                return;
            }//注意，下边语句必须保证数据库文件studentI.mdb在文件夹路径C:\\**中
            string txt1 =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2 = "Update DW Set 坐标点名称 = '" + textBox2.Text + "',坐标点系列 = '" + textBox3.Text + "',UTM时间 = '" + textBox4.Text + "',纬度 = '" + textBox5.Text + "',经度 = '" + textBox6.Text + "',GPS定位状态 = '" + textBox7.Text + "'Where ID = " + textBox1.Text;
            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            conn.Close();*/
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "DW";
            fr1.UpdateGPS();

        }

        private void button3_Click(object sender, EventArgs e)
        {
   /*         if (textBox1.Text == "")
            {
                MessageBox.Show("编号项必须填写！ ");
                return;
            }//注意，下边语句必须保证数据库文件studentI.mdb在文件夹路径C:\\**中
            string txt1 =
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2 = "Delete From DW Where ID = " + textBox1.Text;

            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            conn.Close();*/
        }

        private void button5_Click(object sender, EventArgs e)
        {
      /*      string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT " + comboBox2.Text + "," + comboBox3.Text + " FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            dataGridView2.DataSource = ds;
            dataGridView2.DataMember = "DW";*/
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            QingKongShuJuKu qikong = new QingKongShuJuKu();
            qikong.Show();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void domainUpDown2_TextChanged(object sender, EventArgs e)
        {
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            string txtConn =
      "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            if (textBox1.Text == "")
            {
                if(comboBox1.Text==""){
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker1.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker2.Value + "'))";
                }
                else if(dateTimePicker1.Checked==false||dateTimePicker2.Checked==false){
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (坐标点系列='" + comboBox1.Text + "')";
                }
                else
                txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker1.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker2.Value + "')) and(坐标点系列='" + comboBox1.Text + "')";
            }
            else if (comboBox1.Text == "")
            {
                if (dateTimePicker1.Checked == false || dateTimePicker2.Checked == false)
                {
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (坐标点名称='" + textBox1.Text + "')";
                }
                else
                txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker1.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker2.Value + "'))and (坐标点名称='" + textBox1.Text + "')";
            }
            else if (dateTimePicker1.Checked == false || dateTimePicker2.Checked == false)
            {
                txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (坐标点系列='" + comboBox1.Text + "')and (坐标点名称='" + textBox1.Text + "')";
            }
            else
                txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker1.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker2.Value + "')) and(坐标点系列='" + comboBox1.Text + "')and (坐标点名称='" + textBox1.Text + "')";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            dataGridView1.DataSource = ds;
            dataGridView1.DataMember = "DW";
        }


        public void ToExcel(DataGridView dataGridView1)
        {
            try
            {
                //没有数据的话就不往下执行  
                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("没有数据可以导出！");
                    return;
                }
                //实例化一个Excel.Application对象  
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

                //让后台执行设置为不可见，为true的话会看到打开一个Excel，然后数据在往里写  
                excel.Visible = true;

                //新增加一个工作簿，Workbook是直接保存，不会弹出保存对话框，加上Application会弹出保存对话框，值为false会报错  
                excel.Application.Workbooks.Add(true);
                //生成Excel中列头名称  
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (this.dataGridView1.Columns[i].Visible == true)
                    {
                        excel.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                    }

                }
                //把DataGridView当前页的数据保存在Excel中  
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    System.Windows.Forms.Application.DoEvents();
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (this.dataGridView1.Columns[j].Visible == true)
                        {
                            if (dataGridView1[j, i].ValueType == typeof(string))
                            {
                                excel.Cells[i + 2, j + 1] = "'" + dataGridView1[j, i].Value.ToString();
                            }
                            else
                            {
                                excel.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                            }
                        }

                    }
                }

                //设置禁止弹出保存和覆盖的询问提示框  
                excel.DisplayAlerts = false;
                excel.AlertBeforeOverwriting = false;

                //保存工作簿  
                excel.Application.Workbooks.Add(true).Save();
                //保存excel文件  
                excel.Save();

                //确保Excel进程关闭  
                excel.Quit();
                excel = null;
                GC.Collect();//如果不使用这条语句会导致excel进程无法正常退出，使用后正常退出
                MessageBox.Show(this,"文件已经成功导出！", "信息提示");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误提示");
            }

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //ToExcel(dataGridView1);
            ExportToExcel d = new ExportToExcel();
            d.OutputAsExcelFile(dataGridView1); 
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            string txtConn =
     "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            if (textBox2.Text == "")
            {
                if (comboBox2.Text == "")
                {
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker4.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker3.Value + "'))";
                }
                else if (dateTimePicker4.Checked == false || dateTimePicker3.Checked == false)
                {
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (坐标点系列='" + comboBox2.Text + "')";
                }
                else
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker4.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker3.Value + "')) and(坐标点系列='" + comboBox2.Text + "')";
            }
            else if (comboBox2.Text == "")
            {
                if (dateTimePicker4.Checked == false || dateTimePicker3.Checked == false)
                {
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (坐标点名称='" + textBox2.Text + "')";
                }
                else
                    txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker4.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker3.Value + "'))and (坐标点名称='" + textBox2.Text + "')";
            }
            else if (dateTimePicker4.Checked == false || dateTimePicker3.Checked == false)
            {
                txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (坐标点系列='" + comboBox2.Text + "')and (坐标点名称='" + textBox2.Text + "')";
            }
            else
                txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW Where (cdate(UTM时间) >= cdate('" + dateTimePicker4.Value + "') and cdate(UTM时间)<=cdate('" + dateTimePicker3.Value + "')) and(坐标点系列='" + comboBox2.Text + "')and (坐标点名称='" + textBox2.Text + "')";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            dataGridView2.DataSource = ds;
            dataGridView2.DataMember = "DW";
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
         
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
                dataGridView2.Rows[i].Cells[0].Value = true;
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            
          //  dataGridView2.EndEdit();
          //  MessageBox.Show(dataGridView2.Rows[0].Cells[0].Value.ToString());
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {

                if ((bool)dataGridView2.Rows[i].Cells[0].EditedFormattedValue == true)
                    dataGridView2.Rows[i].Cells[0].Value = false;
                else
                    dataGridView2.Rows[i].Cells[0].Value = true;
            }
                
        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {

                if ((bool)dataGridView2.Rows[i].Cells[0].EditedFormattedValue == true)
                    lt.Add(dataGridView2.Rows[i].Cells[1].Value.ToString());
            }
             
            string txt1 =
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txt1);
            conn.Open();
            for (int i = 0; i < lt.Count; i++)
            {
                if (lt[i] == "")
                {
                    lt = new List<string>();
                    break;
                }
                string txt2 = "Delete From DW Where ID = " + lt[i];
                da = new OleDbCommand();
                da.CommandText = txt2;
                da.Connection = conn;
                da.ExecuteNonQuery();
            }

            conn.Close();
            UpSJK();
        }
        private void UpSJK()
        {
            string txtConn =
      "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            dataGridView2.DataSource = ds;
            dataGridView2.DataMember = "DW";
            conn.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            dataGridView1.DataSource = ds;
             dataGridView1.DataMember = "DW";
           // dataGridView2.DataSource = ds;
          //  dataGridView2.DataMember = "DW";
            //MessageBox.Show(dateTimePicker1.Value + "");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,坐标点名称,坐标点系列,UTM时间,纬度,经度,GPS定位状态 FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            // dataGridView1.DataSource = ds;
            //  dataGridView1.DataMember = "DW";
            dataGridView2.DataSource = ds;
            dataGridView2.DataMember = "DW";
            //MessageBox.Show(dateTimePicker1.Value + "");
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

    }
}
