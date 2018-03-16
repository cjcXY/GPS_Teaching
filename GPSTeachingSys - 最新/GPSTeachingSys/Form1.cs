using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using mshtml;
using System.Web;
using System.Threading;
using System.Web.UI;
using System.Drawing.Imaging;
using System.Data.OleDb;
using System.IO; 

namespace GPSTeachingSys
{
    public partial class Form1 : Form
    {
        List<Button> list = new List<Button>(); 
        OleDbConnection conn;
        OleDbCommand da;
        String mypath2;
        public int ID;
        int b = 0;
        public int count = 1;
        public static Form1 pCurrentWin = null;
        public string strURl = "";
        public string strImgFileName = "123";
        public Page Context = null;  
       // public string path = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
        public Form1()
        {
            InitializeComponent();

            mypath2 = Path.Combine(Application.StartupPath + @"..\..\..\..\data\Database1.mdb");
            
            //mypath2 = Example_01.getPath(path) + "\\GPSTeachingSys - 副本 - 副本\\data\\Database1.mdb";
            pCurrentWin = this;
        }

        private void BLC_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
          String mypath = Path.Combine(Application.StartupPath + @"..\..\..\..\data\BaiDuMaps.html");
         //   String mypath = "https://worldwind.arc.nasa.gov/";
            treeView1.Top = button2.Height + button2.Location.Y;
            webBrowser1.Navigate(mypath);
            UpdateGPS();
            UpdateCS();
            UpdateID();
        }

       

        private void 地名路线ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void ZBDW_Click(object sender, EventArgs e)
        {
            //OtherForms.ZBDW zuobiaodingwei = new OtherForms.ZBDW();
           // zuobiaodingwei.Show();
        }

        private void SS_Click(object sender, EventArgs e)
        {
        }

        private void GB_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void splitContainer1_SplitterMoving(object sender, SplitterCancelEventArgs e)
        {



        }

     

        private void Form1_SizeChanged(object sender, EventArgs e)
        {

        }

        private void splitContainer1_SplitterMoving_1(object sender, SplitterCancelEventArgs e)
        {
            splitContainer1.Panel2MinSize = splitContainer1.Width * 4 / 9;
            splitContainer1.Panel1MinSize = splitContainer1.Width * 2 / 12;
        }

        private void splitContainer1_SplitterMoved_1(object sender, SplitterEventArgs e)
        {
            splitContainer1.Panel2MinSize = 0;
            splitContainer1.Panel1MinSize = 0;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
                webBrowser1.Document.Click += new HtmlElementEventHandler(Document_MouseDown);
              //  GetPoints();
                webBrowser1.Document.InvokeScript("WoDangQianWZ");
            
        }

        private void bnbn_MouseWeel(object sender, MouseEventArgs e)
        {
            toolStripComboBox1.Text = webBrowser1.Document.GetElementById("class").InnerText;
        }

        private void DRSJ_Click(object sender, EventArgs e)
        {
            //String mypath = Example_01.getPath(path) + "\\GPSTeachingSys - 副本 - 副本\\data\\BaiDuMaps.html";
            // webBrowser1.Navigate(mypath);
           // webBrowser1.Document.InvokeScript("Closefff");
            OtherForms.DLSJ daorushuju = new OtherForms.DLSJ();
            daorushuju.Show();
            
        }

        private void toolStripComboBox1_TextChanged(object sender, EventArgs e)
        {
            if (webBrowser1.Document.GetElementById("class").InnerText.Equals(toolStripComboBox1.Text))
            {
            }
            else
            {
               // MessageBox.Show(webBrowser1.Document.GetElementById("class").InnerText);
                webBrowser1.Document.GetElementById("class").InnerText = toolStripComboBox1.Text;
                webBrowser1.Document.InvokeScript("SuoFang");
            }
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void webBrowser1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {

        }

        private void statusStrip1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
        void Document_MouseDown(object sender, HtmlElementEventArgs e)
        {
            toolStripStatusLabel2.Text = webBrowser1.Document.GetElementById("city").InnerText;
            toolStripStatusLabel6.Text = webBrowser1.Document.GetElementById("pointx").InnerText;
            toolStripStatusLabel9.Text = webBrowser1.Document.GetElementById("pointy").InnerText;
            if (!toolStripComboBox1.Text.Equals(webBrowser1.Document.GetElementById("class").InnerText))
            {
                toolStripComboBox1.Text = webBrowser1.Document.GetElementById("class").InnerText;
            }
        }

        private void LX_Click(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel2_TextChanged(object sender, EventArgs e)
        {
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
           /* if (treeView1.SelectedNode.Text.Equals("起点"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableQD");
                }
                else
                {
                    webBrowser1.Document.InvokeScript("disableQD");
                }
            }
            if (treeView1.SelectedNode.Text.Equals("节点1"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableJD1");
                }
                else
                {
                    webBrowser1.Document.InvokeScript("disableJD1");
                }
            }
            if (treeView1.SelectedNode.Text.Equals("节点2"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableJD2");
                }
                else
                {
                    webBrowser1.Document.InvokeScript("disableJD2");
                }
            }
            if (treeView1.SelectedNode.Text.Equals("节点3"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableJD3");
                }
                else
                {
                    webBrowser1.Document.InvokeScript("disableJD3");
                }
            }
            if (treeView1.SelectedNode.Text.Equals("终点"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableZD");
                }
                else
                {
                    webBrowser1.Document.InvokeScript("disableZD");
                }
            }
            if (treeView1.SelectedNode.Text.Equals("连接"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableLX");
                }
                else
                {
                    webBrowser1.Document.InvokeScript("disableLX");
                }
            }
            if (treeView1.SelectedNode.Text.Equals("连线成面"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableLXCM");
                }
                else
                {
                    webBrowser1.Document.InvokeScript("disableLXCM");
                }
            }
            if (treeView1.SelectedNode.Text.Equals("导航"))
            {
                if (treeView1.SelectedNode.Checked == true)
                {
                    webBrowser1.Document.InvokeScript("enableDH");
                }
                else
                {

                    foreach (TreeNode tn in e.Node.Parent.Nodes)
                    {
                        if (tn.Checked == true)
                            tn.Checked = false;
                    }
                    webBrowser1.Document.InvokeScript("disableDH");
                }
            }*/

        }

        private void treeView1_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        {

        }

        private void treeView1_AfterExpand(object sender, TreeViewEventArgs e)
        {
            if(e.Node.Text.Equals("GPS坐标点"))
            {
                UpdateGPS();
                GetPoints();

            }
            if (e.Node.Text.Equals("测试点"))
                UpdateCS();
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {

        }

        private void XCCSSZ_Click(object sender, EventArgs e)
        {
            OtherForms.XCCSSZ xiaochecanshu = new OtherForms.XCCSSZ();
            xiaochecanshu.Show();
        }

        private void HZGJ_Click(object sender, EventArgs e)
        {

        }

        private void 开始绘制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("在地图上单击获得定位点");
            webBrowser1.Document.InvokeScript("addpoints");
            停止绘制ToolStripMenuItem.Enabled = true;
            开始绘制ToolStripMenuItem.Enabled = false;
            if (GJBF.Enabled == true)
                GJBF.Enabled = false;
            if (擦除图层.Enabled == true)
                擦除图层.Enabled = false;
         //   if (BFKZ.Enabled == true)
             //   BFKZ.Enabled = false;
        }

        private void 停止绘制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("轨迹绘制已停止！");
            webBrowser1.Document.InvokeScript("endaddpoints");
            GJBF.Enabled = true;
            擦除图层.Enabled = true;
            开始绘制ToolStripMenuItem.Enabled = true;
            停止绘制ToolStripMenuItem.Enabled = false;
        }

        private void GJBF_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("guijibofang");
            GJBF.Enabled = false;
            BFKZ.Enabled = true;
            if (擦除图层.Enabled == true)
                擦除图层.Enabled = false;
            if (HZGJ.Enabled == true)
                HZGJ.Enabled = false;
            if (开始播放ToolStripMenuItem.Enabled == true)
                开始播放ToolStripMenuItem.Enabled = false;
            if (暂停播放ToolStripMenuItem.Enabled == false)
                暂停播放ToolStripMenuItem.Enabled = true;
            if (停止播放ToolStripMenuItem.Enabled == false)
                停止播放ToolStripMenuItem.Enabled = true;
        }

        private void 开始播放ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("jixu");
            暂停播放ToolStripMenuItem.Enabled = true;
            停止播放ToolStripMenuItem.Enabled = true;
            开始播放ToolStripMenuItem.Enabled = false;
            if (擦除图层.Enabled == true)
                擦除图层.Enabled = false;
            if (HZGJ.Enabled == true)
                HZGJ.Enabled = false;
            if (GJBF.Enabled == true)
                GJBF.Enabled = false;
        }

        private void 暂停播放ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("zanting");
            if (开始播放ToolStripMenuItem.Enabled==false)
            开始播放ToolStripMenuItem.Enabled = true;
            暂停播放ToolStripMenuItem.Enabled = false;
            if (擦除图层.Enabled == true)
                擦除图层.Enabled = false;
        }

        private void 停止播放ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("tingzhi");
            if (开始播放ToolStripMenuItem.Enabled == false)
                开始播放ToolStripMenuItem.Enabled = true;
            停止播放ToolStripMenuItem.Enabled = false;
            擦除图层.Enabled = true;
            if (HZGJ.Enabled == false)
                HZGJ.Enabled = true;
        }

        private void GJHFSZ_Click(object sender, EventArgs e)
        {
            OtherForms.GJHFSZ guiji = new OtherForms.GJHFSZ();
            guiji.Show();
        }

        private void WDWZ_Click(object sender, EventArgs e)
        {
            //webBrowser1.Document.InvokeScript("WoDangQianWZ");
            //webBrowser1.Document.InvokeScript("myposition");
            
        }

        private void 擦除图层_Click(object sender, EventArgs e)
        {
            //webBrowser1.Document.InvokeScript("tingzhi");
            webBrowser1.Document.InvokeScript("resecttt");
            treeView1.Nodes[1].Nodes.Clear();
          
            string txt1 =
           "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2 = "Delete From DW Where ID <0 ";
            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            conn.Close();
            擦除图层.Enabled = false;
            if (GJBF.Enabled == true)
                GJBF.Enabled = false;
            if (BFKZ.Enabled == true)
                BFKZ.Enabled = false;
        }

        private void 驾车路线ToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  OtherForms.LX luxian = new OtherForms.LX();
          //  luxian.Show();
        }

        private void 步行路线ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 路线确认ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 开始步行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void KSCL_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("openceju");
        }

        private void JSCL_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("closeceju");
        }

        private void 开启ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("oplakuangfangda");
        }

        private void 关闭ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("cllakuangfangda");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
         /*   if (checkBox1.Checked == true)
            {
                webBrowser1.Document.InvokeScript("addsuonuetu");
            }
            if (checkBox1.Checked == false)
            {
                webBrowser1.Document.InvokeScript("removesuonuetu");
            }*/
        }

        private void 获取规划ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(webBrowser1.Document.GetElementById("n-result").InnerText);
        }

        private void 路线设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
        //    OtherForms.BXLX buxing = new OtherForms.BXLX();
         //   buxing.Show();
        }

        private void DY_Click(object sender, EventArgs e)
        {
            webBrowser1.Print();
        }

        private void DYYL_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPrintPreviewDialog();
        }

        private void DYSZ_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPrintDialog();
        }

        private void BCTX_Click(object sender, EventArgs e)
        {
           // startImg();
            ExportMapToImage();
            //webBrowser1.ShowSaveAsDialog();
        }
        private void ExportMapToImage()
        {
            try
            {
                SaveFileDialog pSaveDialog = new SaveFileDialog();
                pSaveDialog.FileName = "";
                pSaveDialog.Filter = "JPG图片(*.JPG)|*.jpg|GIF图片(*.GIF)|*.gif|PNG文档(*.PNG)|*.png|BMP图像(*.BMP)|*.bmp";
                pSaveDialog.AddExtension = true;
                pSaveDialog.InitialDirectory = Path.Combine(Application.StartupPath + @"..\..\..\..\data\");
                if (pSaveDialog.ShowDialog() == DialogResult.OK)
                {
                    if (pSaveDialog.FilterIndex == 1)
                    {
                        Size mySize = webBrowser1.Document.Window.Size;
                        Bitmap myPic = new Bitmap(mySize.Width, mySize.Height);
                        Rectangle myRec = new Rectangle(0, 0, mySize.Width, mySize.Height);
                        webBrowser1.Size = mySize;
                        webBrowser1.DrawToBitmap(myPic, myRec);
                        myPic.Save(pSaveDialog.FileName);
                        MessageBox.Show("图片已保存");
                    }
                    if (pSaveDialog.FilterIndex == 2)
                    {
                        Size mySize = webBrowser1.Document.Window.Size;
                        Bitmap myPic = new Bitmap(mySize.Width, mySize.Height);
                        Rectangle myRec = new Rectangle(0, 0, mySize.Width, mySize.Height);
                        webBrowser1.Size = mySize;
                        webBrowser1.DrawToBitmap(myPic, myRec);
                        myPic.Save(pSaveDialog.FileName,ImageFormat.Gif);
                        MessageBox.Show("图片已保存");
                    }
                    if (pSaveDialog.FilterIndex == 3)
                    {
                        Size mySize = webBrowser1.Document.Window.Size;
                        Bitmap myPic = new Bitmap(mySize.Width, mySize.Height);
                        Rectangle myRec = new Rectangle(0, 0, mySize.Width, mySize.Height);
                        webBrowser1.Size = mySize;
                        webBrowser1.DrawToBitmap(myPic, myRec);
                        myPic.Save(pSaveDialog.FileName, ImageFormat.Png);
                        MessageBox.Show("图片已保存");
                    }
                    if (pSaveDialog.FilterIndex == 4)
                    {
                        Size mySize = webBrowser1.Document.Window.Size;
                        Bitmap myPic = new Bitmap(mySize.Width, mySize.Height);
                        Rectangle myRec = new Rectangle(0, 0, mySize.Width, mySize.Height);
                        webBrowser1.Size = mySize;
                        webBrowser1.DrawToBitmap(myPic, myRec);
                        myPic.Save(pSaveDialog.FileName, ImageFormat.Bmp);
                        MessageBox.Show("图像已保存");
                    }

                }

            }
            catch (Exception Err)
            {
                MessageBox.Show(Err.Message, "输出图片", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SZTXSX_Click(object sender, EventArgs e)
        {
            OtherForms.SheZhiTuPianShuXing shezhituxiangshuxing = new OtherForms.SheZhiTuPianShuXing();
            shezhituxiangshuxing.Show();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                webBrowser1.Document.InvokeScript("addcitylist");
            }
            if (checkBox3.Checked == false)
            {
                webBrowser1.Document.InvokeScript("removecitylist");
            }
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void BZWD_Click(object sender, EventArgs e)
        {
         
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
           
        }

        private void toolStripComboBox1_Leave(object sender, EventArgs e)
        {

        }

        private void SJGL_Click(object sender, EventArgs e)
        {
            OtherForms.SJGL shuju = new OtherForms.SJGL();
            shuju.Show();
        }

        private void KSJT_Click(object sender, EventArgs e)
        {
            MessageBox.Show("在地图上单击获得测试点");
            webBrowser1.Document.InvokeScript("addpointsew");
        }

        private void JCJT_Click(object sender, EventArgs e)
        {
            MessageBox.Show("测试点采集完成！");
            webBrowser1.Document.InvokeScript("endaddpointsew");
          //  MessageBox.Show(webBrowser1.Document.GetElementById("ceshidian").InnerText);
           // OtherForms.CeShiDian ceshidian = new OtherForms.CeShiDian();
          //  ceshidian.Show();
           // UpdateCS();
        }
        public void UpdateGPS()
        {
            treeView1.Nodes[0].Nodes.Clear();
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                TreeNode tn = new TreeNode("点" + ds.Tables[0].Rows[i][0].ToString());
                treeView1.Nodes[0].Nodes.Add(tn);
            }
            //  MessageBox.Show(ds.Tables[0].Rows.Count + "");
        }
        public void UpdateCS()
        {
            treeView1.Nodes[1].Nodes.Clear();
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn); 
            string txtCommand = "SELECT ID FROM 测试点";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                TreeNode tn = new TreeNode("测试点" + ds.Tables[0].Rows[i][0].ToString());
                treeView1.Nodes[1].Nodes.Add(tn);
            }
            //  MessageBox.Show(ds.Tables[0].Rows.Count + "");
        }
        public void GetPoints()
        {
           // treeView1.Nodes[0].Nodes.Clear();
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,经度,纬度 FROM DW Where ID >0";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds);
           // int a = ds.Tables[0].Rows.Count;
            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    b = int.Parse(ds.Tables[0].Rows[i][0].ToString());
                   //  MessageBox.Show(b+"");
                    webBrowser1.Document.GetElementById("ceshidianb").InnerText =""+b;
                    string[] pointx = ds.Tables[0].Rows[i][1].ToString().Split('E');
                  //  MessageBox.Show(pointx[0]);
                    webBrowser1.Document.GetElementById("ceshidianx").InnerText = pointx[0];
                  string[] pointy = ds.Tables[0].Rows[i][2].ToString().Split('N');
                    webBrowser1.Document.GetElementById("ceshidiany").InnerText = pointy[0];
                   webBrowser1.Document.InvokeScript("ShengChengDian");
                }
                webBrowser1.Document.InvokeScript("ShengChengXian");
            }
            else { }
        }

        public void RemovePoints()
        {
            // treeView1.Nodes[0].Nodes.Clear();
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID,经度,纬度 FROM DW Where ID >0";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds);
            // int a = ds.Tables[0].Rows.Count;
            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string[] pointx = ds.Tables[0].Rows[i][1].ToString().Split('E');
                    //  MessageBox.Show(pointx[0]);
                    webBrowser1.Document.GetElementById("ceshidianx").InnerText = pointx[0];
                    string[] pointy = ds.Tables[0].Rows[i][2].ToString().Split('N');
                    webBrowser1.Document.GetElementById("ceshidiany").InnerText = pointy[0];
                    webBrowser1.Document.InvokeScript("ShanChuDian");
                }
            }
            else { }
        }

        private void treeView1_AfterCollapse(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Text.Equals("GPS坐标点"))
            {
                RemovePoints();
            }
        }

        private void treeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            
        }

        private void GJHF_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.Parent != null)
            {
                if (treeView1.SelectedNode.Parent.Text.Equals("测试点") )
                    webBrowser1.Document.InvokeScript("guijibofang");
                if (treeView1.SelectedNode.Parent.Text.Equals("GPS坐标点"))
                    webBrowser1.Document.InvokeScript("guijibofang2");
            }
            else
            {
                if (treeView1.SelectedNode.Text.Equals("测试点"))
                    webBrowser1.Document.InvokeScript("guijibofang");
                if (treeView1.SelectedNode.Text.Equals("GPS坐标点"))
                    webBrowser1.Document.InvokeScript("guijibofang2");
            }
        }

        private void XGB_Click(object sender, EventArgs e)
        {
            OtherForms.SJGL shuju = new OtherForms.SJGL();
            shuju.Show();
        }

        private void XGD_Click(object sender, EventArgs e)
        {
            OtherForms.SJGL shuju = new OtherForms.SJGL();
            shuju.Show();
        }

        private void 继续播放ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("jixu");
        }

        private void 暂停播放ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("zanting");
        }

        private void 停止播放ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("tingzhi");
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Point pos = new Point(e.Node.Bounds.X + e.Node.Bounds.Width, e.Node.Bounds.Y + e.Node.Bounds.Height / 2);
                this.contextMenuStrip1.Show(this.treeView1, pos);
            }
        }

        private void YCGD_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.Parent.Text.Equals("GPS坐标点"))
            {
                string tex = treeView1.SelectedNode.Text;
                string result = System.Text.RegularExpressions.Regex.Replace(tex, @"[^0-9]+", "");
                // int a = int.Parse(result);
                string txt1 =
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
                string txt3 = "Select 经度,纬度 From DW Where ID =" + result;
                OleDbDataAdapter daa = new OleDbDataAdapter(txt3, conn);
                DataSet ds = new DataSet("ds");
                daa.Fill(ds);
                string[] pointx = ds.Tables[0].Rows[0][0].ToString().Split('E');
                webBrowser1.Document.GetElementById("ceshidianx").InnerText = pointx[0];
                string[] pointy = ds.Tables[0].Rows[0][1].ToString().Split('N');
                webBrowser1.Document.GetElementById("ceshidiany").InnerText = pointy[0];
                webBrowser1.Document.GetElementById("ceshidianb").InnerText = result;
                webBrowser1.Document.InvokeScript("ShanChuYiDian");

                // int a = ds.Tables[0].Rows.Count;
                string txt2 = "Delete From DW Where ID = " + result;
                conn = new OleDbConnection(txt1);
                conn.Open();
                da = new OleDbCommand();
                da.CommandText = txt2;
                da.Connection = conn;
                da.ExecuteNonQuery();
                conn.Close();
            }
         

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (treeView1.SelectedNode.Text.Equals("GPS坐标点"))
            {
                YCGD.Visible = false;
                SFDD.Visible = false;
                CXB.Visible = false;
            }else if (treeView1.SelectedNode.Text.Equals("测试点"))
            {
                YCGD.Visible = false;
                SFDD.Visible = false;
              //  DCZ.Visible = false;
                BCZ.Visible = false;
                CXB.Visible = true;
            }
            else if (treeView1.SelectedNode.Parent.Text.Equals("测试点"))
            {
                YCGD.Visible = false;
                SFDD.Visible = true;
               // DCZ.Visible = false;
                BCZ.Visible = false;
                CXB.Visible = false;
            }    
            else
            {
           //     DCZ.Visible = true;
                BCZ.Visible = true;
                YCGD.Visible = true;
                SFDD.Visible = true;
                CXB.Visible = false;
            }
        }

        private void SFDD_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.Parent.Text.Equals("GPS坐标点"))
            {
                string tex = treeView1.SelectedNode.Text;
                string result = System.Text.RegularExpressions.Regex.Replace(tex, @"[^0-9]+", "");
                // int a = int.Parse(result);
                string txt1 =
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
                string txt3 = "Select 经度,纬度 From DW Where ID =" + result;
                OleDbDataAdapter daa = new OleDbDataAdapter(txt3, conn);
                DataSet ds = new DataSet("ds");
                daa.Fill(ds);
                string[] pointx = ds.Tables[0].Rows[0][0].ToString().Split('E');
                webBrowser1.Document.GetElementById("ceshidianx").InnerText = pointx[0];
                string[] pointy = ds.Tables[0].Rows[0][1].ToString().Split('N');
                webBrowser1.Document.GetElementById("ceshidiany").InnerText = pointy[0];
                webBrowser1.Document.GetElementById("ceshidianb").InnerText = result;
                webBrowser1.Document.InvokeScript("SuoFangDaoDian");
                conn.Close();
            }
            if (treeView1.SelectedNode.Parent.Text.Equals("测试点"))
            {
                string tex = treeView1.SelectedNode.Text;
                string result = System.Text.RegularExpressions.Regex.Replace(tex, @"[^0-9]+", "");
                 int a = int.Parse(result);
                 result = "" + a;
                string txt1 =
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
                string txt3 = "Select 经度,纬度 From 测试点 Where ID =" + result;
                OleDbDataAdapter daa = new OleDbDataAdapter(txt3, conn);
                DataSet ds = new DataSet("ds");
                daa.Fill(ds);
                string[] pointx = ds.Tables[0].Rows[0][0].ToString().Split('E');
                webBrowser1.Document.GetElementById("ceshidianx").InnerText = pointx[0];
                string[] pointy = ds.Tables[0].Rows[0][1].ToString().Split('N');
                webBrowser1.Document.GetElementById("ceshidiany").InnerText = pointy[0];
                webBrowser1.Document.GetElementById("ceshidianb").InnerText = result;
                webBrowser1.Document.InvokeScript("SuoFangDaoDian");
                conn.Close();
            }


        }

        private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //webBrowser1.Document.InvokeScript("tingzhi");
            webBrowser1.Document.InvokeScript("resecttt");
            treeView1.Nodes[1].Nodes.Clear();

            string txt1 =
           "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2 = "Delete From DW Where ID <0 ";
            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            conn.Close();
            treeView1.Nodes[0].Collapse();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            webBrowser1.Document.InvokeScript("tingzhi");
            webBrowser1.Document.InvokeScript("resecttt");
            treeView1.Nodes[1].Nodes.Clear();

            string txt1 =
           "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            string txt2 = "Delete From 测试点 Where ID >0 ";
            conn = new OleDbConnection(txt1);
            conn.Open();
            da = new OleDbCommand();
            da.CommandText = txt2;
            da.Connection = conn;
            da.ExecuteNonQuery();
            conn.Close();
        }

        private void CXB_Click(object sender, EventArgs e)
        {
            OtherForms.CeShiDian ceshidian = new OtherForms.CeShiDian();
            ceshidian.Show();
            UpdateCS();
        }

        private void GSZH_Click(object sender, EventArgs e)
        {
            OtherForms.GpxShuJuZhuanHuan geshi = new OtherForms.GpxShuJuZhuanHuan();
            geshi.Show();
        }
        public void UpdateID()
        {
            string txtConn =
       "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mypath2;
            conn = new OleDbConnection(txtConn);
            string txtCommand = "SELECT ID FROM DW";
            OleDbDataAdapter da = new OleDbDataAdapter(txtCommand, conn);
            DataSet ds = new DataSet("ds");
            da.Fill(ds, "DW");
            int a = ds.Tables[0].Rows.Count;
           // MessageBox.Show(""+a);
            if (a > 0)
            {
                ID = int.Parse(ds.Tables[0].Rows[a - 1][0].ToString());
            }
          //  MessageBox.Show(""+ID);

        }

        private void CKSZ_Click(object sender, EventArgs e)
        {
            //UpdateID();
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                webBrowser1.Document.InvokeScript("AddMapType");
            }
            if (checkBox1.Checked == false)
            {
                webBrowser1.Document.InvokeScript("RemoveMapType");
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                webBrowser1.Document.InvokeScript("AddBiLiChi");
            }
            if (checkBox2.Checked == false)
            {
                webBrowser1.Document.InvokeScript("RemoveBiLiChi");
            }
        }

        private void 全景ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("QuanJing");
            toolStripStatusLabel2.Text="北京市";
            toolStripStatusLabel6.Text = "116.403983";
            toolStripStatusLabel9.Text = "39.9151103";
        }

        private void 刷新ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("ShuaXin");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            treeView1.Top = button2.Height + button2.Location.Y;
            // Get the clicked button...
            Button clickedButton = (Button)sender;
            // ... and it's tabindex
            int clickedButtonTabIndex = clickedButton.TabIndex;

            // Send each button to top or bottom as appropriate
            foreach (System.Windows.Forms.Control ctl in panel1.Controls)
            {
                if (ctl is Button)
                {
                    Button btn = (Button)ctl;
                    if (btn.TabIndex > clickedButtonTabIndex)
                    {
                        if (btn.Dock != DockStyle.Bottom)
                        {
                            btn.Dock = DockStyle.Bottom;
                            // This is vital to preserve the correct order
                            btn.BringToFront();
                        }
                    }
                    else
                    {
                        if (btn.Dock != DockStyle.Top)
                        {
                            btn.Dock = DockStyle.Top;
                            // This is vital to preserve the correct order
                            btn.BringToFront();
                        }
                    }
                }
            }

            // Determine which button was clicked.
            switch (clickedButton.Text)
            {
                case "搜索定位":
                    CreateCarList();
                   // CreateOutlookList();
                    break;

                case "图层控制":
                    CreateOutlookList();
                    break;

              //  case "Zip Files":
               //     CreateZipList();
               //     break;
            }
           // listView1.BringToFront();  // Without this, the buttons will hide the items.
        }

        private void button2_Click(object sender, EventArgs e)
        {
            treeView1.Top = button2.Height + button2.Location.Y;
            // Get the clicked button...
            Button clickedButton = (Button)sender;
            // ... and it's tabindex
            int clickedButtonTabIndex = clickedButton.TabIndex;

            // Send each button to top or bottom as appropriate
            foreach (System.Windows.Forms.Control ctl in panel1.Controls)
            {
                if (ctl is Button)
                {
                    Button btn = (Button)ctl;
                    if (btn.TabIndex > clickedButtonTabIndex)
                    {
                        if (btn.Dock != DockStyle.Bottom)
                        {
                            btn.Dock = DockStyle.Bottom;
                            // This is vital to preserve the correct order
                            btn.BringToFront();
                        }
                    }
                    else
                    {
                        if (btn.Dock != DockStyle.Top)
                        {
                            btn.Dock = DockStyle.Top;
                            // This is vital to preserve the correct order
                            btn.BringToFront();
                        }
                    }
                }
            }

            // Determine which button was clicked.
            switch (clickedButton.Text)
            {
                case "搜索定位":
                    CreateCarList();
                    break;

                case "图层控制":
                    CreateOutlookList();
                    break;

                //  case "Zip Files":
                //     CreateZipList();
                //     break;
            }
         //   listView1.BringToFront();  // Without this, the buttons will hide the items.
        }

        private void CreateCarList()
        {
         //   listView1.Items.Clear();
          //  listView1.LargeImageList = imageListCars;
            //listView1.Items.Add(tabControl1);
            tabControl1.Visible = true;
           // treeView1.Visible = false;
            treeView1.Top = button2.Height + button2.Location.Y;
        }
        private void CreateOutlookList()
        {
         //   listView1.Items.Clear();
          //  listView1.LargeImageList = imageListOutlook;
           // listView1.Items.Add("Outlook Today", 0);
          //  listView1.Items.Add("Inbox", 1);
         //   listView1.Items.Add("Calendar", 2);
          //  listView1.Items.Add("Contacts", 3);
          //  listView1.Items.Add("Tasks", 4);
          //  listView1.Items.Add("Deleted Items", 5);
            tabControl1.Visible = false;
         //  treeView1.Visible = true;
            treeView1.Top = button2.Height + button2.Location.Y;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                webBrowser1.Document.GetElementById("cityname").InnerText = textBox1.Text;
                webBrowser1.Document.InvokeScript("theLocation");
              //  this.Close();
            }
            else
                MessageBox.Show("请输入有效地地名！");
            textBox1.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "" || textBox3.Text == "")
            {
                MessageBox.Show("请输入有效的地名！");
            }
            else
            {
                webBrowser1.Document.GetElementById("point1").InnerText = textBox2.Text;
                webBrowser1.Document.GetElementById("point2").InnerText = textBox3.Text;
                webBrowser1.Document.InvokeScript("ssss");
               // webBrowser1.Document.InvokeScript("jisuanluchengshijian");
               // this.Close();
                OtherForms.QRDH queren = new OtherForms.QRDH();
                queren.Show();
            }
            textBox2.Text = "";
            textBox3.Text = "";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox4.Text == "" || textBox5.Text == "")
            {
                MessageBox.Show("请输入有效的经纬度！");
            }
            else
            {
                webBrowser1.Document.GetElementById("longittude").InnerText = textBox4.Text;
                webBrowser1.Document.GetElementById("latitude").InnerText = textBox5.Text;
                webBrowser1.Document.InvokeScript("theLocation2");
                textBox4.Text = "";
                textBox5.Text = "";
                //this.Close();
            }
        }

        private void 定位我的位置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OtherForms.MyPosition mp = new OtherForms.MyPosition();
            mp.Show();
        }

        private void 初始化我的位置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("SFGH");
        }

    }
}
 