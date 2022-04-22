using MySql.Data.MySqlClient;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace excelScreen
{
    public partial class Form1 : CCWin.Skin_DevExpress
    {
        public static string[] arrayData = ConfigurationManager.AppSettings["arrayData"].Split(',');
        public static readonly Dictionary<string, string> TextSwitch = new Dictionary<string, string>();
        private static readonly log4net.ILog _log = log4net.LogManager.GetLogger("日志");
        public Form1()
        {
            InitializeComponent();
        }

        private void skinButton6_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "EXCEl文件|*.xlsx|所有文件(*.*)|*.*";
            //循环遍历所有选中的checkbox
            List<string> arrCheck = new List<string> { };
            for (int i = 0; i < uiTabControl1.TabPages.Count; i++)
            {
                foreach (var item in uiTabControl1.TabPages[i].Controls)
                {
                    if (item is HZH_Controls.Controls.UCCheckBox)
                    {
                        HZH_Controls.Controls.UCCheckBox ucCheckB = ((HZH_Controls.Controls.UCCheckBox)item);
                        if (ucCheckB.Checked)
                        {
                            string[] checkname = ucCheckB.Tag.ToString().Split(':');
                            arrCheck.Add(checkname[1]);
                        }
                    }
                }
            }
            
          
            if (arrCheck.Count>0)
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    //开启线程异步导入
                    TaskFactor.NewTask(SelectFile, new TaskFactor.TaskPara()
                    {
                        Paras = ofd.FileNames,
                        Checknum = arrCheck,
                        Callback = SelectFileCallBack
                    });

                }
            }
            else
            {
                CCWin.MessageBoxEx.Show("请选择结构类型");
            }
           
        }

        public static DataTable UniteDataTable(DataTable dt1, DataTable dt2, string DTName)
        {
            DataTable dt3 = dt1.Clone();
            for (int i = 0; i < dt2.Columns.Count; i++)
                dt3.Columns.Add(dt2.Columns[i].ColumnName);
            object[] obj = new object[dt3.Columns.Count];
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                dt1.Rows[i].ItemArray.CopyTo(obj, 0);
                dt3.Rows.Add(obj);
            }
            if (dt1.Rows.Count >= dt2.Rows.Count)
            {
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    for (int j = 0; j < dt2.Columns.Count; j++)
                        dt3.Rows[i][j + dt1.Columns.Count] = dt2.Rows[i][j].ToString();
                }
            }
            else
            {
                DataRow dr3;
                for (int i = 0; i < dt2.Rows.Count - dt1.Rows.Count; i++)
                {
                    dr3 = dt3.NewRow();
                    dt3.Rows.Add(dr3);
                }
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    for (int j = 0; j < dt2.Columns.Count; j++)
                        dt3.Rows[i][j + dt1.Columns.Count] = dt2.Rows[i][j].ToString();
                }
            }
            dt3.TableName = DTName; //设置DT的名字
            return dt3;
        }


        public void SelectFile(object obj)
        {
            TaskFactor.TaskPara tt = (TaskFactor.TaskPara)obj;
            string[] ofd = (string[])tt.Paras;
            string[] lists = ((List<string>)tt.Checknum).ToArray();
          //选择多个表格和每个表拆分数据
            DataSet d3 = new DataSet();
            bool IsFirst = true;
            foreach (string Wenjian in ofd)
            {
                DataTable dt = new DataTable();
                DataTable dt1 = new DataTable();
                for (int i = 0; i < lists.Length; i++)
                {
                    DataTable d1 = NPOIExcel.ExcelToTable2(Wenjian, lists[i]);
                    DataTable d2 = NPOIExcel.ExcelToTable2(Wenjian, null, "A4:B23");
                    dt1 = UniteDataTable(d2, d1, lists[i]);
                    if (!IsFirst)
                    {
                        for (int j = 0; j < d3.Tables.Count; j++)
                        {
                            if (d3.Tables[j].TableName== dt1.TableName)
                            {
                                d3.Tables[j].Merge(dt1);
                            }
                        }
                    }
                    else
                    {
                        d3.Tables.Add(dt1);
                    }
                   
                }
                IsFirst = false;
            }
           
            tt.Invoke(d3);
        }

        public void SelectFileCallBack(object obj)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action<object>(SelectFileCallBack), obj);
                }
                else
                {
                    for (int j = 0; j < uiTabControl1.TabPages.Count; j++)
                    {
                        foreach (var item in uiTabControl1.TabPages[j].Controls)
                        {

                            if (item is DataGridView)
                            {
                                DataTable dt = new DataTable();
                                dt.Rows.Clear();
                                ((DataGridView)item).DataSource = dt;
                                for (int i = 0; i < ((DataSet)obj).Tables.Count; i++)
                                {
                                    string[] tag = ((DataGridView)item).Tag.ToString().Split(':');
                                    if (tag[1] == ((DataSet)obj).Tables[i].TableName)
                                    {
                                        ((DataGridView)item).DataSource = ((DataSet)obj).Tables[i];

                                    }
                                }

                            }
                        }
                    }
                   
                    
                }
            }
            catch (Exception ex)
            {
                CCWin.MessageBoxEx.Show(ex.Message);
            }
        }


        private void ucBtnExt1_BtnClick(object sender, EventArgs e)
        {
            //循环遍历所有选中的checkbox
            List<string> arrCheck = new List<string> { };
            for (int i = 0; i < uiTabControl1.TabPages.Count; i++)
            {
                foreach (var item in uiTabControl1.TabPages[i].Controls)
                {
                    if (item is HZH_Controls.Controls.UCCheckBox)
                    {
                        HZH_Controls.Controls.UCCheckBox ucCheckB = ((HZH_Controls.Controls.UCCheckBox)item);
                        string _ucCheckB = ucCheckB.Tag.ToString().Split(':')[0];
                        if (ucCheckB.Checked)
                        {
                            foreach (var item2 in uiTabControl1.TabPages[i].Controls)
                            {
                                if (item2 is DataGridView)
                                {
                                    DataGridView DbView = ((DataGridView)item2);
                                    string _dbview = DbView.Tag.ToString().Split(':')[0];
                                    if (_dbview == _ucCheckB)
                                    {
                                        DataTable dt = DbView.DataSource as DataTable;
                                        TaskFactor.NewTask(SaveData, new TaskFactor.TaskPara()
                                        {
                                            Paras = new object[] { dt, ucCheckB.Tag.ToString().Split(':')[2], ucCheckB.Tag.ToString().Split(':')[0] },
                                            Callback = SaveDataCallBack
                                        });
                                    }
                                }
                            }

                        }

                    }
                }
            }
            
          
        }
        public void SaveData(object obj)
        {
            TaskFactor.TaskPara tt = (TaskFactor.TaskPara)obj;
            try
            {
                string str = ConfigurationManager.AppSettings[tt.Paras[1].ToString()];
                string name = tt.Paras[2].ToString();
                DBHelper db = new DBHelper();
                DataTable dt = (DataTable)tt.Paras[0];
                int a = db.UpSqlData(str, dt);

                tt.Invoke(new object[] { a, name });
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void SaveDataCallBack(object obj)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<object>(SaveDataCallBack), obj);
            }
            else
            {
                int a = int.Parse(((object[])obj)[0].ToString());
                string tabname = ((object[])obj)[1].ToString();
                if (a > 0)
                {
                    CCWin.MessageBoxEx.Show(tabname + "  导入成功数据条数： " + a);
                }
            }
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            //NPOIExcel.IsMergeCell()
        }

        public static void Initparas(string path,string txt)
        {
            if (File.Exists(path))
            {
                File.WriteAllText(path, txt);
                //TextSwitch.Clear();
                //string[] tts = File.ReadAllLines(path, Encoding.UTF8);
                //foreach (string tt in tts)
                //{
                //    if (string.IsNullOrEmpty(tt)) continue;
                //    string[] p = tt.Split('&');
                //    if (p.Length < 2) continue;
                //    if (TextSwitch.ContainsKey(p[0]))
                //    {
                //        TextSwitch[p[0]] = p[1];
                //    }
                //    else
                //    {
                //        TextSwitch.Add(p[0], p[1]);
                //    }
                //}
            }
            else
            {
                File.WriteAllText(path, txt);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Start();
            //初始化操作
            loadFunction();
        }

        public void loadFunction()
        {
            //创建绑定关系文件
            string[] _arrayData = ConfigurationManager.AppSettings["arrayData"].Split(',');
            for (int k = 0; k < _arrayData.Length; k++)
            {
                string _fileName = _arrayData[k].Split(':')[0];
                string txt = _arrayData[k].ToString();
                uiTabControl1.TabPages[k].Text = _fileName;
                Initparas(@"绑定关系文件\" + _fileName + ".txt", txt);
            }
            //循环给指定控件tag赋值
            int controlCount = 0;
            for (int i = 0; i < uiTabControl1.TabPages.Count; i++)
            {
                foreach (var item in uiTabControl1.TabPages[i].Controls)
                {

                    if (item is HZH_Controls.Controls.UCCheckBox)
                    {
                        controlCount++;
                    }
                }
            }
            
            for (int i = 1; i <= controlCount; i++)
            {
                ((HZH_Controls.Controls.UCCheckBox)this.Controls.Find("ucCheckBox" + i, true)[0]).Tag = _arrayData[i-1];
                ((HZH_Controls.Controls.UCCheckBox)this.Controls.Find("ucCheckBox" + i, true)[0]).TextValue = _arrayData[i-1].Split(':')[0];
                this.Controls.Find("uiDataGridView" + i, true)[0].Tag = _arrayData[i-1];
            }
           
        }

        public void Log(string txt)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(Log), txt);
                return;
            }
            else
            {
               // listBox1.Items.Add("【" + DateTime.Now.ToString("HH:mm:ss") + "】" + txt);
                _log.Info(txt);
            }
        }
        public static void WriteLog(Exception ex)
        {
            _log.Error(ex);
            string LogAddress = "";
            //如果日志文件为空，则默认在Debug目录下新建 YYYY-mm-dd_Log.log文件
            if (LogAddress == "")
            {
                LogAddress = Environment.CurrentDirectory + '\\' +
                    DateTime.Now.Year + '-' +
                    DateTime.Now.Month + '-' +
                    DateTime.Now.Day + "_Log.log";
            }
            //把异常信息输出到文件
            StreamWriter fs = new StreamWriter(LogAddress, true);
            fs.WriteLine("当前时间：" + DateTime.Now.ToString());
            fs.WriteLine("异常信息：" + ex.Message);
            fs.WriteLine("异常对象：" + ex.Source);
            fs.WriteLine("调用堆栈：\n" + ex.StackTrace.Trim());
            fs.WriteLine("触发方法：" + ex.TargetSite);
            fs.WriteLine();
            fs.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            SetTapTips();
        }

        public void SetTapTips()
        {
            for (int i = 0; i < uiTabControl1.TabPages.Count; i++)
            {
                foreach (var item in uiTabControl1.TabPages[i].Controls)
                {
                    if (item is HZH_Controls.Controls.UCCheckBox)
                    {
                        if (((HZH_Controls.Controls.UCCheckBox)item).Checked)
                        {
                            uiTabControl1.SetTipsText(uiTabControl1.TabPages[i], "1");
                        }
                        else
                        {
                            uiTabControl1.SetTipsText(uiTabControl1.TabPages[i], "");
                        }

                    }
                }
            }

        }
    }
}
