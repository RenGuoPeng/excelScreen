
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Sunny.UI;

namespace SelectData
{
    public partial class SelectData : UIForm
    {
        public static readonly Dictionary<string,string> BtnText = new Dictionary<string, string>();
        public static string sqlData = "";
        public static int pageCount = 0;//一共多少数据
        public static StringBuilder strb = new StringBuilder();
        public SelectData()
        {
            InitializeComponent();
        }
        DataSet ds = new DataSet();
        private void SelectData_Load(object sender, EventArgs e)
        {
            ReadBtn();
            uiRadioButtonGroup1.SelectedIndex = 0;
            uiPagination1.PageSize = int.Parse(uiComboBox1.Text);
            //
        }

        public void shiti()
        {
            List<object> lstSource = new List<object>();
            DataSet ds = new DataSet();
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                h_ave_0_11 model = new h_ave_0_11()
                {
                    _40CH_AX = ds.Tables[0].Rows[i]["40CH_AX"].ToString(),
                    _40CH_BD_0_5 = ds.Tables[0].Rows[i]["40CH_BD_0.5"].ToString(),
                    _40CH_BD_1 = ds.Tables[0].Rows[i]["40CH_BD_1"].ToString(),
                    _40CH_BD_20 = ds.Tables[0].Rows[i]["40CH_BD_20"].ToString(),
                    _40CH_BD_3 = ds.Tables[0].Rows[i]["40CH_BD_3"].ToString(),
                    _40CH_IL = ds.Tables[0].Rows[i]["40CH_IL"].ToString(),
                    _40CH_NX = ds.Tables[0].Rows[i]["40CH_NX"].ToString(),
                    _40CH_offset = ds.Tables[0].Rows[i]["40CH_offset"].ToString(),
                    _40CH_PDL = ds.Tables[0].Rows[i]["40CH_PDL"].ToString(),
                    _40CH_ripple = ds.Tables[0].Rows[i]["40CH_ripple"].ToString(),
                    _40CH_TX = ds.Tables[0].Rows[i]["40CH_TX"].ToString(),
                    _40CH工作波段 = ds.Tables[0].Rows[i]["40CH工作波段"].ToString(),
                    _40CH工作通道 = ds.Tables[0].Rows[i]["40CH工作通道"].ToString(),
                    _48CH_AX = ds.Tables[0].Rows[i]["48CH_AX"].ToString(),
                    _48CH_BD_0_5 = ds.Tables[0].Rows[i]["48CH_BD_0.5"].ToString(),
                    _48CH_BD_1 = ds.Tables[0].Rows[i]["48CH_BD_1"].ToString(),
                    _48CH_BD_20 = ds.Tables[0].Rows[i]["48CH_BD_20"].ToString(),
                    _48CH_BD_3 = ds.Tables[0].Rows[i]["48CH_BD_3"].ToString(),
                    _48CH_IL = ds.Tables[0].Rows[i]["48CH_IL"].ToString(),
                    _48CH_NX = ds.Tables[0].Rows[i]["48CH_NX"].ToString(),

                    _48CH_offset = ds.Tables[0].Rows[i]["48CH_offset"].ToString(),
                    _48CH_PDL = ds.Tables[0].Rows[i]["48CH_PDL"].ToString(),
                    _48CH_ripple = ds.Tables[0].Rows[i]["48CH_ripple"].ToString(),
                    _48CH_TX = ds.Tables[0].Rows[i]["48CH_TX"].ToString(),
                    _48CH工作波段 = ds.Tables[0].Rows[i]["48CH工作波段"].ToString(),
                    _48CH工作通道 = ds.Tables[0].Rows[i]["48CH工作通道"].ToString(),
                    chip_code = ds.Tables[0].Rows[i]["芯片编号"].ToString(),
                    wafer_code = ds.Tables[0].Rows[i]["晶圆编号"].ToString(),
                };
                lstSource.Add(model);
            }
        }

        private void uiPagination1_PageChanged(object sender, object pagingSource, int pageIndex, int count)
        {
            //未连接数据库，通过模拟数据来实现
            //一般通过ORM的分页去取数据来填充
            //pageIndex：第几页，和界面对应，从1开始，取数据可能要用pageIndex - 1
            //count：单页数据量，也就是PageSize值
            //List<Data> data = new List<Data>();
            if (uiPagination1.TotalCount!=0)
            {
                int sqlpage = (pageIndex - 1) * count;
                int sqlpagecount = count;
                DBHelper db = new DBHelper();
                GetDataString(sqlpage.ToString(), count.ToString());
            }

            //uiDataGridView1.DataSource = ds.Tables[0];
            //uiDataGridViewFooter1.Clear();
            //uiDataGridViewFooter1["Column1"] = "合计：";
            //uiDataGridViewFooter1["Column2"] = "Column2_" + pageIndex;
            //uiDataGridViewFooter1["Column3"] = "Column3_" + pageIndex;
            //uiDataGridViewFooter1["Column4"] = "Column4_" + pageIndex;
        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem delete = sender as ToolStripMenuItem;
            UIContextMenuStrip menu = delete.Owner as UIContextMenuStrip;
            UIButton ubn= menu.SourceControl as UIButton;
            uiFlowLayoutPanel1.Remove(ubn);
            BtnText.Remove(ubn.Name);
        }

        private void uiButton1_Click(object sender, EventArgs e)
        {
            string value = "请输入字符串";
            if (this.InputStringDialog(ref value, true, "请输入条件：",uiStyleManager1.Style, true))
            {
                Creatbutton(value, value);
            }
        }
        public void Creatbutton(string buttonName, string buttonText)
        {
            UIButton ub = new UIButton();
            ub.Name = buttonName;
            ub.Text = buttonText;
            ub.ContextMenuStrip = uiContextMenuStrip1;
            ub.AutoSize = true;
            uiFlowLayoutPanel1.Controls.Add(ub);
            if (!BtnText.ContainsKey(ub.Name))
            {
                BtnText.Add(buttonName, buttonText);
            }
          
        }

        public static void Initparas(string path, string txt)
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

        private void uiButton2_Click(object sender, EventArgs e)
        {
            StringBuilder strbu = new StringBuilder();
            foreach (var item in this.uiFlowLayoutPanel1.Controls["FlowLayoutPanel"].Controls)
            {
                if (item is UIButton)
                {
                    strbu.AppendLine(((UIButton)item).Text);
                }
                
            }
            Debug.WriteLine(strbu);
            Initparas(@"查询条件\查询条件.txt", strbu.ToString());
        }
        public void ReadBtn()
        {
            BtnText.Clear();
            string[] tts = File.ReadAllLines(@"查询条件\查询条件.txt", Encoding.UTF8);
            foreach (string tt in tts)
            {
                if (string.IsNullOrEmpty(tt)) continue;
                Creatbutton(tt, tt);
               
            }
        }

        private void 修改ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem delete = sender as ToolStripMenuItem;
            UIContextMenuStrip menu = delete.Owner as UIContextMenuStrip;
            UIButton ubn = menu.SourceControl as UIButton;
            string value = ubn.Text;
            if (this.InputStringDialog(ref value, true, "请输入修改的规则：", uiStyleManager1.Style, true))
            {
               
                BtnText.Remove(ubn.Text);
                ubn.Name = value;
                ubn.Text = value;
               
                BtnText.Add(ubn.Name, ubn.Text);
            }
           
        }

        private void uiButton3_Click(object sender, EventArgs e)
        {
            DBHelper db = new DBHelper();
            string tabName = uiRadioButtonGroup1.Items[uiRadioButtonGroup1.SelectedIndex].ToString();//数据库名
            string sqlpagecount= uiComboBox1.Text;//一页显示多少数据
            string sqlpage = "0";//第几页
            //int pageCount = 0;//一共多少数据
            foreach (var item in BtnText.Keys)
            {
                strb.Append(" and " + BtnText[item] + "");
            }
            sqlData = "select *  from `" + tabName + "` where id > (select id from `" + tabName + "` where 1=1 " + strb + " order by id limit {0}, 1) " + strb + " limit {1}";
            
            pageCount = db.select("select count(*) from `" + tabName + "` where 1=1 " + strb + ";");
            GetDataString(sqlpage, sqlpagecount);
            uiPagination1.ActivePage = 1;
        }

        public void GetDataString(string sqlpage,string sqlpagecount)
        {
            DBHelper db = new DBHelper();
            DataSet ds = new DataSet();
            ds = db.select1(string.Format(sqlData, sqlpage, sqlpagecount));
            uiPagination1.TotalCount = pageCount;
            uiPagination1.PageSize = int.Parse(uiComboBox1.Text);
            if (ds.Tables.Count> 0)
            {
                uiDataGridView1.DataSource = ds.Tables[0];
            }
        }

        private void uiButton4_Click(object sender, EventArgs e)
        {
            DataTable d1 = uiDataGridView1.DataSource as DataTable;
            if (d1 != null)
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "EXCEl文件|*.xlsx";       //设置文件类型
                dialog.FileName = DateTime.Now.ToShortDateString();   //设置默认文件名
                dialog.DefaultExt = ".xlsx";                              //设置默认格式（可以不设）
                dialog.AddExtension = true;                             //设置自动在文件名中添加扩展名
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    NPOIExcel.TableToExcel(d1, dialog.FileName);
                }
            }
          
        }
    }
}
