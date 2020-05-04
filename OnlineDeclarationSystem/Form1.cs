using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Reflection;
using log4net.Config;
using log4net;

namespace OnlineDeclarationSystem
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static void InitLog4Net()
        {
            var logCfg = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "log4net.config");
            XmlConfigurator.ConfigureAndWatch(logCfg);
        }

        public ILog logger;
        struct StaffInfo { public string name; public string subDept;}
        DataSet DataSet_StaffList = new DataSet();
        List<StaffInfo> StaffList = new List<StaffInfo>();
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = this.Text + "   V" + Assembly.GetExecutingAssembly().GetName().Version + "";
            DataSet_StaffList = ExcelToDataSet("FEC Working Hour for CTO DE.xlsx");
            DataTable dt_Name = DataSet_StaffList.Tables[0].DefaultView.ToTable(true, "Name");
            cBox_Name.DataSource = ItemList2StringArray(dt_Name);
            for (int i = 0; i < DataSet_StaffList.Tables[0].Rows.Count; i++)
            {
                StaffList.Add(new StaffInfo { name = DataSet_StaffList.Tables[0].Rows[i][1].ToString(), 
                    subDept = DataSet_StaffList.Tables[0].Rows[i][3].ToString() });
            }
            DataTable dt_SubDept = DataSet_StaffList.Tables[0].DefaultView.ToTable(true, "Sub-department");
            cBox_SubDept.DataSource = ItemList2StringArray(dt_SubDept);
            cBox_Name.SelectedIndex = -1;
            cBox_SubDept.SelectedIndex = -1;
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker1.MinDate = new DateTime(DateTime.Today.Year, 1, 1);
            switch (DateTime.Today.DayOfWeek)
            {
                case DayOfWeek.Friday:
                    dateTimePicker1.MaxDate = DateTime.Today.AddDays(9);
                    break;
                case DayOfWeek.Monday:
                    dateTimePicker1.MaxDate = DateTime.Today.AddDays(13);
                    break;
                case DayOfWeek.Saturday:
                    dateTimePicker1.MaxDate = DateTime.Today.AddDays(8);
                    break;
                case DayOfWeek.Sunday:
                    dateTimePicker1.MaxDate = DateTime.Today.AddDays(7);
                    break;
                case DayOfWeek.Thursday:
                    dateTimePicker1.MaxDate = DateTime.Today.AddDays(10);
                    break;
                case DayOfWeek.Tuesday:
                    dateTimePicker1.MaxDate = DateTime.Today.AddDays(12);
                    break;
                case DayOfWeek.Wednesday:
                    dateTimePicker1.MaxDate = DateTime.Today.AddDays(11);
                    break;
                default:
                    break;
            }
            logger = LogManager.GetLogger(typeof(Program));
        }

        /// <summary>
        /// 利用excelFileName生成OleDbConnection，可用来赋值全局变量OleDbConnection
        /// </summary>
        /// <param name="pathName">完整的文件路径</param>
        /// <returns></returns>
        public static OleDbConnection GetConnection(string pathName)
        {
            string strConn = string.Empty;
            //string pathName = @"C:\Users\Public\Music\" + excelFileName;
            FileInfo file = new FileInfo(pathName);
            if (!file.Exists) { throw new Exception("文件不存在"); }
            string extension = file.Extension.ToLower();
            switch (extension)
            {
                case ".xls":
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    break;
                case ".xlsx":
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                    break;
                default:
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    break;
            }
            //链接Excel
            OleDbConnection cnnxls = new OleDbConnection(strConn);
            return cnnxls;
        }

        /// <summary>
        /// C#中获取Excel文件的第一个表名
        /// Excel文件中第一个表名的缺省值是Sheet1$, 但有时也会被改变为其他名字. 如果需要在C#中使用OleDb读写Excel文件, 就需要知道这个名字是什么. 以下代码就是实现这个功能的:
        /// </summary>
        /// <param name="pathName"></param>
        /// <returns></returns>
        public static string GetExcelFirstTableName(string pathName)
        {
            //string pathName = @"C:\Users\Public\Music\" + excelFileName;
            string tableName = null;
            if (File.Exists(pathName))
            {
                OleDbConnection conn_Cur = null;
                using (conn_Cur = GetConnection(pathName))
                {
                    conn_Cur.Open();
                    DataTable dt = conn_Cur.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    tableName = dt.Rows[0][2].ToString().Trim();
                }
            }
            return tableName;
        }

        /// <summary>
        /// 把数据从Excel装载到DataSet
        /// </summary>
        /// <param name="pathName">带路径的Excel文件名</param>
        /// <param name="sheetName">工作表名</param>
        /// <param name="tbContainer">将数据存入的DataSet</param>
        /// <returns></returns>
        public DataSet ExcelToDataSet(string pathName, string sheetName = null)
        {
            DataSet ds = new DataSet();
            string strConn = string.Empty;
            
            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = GetExcelFirstTableName(pathName);
            }
            else if (!sheetName.EndsWith("$"))
            {
                sheetName = sheetName + "$";
            }
            if (!File.Exists(pathName)) { throw new Exception("文件不存在"); }

            //读取Excel里面有 表Sheet1
            OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}]", sheetName), GetConnection(pathName));
            //将Excel里面有表内容装载到内存表中！
            oda.Fill(ds);
            oda.Dispose();
            return ds;
        }

        public string[] ItemList2StringArray(DataTable mItemList)
        {
            string[] arrayString = new string[mItemList.Rows.Count];
            for (int i = 0; i < mItemList.Rows.Count; i++)
            {
                arrayString[i] = (string)mItemList.Rows[i].ItemArray[0];
            }
            return arrayString;
        }

        public DataSet SelectCMD(string sqlStr, string pathName, string sheetName = null)
        {
            DataSet ds = new DataSet();
            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = GetExcelFirstTableName(pathName);
            }
            else if (!sheetName.EndsWith("$"))
            {
                sheetName = sheetName + "$";
            }
            OleDbConnection cnnxls = GetConnection(pathName);
            OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}]", sheetName) + sqlStr, cnnxls);
            //将Excel里面有表内容装载到内存表中！
            oda.Fill(ds);
            oda.Dispose();
            cnnxls.Close();
            return ds;
        }

        private void cBox_Name_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string sqlStr = "";
            if (cBox_Name.SelectedValue!=null)
            {
                for (int i = 0; i < StaffList.Count; i++)
                {
                    if (StaffList[i].name == cBox_Name.SelectedValue.ToString())
                    {
                        for (int j = 0; j < cBox_SubDept.Items.Count; j++)
                        {
                            if (cBox_SubDept.Items[j].ToString() == StaffList[i].subDept)
                            {
                                cBox_SubDept.SelectedIndex = j;
                            }
                        }
                    }
                }
                //sqlStr = " where [Name]='" + cBox_Name.SelectedValue.ToString() + "'";
                //DataSet mResult = SelectCMD(sqlStr, "FEC Working Hour for CTO DE.xlsx");
                //if (mResult.Tables[0].Rows.Count>0)
                //{
                //    for (int i = 0; i < cBox_SubDept.Items.Count; i++)
                //    {
                //        if (cBox_SubDept.Items[i].ToString()==mResult.Tables[0].Rows[0][3].ToString())
                //        {
                //            cBox_SubDept.SelectedIndex = i;
                //        }
                //    }
                //}
            }
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Directory.Exists(@"Q:\CNGrp095\FEC&I\Working Hours Management"))
                {
                    //网盘映射有问题
                    logger.Error("网盘映射有问题");
                }
                else
                {
                    string str_aorp = "";
                    if (checkBox1.Checked && checkBox2.Checked)
                    {
                        str_aorp = "上午&下午";
                    }
                    else if (checkBox1.Checked && !checkBox2.Checked)
                    {
                        str_aorp = "上午";
                    }
                    else if (!checkBox1.Checked && checkBox2.Checked)
                    {
                        str_aorp = "下午";
                    }
                    else
                    {
                        MessageBox.Show("请选择具体时段~");
                    }
                    if (str_aorp != "")
                    {
                        try
                        {
                            if (writeCSV(cBox_SubDept.SelectedValue.ToString().Split('-')[cBox_SubDept.SelectedValue.ToString().Split('-').Length - 1] + "-" + cBox_Name.SelectedValue.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd") + " " + str_aorp,
                        cBox_Type.SelectedItem.ToString(), "\"" + textBox1.Text.Replace('"', ' ') + "\""))
                            {
                                MessageBox.Show("提交成功！");
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Fatal(ex.Message);
                            MessageBox.Show(ex.StackTrace,"Exception");
                        }

                    }
                    //if (!File.Exists(@"Q:\CNGrp095\FEC&I\Working Hours Management\" + cBox_Name.SelectedValue.ToString() + ".xlsx"))
                    //{
                    //    File.Create(@"Q:\CNGrp095\FEC&I\Working Hours Management\" + cBox_Name.SelectedValue.ToString() + ".xlsx");
                    //}
                }
            }
            catch (Exception ex2)
            {
                logger.Fatal(ex2.Message);
                MessageBox.Show(ex2.StackTrace,"Exception");
            }
        }

        public bool writeCSV(string filename,string str_date, string str_type, string str_desc)
        {
            bool b_result = false;
            bool b_write = true;
            bool b_delete = false;
            string vText = "";
            string line = string.Empty;
            const string LOG_DIR = @"Q:\CNGrp095\FEC&I\Working Hours Management";
            string csvFilePath = Path.Combine(LOG_DIR, filename + ".csv");

            if (!File.Exists(csvFilePath))
            {
                //写入表头
                using (StreamWriter csvFile = new StreamWriter(csvFilePath, true, Encoding.UTF8))
                {
                    line = "申请时间,类型,描述,提交时间,工时";
                    csvFile.WriteLine(line);
                }
            }
            else
            {
                //查重复申报
                string strValue = string.Empty;
                bool b_recordExist = false;
                using (StreamReader read = new StreamReader(csvFilePath, true))
                {
                    do
                    {
                        strValue = read.ReadLine();
                        if (strValue!=null)
                        {
                            if (strValue.Split(',')[0].Split(' ')[0] == str_date.Split(' ')[0])
                            {
                                if (strValue.Split(',')[0].Split(' ')[1] == "上午&下午" || str_date.Split(' ')[1] == "上午&下午")
                                {
                                    b_recordExist = true;
                                }
                                else
                                {
                                    if (strValue.Split(',')[0].Split(' ')[1] == str_date.Split(' ')[1])
                                    {
                                        b_recordExist = true;
                                    }
                                }
                                if (b_recordExist)
                                {
                                    if (MessageBox.Show("已存在该时段的预约，取消或覆盖", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                                    {

                                        b_write = true;
                                        b_delete = true;
                                        strValue = "";
                                    }
                                    else
                                    {
                                        b_write = false;
                                    }
                                }                                
                            }
                            if (strValue != "")
                            {
                                vText += strValue + "\r\n";
                            }                           

                        }
                    } while (strValue != null);
                }
            }
            if (b_delete)
            {
                StreamWriter vStreamWriter = new StreamWriter(csvFilePath, false, Encoding.UTF8);
                vStreamWriter.Write(vText);
                vStreamWriter.Close();
            }
            if (b_write)
            {
                using (StreamWriter csvFile = new StreamWriter(csvFilePath, true, Encoding.UTF8))
                {
                    if (str_date.Split(' ')[1] == "上午&下午")
                    {
                        line = str_date + "," + str_type + "," + str_desc + "," + DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + ",8";
                    }
                    else
                    {
                        line = str_date + "," + str_type + "," + str_desc + "," + DateTime.Now.ToString("yyyyMMdd HH:mm:ss") + ",4";
                    }
                    csvFile.WriteLine(line);
                }
                b_result = true;
            }
            return b_result;
        }
    }
}
