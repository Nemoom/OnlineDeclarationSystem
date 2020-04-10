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

namespace OnlineDeclarationSystem
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        struct StaffInfo { public string name; public string subDept;}
        DataSet DataSet_StaffList = new DataSet();
        List<StaffInfo> StaffList = new List<StaffInfo>();
        private void Form1_Load(object sender, EventArgs e)
        {
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
    }
}
