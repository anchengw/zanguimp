using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOIExcelOpa;
using RichTextLog;

namespace zanguimp
{
    public partial class Form1 : Form
    {
        DataTable excelDt = null;
        string connStr = @"Data Source = 192.168.1.15; Initial Catalog = UFDATA_003_2016; User ID = sa; Password=knfz@2013";
        Dictionary<string, string> autoidList = new Dictionary<string, string>();
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 检查数据
        /// </summary>
        private bool checkData()
        {
            try
            {
                LogRichTextBox.logMesg("正在检查导入数据的完整性。。。");
                int i = 1;
                foreach (DataRow dr in excelDt.Rows)
                {
                    if (string.IsNullOrEmpty(dr["金额"].ToString()) || string.IsNullOrEmpty(dr["数量"].ToString()) || string.IsNullOrEmpty(dr["单据号"].ToString()) || string.IsNullOrEmpty(dr["存货编码"].ToString()))
                    {
                        LogRichTextBox.logMesg("第" + i.ToString() + "行记录的单据号、存货编码、金额和数量等关键列不能为空！数据完整性检查失败！请核对要导入的EXCEL文件。", 1);
                        return false;
                    }
                    i++;
                }
                LogRichTextBox.logMesg("数据完整,可以导入。");
            }
            catch (Exception e)
            {
                LogRichTextBox.logMesg("数据完整性检查程序出错！错误原因：" + e.ToString(), 2);
                return false;
            }
            return true;
        }
        /// <summary>
        /// 提交数据库
        /// </summary>
        public void commitDatabase()
        {
            string autoid = "";
            Double jine, num;
            string price = "";
            autoidList.Clear();

            foreach (DataRow dr in excelDt.Rows)
            {
                string sqlstr = @"select autoid from rdrecord01 as RdRecord inner join rdrecords01 as RdRecords on rdrecord.id=rdrecords.id  where RdRecord.cCode = {0} and RdRecords.cInvCode = '{1}';";
                string upsql = @"Update rdrecords01 set iUnitCost = {0}, faCost = {0}, iPrice = {1}, iAPrice = {1} Where autoid = {2}";

                sqlstr = string.Format(sqlstr, dr["单据号"].ToString(), dr["存货编码"].ToString());
                autoid = (Sqlhelper.DbHelperSQL.GetSingle(sqlstr)).ToString();
                num = Convert.ToDouble(dr["数量"]);
                jine = Convert.ToDouble(dr["金额"]);
                price = (jine / num).ToString("f10"); //0.3200000000
                if (string.IsNullOrEmpty(autoid))
                {
                    throw new Exception("获取U8系统的AUTOID失败！导入终止！请稍后尝试或联系系统工程师解决此问题！");
                }
                upsql = string.Format(upsql, price, jine, autoid);
                autoidList.Add(autoid, upsql);
            }
            if (autoidList.Count == 0)
                throw new Exception("SQL语句字典为空！");
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                SqlTransaction tx = conn.BeginTransaction();
                cmd.Transaction = tx;
                try
                {
                    foreach (var item in autoidList)
                    {
                        string strsql = item.Value;
                        int id = int.Parse(item.Key);
                        if (strsql.Trim().Length > 1)
                        {
                            cmd.CommandText = strsql;
                            cmd.ExecuteNonQuery();
                            //执行存储过程
                            SqlCommand sqlCmd = new SqlCommand("Pu_WBRkdCostPrice", conn);
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            /*
                            sqlCmd.Parameters.Add(new SqlParameter("@sRdsID", SqlDbType.Int));
                            sqlCmd.Parameters["@sRdsID"].Value = id;

                            sqlCmd.Parameters.Add(new SqlParameter("@QuanPoint", SqlDbType.Int));
                            sqlCmd.Parameters["@QuanPoint"].Value = 2;

                            sqlCmd.Parameters.Add(new SqlParameter("@PricePoint", SqlDbType.Int));
                            sqlCmd.Parameters["@PricePoint"].Value = 2;

                            sqlCmd.Parameters.Add(new SqlParameter("@NumPoint", SqlDbType.Int));
                            sqlCmd.Parameters["@NumPoint"].Value = 2;

                            sqlCmd.Parameters.Add(new SqlParameter("@iError", SqlDbType.Int));
                            sqlCmd.Parameters["@iError"].Value = DBNull.Value;
                            */
                            SqlParameter[] param = new SqlParameter[]
                            {
                                   new SqlParameter("@sRdsID",id),
                                   new SqlParameter("@QuanPoint",2),
                                   new SqlParameter("@PricePoint",2),
                                   new SqlParameter("@NumPoint", 2),
                                   new SqlParameter("@iError",SqlDbType.Int,4)
                            };
                            param[4].Value = DBNull.Value;
                            param[4].Direction = ParameterDirection.Output;
                            //param[4].Direction = ParameterDirection.ReturnValue;
                            foreach (SqlParameter parameter in param)
                            {
                                sqlCmd.Parameters.Add(parameter);
                            }
                            sqlCmd.Transaction = tx;
                            sqlCmd.ExecuteNonQuery();
                            object obj = sqlCmd.Parameters["@iError"].Value;
                            if (obj.ToString() == "-1") //出错
                            {
                                throw new Exception("存储过程错误!");
                            }
                        }
                    }
                    tx.Commit();
                }
                catch (System.Data.SqlClient.SqlException E)
                {
                    tx.Rollback();
                    throw new Exception("导入出错，已进行回滚。原因：" + E.ToString());
                }
            }
        }
        private void openExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = ofd.FileName;
                try
                {
                    excelDt = NPOIExcel.readExcel(strFileName, true);
                    int index = excelDt.Rows.Count - 1;
                    string heji = excelDt.Rows[index]["单据日期"].ToString();
                    if (heji.Equals("合计"))
                        excelDt.Rows.RemoveAt(index);
                    LogRichTextBox.logMesg("文件『" + strFileName + "』读取成功！共读取【" + excelDt.Rows.Count.ToString() + "】条记录。");
                }
                catch(Exception ex)
                {
                    LogRichTextBox.logMesg("文件『" + strFileName + "』读取失败！失败原因为：" + ex.ToString(),2);
                    return;
                }
                impU8.Enabled = checkData();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LogRichTextBox.richTextBoxRemote = this.richTextBox1;
            impU8.Enabled = false;
            //this.ControlBox = false; //去所有按钮
            this.MaximizeBox = false;//去大最化按钮
            this.MinimizeBox = false;//去最小化按钮
            this.BackColor = Color.FromArgb(194, 216, 240);
            this.ShowIcon = false;

            Sqlhelper.DbHelperSQL.setConnectStr(connStr);
        }

        private void impU8_Click(object sender, EventArgs e)
        {
          
            try
            {
                LogRichTextBox.logMesg("请等候，开始导入U8系统。。。");
                impU8.Enabled = false;
                commitDatabase();
                LogRichTextBox.logMesg("数据导入成功！共导入【" + autoidList.Count.ToString() + "】条记录！");
                excelDt.Rows.Clear();//导入成功，清空数据
            }
            catch(Exception ex)
            {
                impU8.Enabled = true;
                LogRichTextBox.logMesg("程序出错！错误原因为：" + ex.ToString(), 2);
            }
        }

        private void 清除日志ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogRichTextBox.LogClear();
        }
    }
}
