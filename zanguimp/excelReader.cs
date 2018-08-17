using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace ExcelComOle
{
    public class excelReader
    {
        public static string connectionString = "";//数据库连接字符串
        //连接字符串 如果第一行是数据而不是标题的话, 应该写: "HDR=No;" 否则是HDR=Yes "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=35;Extended Properties=Excel 8.0;HDR=\"{1}\";Persist Security Info=False"
        //"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可避免数据类型冲突。
        public static string connstring07 = "Provider=Microsoft.Ace.OleDb.12.0;Data Source={path};Extended Properties='Excel 12.0;HDR={HDR};IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
        public static string connstring03 = "Provider=Microsoft.JET.OLEDB.4.0;Data Source={path};Extended Properties='Excel 8.0;HDR={HDR};IMEX=1';"; //Office 07以下版本 
        #region OleDb读取Excel
        /// <summary>
        /// 将 Excel 文件转成 DataTable 后,再把 DataTable中的数据写入表Products
        /// </summary>
        /// <param name="serverMapPathExcelAndFileName"></param>
        /// <param name="excelFileRid"></param>
        /// <returns></returns>
        public static int WriteExcelToDataBase(string excelFileName)
        {
            int rowsCount = 0;
            OleDbConnection objConn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFileName + ";" + "Extended Properties=Excel 8.0;");
            objConn.Open();
            try
            {
                DataTable schemaTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                string sheetName = string.Empty;
                for (int j = 0; j < schemaTable.Rows.Count; j++)
                {
                    sheetName = schemaTable.Rows[j][2].ToString().Trim();//获取 Excel 的表名，默认值是sheet1 
                    DataTable excelDataTable = ExcelToDataTable(excelFileName, sheetName, true);
                    if (excelDataTable.Columns.Count > 1)
                    {
                        SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction);
                        sqlbulkcopy.DestinationTableName = "Products";//数据库中的表名


                        sqlbulkcopy.WriteToServer(excelDataTable);
                        sqlbulkcopy.Close();
                    }
                }
            }
            catch (SqlException ex)
            {
                throw ex;
            }
            finally
            {
                objConn.Close();
                objConn.Dispose();
            }
            return rowsCount;
        }
        /// <summary>
        /// 读取Excel
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static DataSet ExcelToDS(string Path, string sheetName)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";  //HDR=Yes;
            DataSet ds = null;
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                OleDbDataAdapter myCommand = null;
                try
                {
                    conn.Open();
                    string strExcel = "";
                    strExcel = "select * from [" + sheetName + "$]";
                    myCommand = new OleDbDataAdapter(strExcel, strConn);
                    ds = new DataSet();
                    myCommand.Fill(ds, "table1");
                }
                catch (SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    myCommand.Dispose();
                    conn.Close();
                }
                return ds;
            }
        }
        /// <summary>
        /// 将 Excel 文件转成 DataTable
        /// </summary>
        /// <param name="serverMapPathExcel">Excel文件及其路径</param>
        /// <param name="strSheetName">工作表名,如:Sheet1</param>
        /// <param name="isTitleOrDataOfFirstRow">True 第一行是标题,False 第一行是数据</param>
        /// <returns>DataTable</returns>
        public static DataTable ExcelToDataTable(string serverMapPathExcel, string strSheetName, bool isTitleOrDataOfFirstRow)
        {


            string HDR = string.Empty;//如果第一行是数据而不是标题的话, 应该写: "HDR=No;"
            if (isTitleOrDataOfFirstRow)
            {
                HDR = "YES";//第一行是标题
            }
            else
            {
                HDR = "NO";//第一行是数据
            }
            //源的定义 
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + serverMapPathExcel + ";" + "Extended Properties='Excel 8.0;HDR=" + HDR + ";IMEX=1';";
            //Sql语句
            //string strExcel = string.Format("select * from [{0}$]", strSheetName); 这是一种方法
            string strExcel = "select * from   [" + strSheetName + "]";
            //定义存放的数据表
            DataSet ds = new DataSet();
            //连接数据源
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                try
                {
                    conn.Open();
                    //适配到数据源
                    OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, strConn);

                    adapter.Fill(ds, strSheetName);
                }
                catch (System.Data.SqlClient.SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
            return ds.Tables[strSheetName];
        }
        public static DataSet ExcelToDS(string Path)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "select * from [sheet1$]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);
            ds = new DataSet();
            myCommand.Fill(ds, "table1");
            return ds;
        }
        public static DataTable ExcelToTable(string excelFilename, bool isTitleOrDataOfFirstRow)
        {
            if (!File.Exists(excelFilename))
                return null;
            string HDR = string.Empty;//如果第一行是数据而不是标题的话, 应该写: "HDR=No;"
            string connectionString = string.Empty;
            if (isTitleOrDataOfFirstRow)
            {
                HDR = "YES";//第一行是标题
            }
            else
            {
                HDR = "NO";//第一行是数据
            }
            if(Is2007Excel(excelFilename))
            {
                connectionString = connstring07.Replace("{path}", excelFilename).Replace("{HDR}",HDR);
            }
            else
            {
                connectionString = connstring03.Replace("{path}", excelFilename).Replace("{HDR}", HDR);
            }
            DataSet ds = new DataSet();
            string tableName;
            using (System.Data.OleDb.OleDbConnection connection = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                connection.Open();
                DataTable table = connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                tableName = table.Rows[0]["Table_Name"].ToString();
                string strExcel = "select * from " + "[" + tableName + "]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, connectionString);
                adapter.Fill(ds, tableName);
                connection.Close();
            }
            return ds.Tables[tableName];
        }
        #endregion
        /// <summary>
        /// 判断文件格式
        /// http://www.cnblogs.com/babycool 
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static bool Is2007Excel(string filePath)
        {
            string fileclass = "";
            string extension = "";

            extension = System.IO.Path.GetExtension(filePath);
            FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            BinaryReader reader = new BinaryReader(stream);

            // byte buffer;
            try
            {

                //buffer = reader.ReadByte();
                //fileclass = buffer.ToString();
                //buffer = reader.ReadByte();
                //fileclass += buffer.ToString();

                for (int i = 0; i < 2; i++)
                {
                    fileclass += reader.ReadByte().ToString();
                }

            }
            catch (Exception)
            {

                throw;
            }
            if (fileclass == "8075" && extension.ToLower().Equals(".xlsx"))
            {
                return true;
            }
            else
            {
                return false;
            }

            /*文件扩展名说明
             * 255216 jpg
             * 208207 doc xls ppt wps
             * 8075 docx pptx xlsx zip
             * 5150 txt
             * 8297 rar
             * 7790 exe
             * 3780 pdf      
             * 
             * 4946/104116 txt
             * 7173        gif 
             * 255216      jpg
             * 13780       png
             * 6677        bmp
             * 239187      txt,aspx,asp,sql
             * 208207      xls.doc.ppt
             * 6063        xml
             * 6033        htm,html
             * 4742        js
             * 8075        xlsx,zip,pptx,mmap,zip
             * 8297        rar   
             * 01          accdb,mdb
             * 7790        exe,dll
             * 5666        psd 
             * 255254      rdp 
             * 10056       bt种子 
             * 64101       bat 
             * 4059        sgf    
             */

        }

    }
}
