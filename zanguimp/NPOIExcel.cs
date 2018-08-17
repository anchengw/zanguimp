using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace NPOIExcelOpa
{
    public class NPOIExcel
    {
        /// <summary>
        /// 读取2007Excel[.xlsx]或读取2003Excel[.xls](返回DataTable)
        /// </summary>
        /// <param name="path">Excel路径</param>
        /// <param name="isheader">第一行是否列名，否表默认列名为column1，column2等</param>
        /// <returns>表</returns>
        private static DataTable ReadExcel(string path,bool isheader)
        {
            try
            {
                DataTable dt = new DataTable();
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //IWorkbook workbook = new XSSFWorkbook(fs);//2007
                    //IWorkbook workbook = new HSSFWorkbook(fs);//2003
                    IWorkbook workbook = WorkbookFactory.Create(fs);//工厂模式
                    ISheet sheet = workbook.GetSheetAt(0); //取第一个工作表
                    int rfirst = sheet.FirstRowNum;//工作表第一行
                    int rlast = sheet.LastRowNum; //工作表最后一行
                    IRow row = sheet.GetRow(rfirst);
                    int cfirst = row.FirstCellNum;//工作表第一列
                    int clast = row.LastCellNum;//工作表最后一列
                    //构建表列
                    if (isheader)
                    {
                        for (int i = cfirst; i < clast; i++)
                        {
                            if (row.GetCell(i) != null)
                                dt.Columns.Add(row.GetCell(i).StringCellValue, System.Type.GetType("System.String"));
                        }
                        rfirst = rfirst + 1;
                        row = null;
                    }
                    else
                    {
                        for (int i = cfirst; i < clast; i++)
                        {
                            //DataColumn column = new DataColumn("column" + (i + 1));
                            dt.Columns.Add("column" + (i + 1), System.Type.GetType("System.String"));
                        }
                    }
                    for (int i = rfirst; i <= rlast; i++)
                    {
                        DataRow r = dt.NewRow();
                        IRow ir = sheet.GetRow(i);
                        for (int j = cfirst; j < clast; j++)
                        {
                            if (ir.GetCell(j) != null)
                            {
                                r[j] = ir.GetCell(j).ToString();
                            }
                        }
                        dt.Rows.Add(r);
                        ir = null;
                        r = null;
                    }
                    sheet = null;
                    workbook = null;
                }
                return dt;
            }
            catch (Exception ex)
            {
                throw (ex);//System.Windows.Forms.MessageBox.Show("Excel格式错误或者Excel正由另一进程在访问");
            }
        }
        /// <summary>
        /// 判断文件格式
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static bool Is2007Excel(string filePath)
        {
            string fileclass = "";
            //string extension = "";

            //extension = System.IO.Path.GetExtension(filePath);
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
            finally
            {
                reader.Close();
                stream.Close();
                stream.Dispose();
            }
            if (fileclass == "8075")// && extension.ToLower().Equals(".xlsx"))
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
        /// <summary>
        /// 根据路径删除文件
        /// </summary>
        /// <param name="path"></param>
        public static void DeleteFile(string path)
        {
            FileAttributes attr = File.GetAttributes(path);
            if (attr == FileAttributes.Directory)
            {
                Directory.Delete(path, true);
            }
            else
            {
                File.Delete(path);
            }
        }
        /// <summary>
        /// 自动建立目录
        /// </summary>
        /// <param name="path"></param>
        public static void autoCreateDir(string path)
        {
            string filePath = Path.GetDirectoryName(path);
            if (!System.IO.Directory.Exists(filePath))
            {
                System.IO.Directory.CreateDirectory(path);//不存在就创建目录 
            }
        }
        /// <summary>
        /// 另存为2003格式的EXCEL
        /// </summary>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        public static string excelSaveas(string excelFile)
        {
            string excelFullName = Environment.CurrentDirectory + "\\temp\\" + Path.GetFileNameWithoutExtension(excelFile) + "03.xls";//excelFile.Substring(0, excelFile.LastIndexOf(".")) + "03.xls";

            autoCreateDir(excelFullName);
            if (File.Exists(excelFullName))
                DeleteFile(excelFullName);
            object missing = System.Reflection.Missing.Value;
            Excel.Application myExcel = new Excel.Application();//lauch excel application
            if (myExcel == null)
            {
                return excelFile;//打开EXCEL应用失败 
            }
            else
            {
                myExcel.Visible = false;
                myExcel.UserControl = true;
                //以只读的形式打开EXCEL文件
                Excel.Workbook myBook = myExcel.Application.Workbooks.Open(excelFile, missing, true, missing, missing, missing,
                     missing, missing, missing, true, missing, missing, missing, missing, missing);
                myBook.SaveAs(excelFullName, Excel.XlFileFormat.xlExcel8, null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
                myExcel.Quit();  //退出Excel文件
                myExcel = null;
                System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName("excel");
                foreach (System.Diagnostics.Process pro in procs)
                {
                    pro.Kill();//杀掉进程
                }
                System.Threading.Thread.Sleep(100);//等待进程退出，否则会出现文件正在被另一进程访问的错误
                GC.Collect();
            }
            return excelFullName;
        }
        public static DataTable readExcel(string excelFullName ,bool isHeader)
        {
            if (!File.Exists(excelFullName))
                return null;
            DataTable dt = null;
            try
            {
                dt = ReadExcel(excelFullName,isHeader);
            }
            catch(Exception ex)
            {
                dt = ReadExcel(excelSaveas(excelFullName), isHeader);
            }
            return dt;
        }
    }
}
