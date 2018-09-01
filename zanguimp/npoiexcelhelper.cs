using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Threading;
using NPOI.HSSF.UserModel;

public static class ExcelHelper
{
    public static DataSet ExcelToDataSet(string fileName)
    {
        return ExcelToDataSet(fileName, true);
    }

    public static DataSet ExcelToDataSet(string fileName, bool firstRowAsHeader)
    {
        int sheetCount = 0;
        return ExcelToDataSet(fileName, firstRowAsHeader, out sheetCount);
    }

    public static DataSet ExcelToDataSet(string fileName, bool firstRowAsHeader, out int sheetCount)
    {
        using (DataSet ds = new DataSet())
        {
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = WorkbookFactory.Create(fileStream);
                IFormulaEvaluator evaluator = WorkbookFactory.CreateFormulaEvaluator(workbook);

                sheetCount = workbook.NumberOfSheets;

                for (int i = 0; i < sheetCount; ++i)
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    DataTable dt = ExcelToDataTable(sheet, evaluator, firstRowAsHeader);
                    ds.Tables.Add(dt);
                }
                return ds;
            }
        }
    }

    public static DataTable ExcelToDataTable(string fileName, string sheetName)
    {
        return ExcelToDataTable(fileName, sheetName, true);
    }

    public static DataTable ExcelToDataTable(string fileName, string sheetName, bool firstRowAsHeader)
    {
        using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = WorkbookFactory.Create(fileStream);

            IFormulaEvaluator evaluator = new HSSFFormulaEvaluator(workbook);

            ISheet sheet = workbook.GetSheet(sheetName);

            return ExcelToDataTable(sheet, evaluator, firstRowAsHeader);
        }
    }

    private static DataTable ExcelToDataTable(ISheet sheet, IFormulaEvaluator evaluator, bool firstRowAsHeader)
    {
        if (firstRowAsHeader)
        {
            return ExcelToDataTableFirstRowAsHeader(sheet, evaluator);
        }
        else
        {
            return ExcelToDataTable(sheet, evaluator);
        }
    }

    private static DataTable ExcelToDataTableFirstRowAsHeader(ISheet sheet, IFormulaEvaluator evaluator)
    {
        using (DataTable dt = new DataTable())
        {
            IRow firstRow = sheet.GetRow(0);
            int cellCount = GetCellCount(sheet);

            for (int i = 0; i < cellCount; i++)
            {
                if (firstRow.GetCell(i) != null)
                {
                    dt.Columns.Add(firstRow.GetCell(i).StringCellValue ?? string.Format("F{0}", i + 1), typeof(string));
                }
                else
                {
                    dt.Columns.Add(string.Format("F{0}", i + 1), typeof(string));
                }
            }

            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dr = dt.NewRow();
                FillDataRowByRow(row, evaluator, ref dr);
                dt.Rows.Add(dr);
            }

            dt.TableName = sheet.SheetName;
            return dt;
        }
    }

    private static DataTable ExcelToDataTable(ISheet sheet, IFormulaEvaluator evaluator)
    {
        using (DataTable dt = new DataTable())
        {
            if (sheet.LastRowNum != 0)
            {
                int cellCount = GetCellCount(sheet);

                for (int i = 0; i < cellCount; i++)
                {
                    dt.Columns.Add(string.Format("F{0}", i), typeof(string));
                }

                for (int i = 0; i < sheet.FirstRowNum; ++i)
                {
                    DataRow dr = dt.NewRow();
                    dt.Rows.Add(dr);
                }

                for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    DataRow dr = dt.NewRow();
                    FillDataRowByRow(row, evaluator, ref dr);
                    dt.Rows.Add(dr);
                }
            }

            dt.TableName = sheet.SheetName;
            return dt;
        }
    }

    /// <summary>
    /// 填充数据
    /// </summary>
    /// <param name="row"></param>
    /// <param name="evaluator"></param>
    /// <param name="dr"></param>
    private static void FillDataRowByRow(IRow row, IFormulaEvaluator evaluator, ref DataRow dr)
    {
        if (row != null)
        {
            for (int j = 0; j < dr.Table.Columns.Count; j++)
            {
                ICell cell = row.GetCell(j);

                if (cell != null)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Blank:
                            {
                                dr[j] = DBNull.Value;
                                break;
                            }
                        case CellType.Boolean:
                            {
                                dr[j] = cell.BooleanCellValue;
                                break;
                            }
                        case CellType.Numeric:
                            {
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    dr[j] = cell.DateCellValue;
                                }
                                else
                                {
                                    dr[j] = cell.NumericCellValue;
                                }
                                break;
                            }
                        case CellType.String:
                            {
                                dr[j] = cell.StringCellValue;
                                break;
                            }
                        case CellType.Error:
                            {
                                dr[j] = cell.ErrorCellValue;
                                break;
                            }
                        case CellType.Formula:
                            {
                                cell = evaluator.EvaluateInCell(cell) as HSSFCell;
                                dr[j] = cell.ToString();
                                break;
                            }
                        default:
                            throw new NotSupportedException(string.Format("Unsupported format type:{0}", cell.CellType));
                    }
                }
            }
        }
    }

    /// <summary>
    /// 获取单元格
    /// </summary>
    /// <param name="sheet"></param>
    /// <returns></returns>
    private static int GetCellCount(ISheet sheet)
    {
        int firstRowNum = sheet.FirstRowNum;

        int cellCount = 0;

        for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; ++i)
        {
            IRow row = sheet.GetRow(i);

            if (row != null && row.LastCellNum > cellCount)
            {
                cellCount = row.LastCellNum;
            }
        }
        return cellCount;
    }

    /// <summary>
    /// 数据导出
    /// </summary>
    /// <param name="data"></param>
    /// <param name="sheetName"></param>
    public static void ExportToExcel(this DataTable data)
    {
        ExportToExcel(data, "Sheet1");
    }

    /// <summary>
    /// 数据导出
    /// </summary>
    /// <param name="data"></param>
    /// <param name="sheetName"></param>
    public static void ExportToExcel(this DataTable data, string sheetName)
    {
        SaveFileDialog fileDialog = new SaveFileDialog();
        fileDialog.Filter = "Excel(97-2003)|*.xls|Excel(2007-2013)|*.xlsx";
        if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
        {
            return;
        }


        IWorkbook workbook = new XSSFWorkbook();
        ISheet sheet = workbook.CreateSheet(sheetName);
        IRow rowHead = sheet.CreateRow(0);


        //填写表头
        for (int i = 0; i < data.Columns.Count; i++)
        {
            rowHead.CreateCell(i, CellType.String).SetCellValue(data.Columns[i].ColumnName.ToString());
        }
        //填写内容
        for (int i = 0; i < data.Rows.Count; i++)
        {
            IRow row = sheet.CreateRow(i + 1);
            for (int j = 0; j < data.Columns.Count; j++)
            {
                row.CreateCell(j, CellType.String).SetCellValue(data.Rows[i][j].ToString());
            }
        }

        for (int i = 0; i < data.Columns.Count; i++)
        {
            sheet.AutoSizeColumn(i);
        }

        using (FileStream stream = File.OpenWrite(fileDialog.FileName))
        {
            workbook.Write(stream);
            stream.Close();
        }
        MessageBox.Show("导出数据成功!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        GC.Collect();
    }
}

