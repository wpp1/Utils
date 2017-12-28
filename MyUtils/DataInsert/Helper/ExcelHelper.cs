using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataInsert.Helper
{
    public class ExcelHelper
    {
        public static void WriteDataToExceel(string fileName, DataSet ds)
        {
            if (File.Exists(fileName))
                throw new Exception("File is Exists!");
            using (FileStream stream = new FileStream(fileName, FileMode.Create))
            {
                try
                {
                    IWorkbook workbook = ReadToWorkBook(ds, fileName);
                    workbook.Write(stream);
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
        }

        public static void WriteDataToExceel(string fileName, DataTable dt)
        {
            if (File.Exists(fileName))
                throw new Exception("File is Exists!");
            using (FileStream stream = new FileStream(fileName, FileMode.Create))
            {
                try
                {
                    IWorkbook workbook = ReadToWorkBook(dt, fileName);
                    workbook.Write(stream);
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
        }


        public static MemoryStream WriteDataToExcel(DataSet ds, string Extension = ".xls")
        {
            MemoryStream memoryStream = new MemoryStream();
            try
            {
                IWorkbook workbook = ReadToWorkBook(ds, Extension);
                workbook.Write(memoryStream);

                workbook = null;
            }
            catch (Exception exception)
            {
                throw exception;
            }
            return memoryStream;
        }

        private static IWorkbook ReadToWorkBook(DataSet ds, string Extension)
        {
            IWorkbook workbook = null;
            if (Extension.ToLower().EndsWith(".xls"))
                workbook = new HSSFWorkbook();
            else
                workbook = new XSSFWorkbook();
            foreach (DataTable table in ds.Tables)
            {
                ISheet sheet = workbook.CreateSheet(table.TableName);
                IRow headerRow = sheet.CreateRow(0);
                foreach (DataColumn column in table.Columns)
                {
                    headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                }
                int rowIndex = 1;
                foreach (DataRow row in table.Rows)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);
                    DateTime temp = DateTime.Now;
                    foreach (DataColumn column in table.Columns)
                    {
                        if (DateTime.TryParse(row[column].ToString(), out temp))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue(temp.ToString("yyyy-MM-dd"));
                        }
                        else
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                        }
                        sheet.SetColumnWidth(column.Ordinal, 4000);
                    }
                    rowIndex++;
                }
                sheet = null;
                headerRow = null;
            }

            return workbook;
        }

        private static IWorkbook ReadToWorkBook(DataTable dt, string Extension)
        {
            IWorkbook workbook = null;
            if (Extension.ToLower().EndsWith(".xls"))
                workbook = new HSSFWorkbook();
            else
                workbook = new XSSFWorkbook();
            //foreach (DataTable table in ds.Tables)
            //{
            ISheet sheet = workbook.CreateSheet(dt.TableName);
            IRow headerRow = sheet.CreateRow(0);
            foreach (DataColumn column in dt.Columns)
            {
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
            }
            int rowIndex = 1;
            foreach (DataRow row in dt.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dt.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }

                rowIndex++;
            }
            sheet = null;
            headerRow = null;
            //  }

            return workbook;
        }

        public static DataSet ExcelToDataSet(string excelPath)
        {
            return ExcelToDataSet(excelPath, true);
        }
        public static DataSet ExcelToDataSet(string excelPath, bool firstRowAsHeader)
        {
            int sheetCount;
            try
            {
                return ExcelToDataSet(excelPath, firstRowAsHeader, out sheetCount);
            }
            catch
            {
                return ExcelToDataSet(excelPath, firstRowAsHeader, out sheetCount, true);
            }
        }

        public static DataSet ExcelToDataSet(string excelPath, bool firstRowAsHeader, out int sheetCount, bool isXlsx = false)
        {
            using (DataSet ds = new DataSet())
            {
                using (FileStream fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook;
                    IFormulaEvaluator evaluator;
                    GetWorkBook(excelPath, fileStream, out workbook, out evaluator, isXlsx);
                    sheetCount = workbook.NumberOfSheets;
                    for (int i = 0; i < sheetCount; ++i)
                    {
                        DataTable dt = ExcelToDataTable(workbook.GetSheetAt(i), evaluator, firstRowAsHeader);
                        ds.Tables.Add(dt);
                    }
                    return ds;
                }
            }
        }

        public static DataTable ExcelToDataTable(string excelPath, string sheetName, bool firstRowAsHeader = true)
        {
            using (FileStream fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                IFormulaEvaluator evaluator;
                GetWorkBook(excelPath, fileStream, out workbook, out evaluator);
                return ExcelToDataTable(workbook.GetSheet(sheetName), evaluator, firstRowAsHeader);
            }
        }

        public static DataTable ExcelToDataTable(string excelPath, int pageIndex, Stream stream, bool firstRowAsHeader = true)
        {
            using (Stream fileStream = stream)
            {
                IWorkbook workbook;
                IFormulaEvaluator evaluator;
                GetWorkBook(excelPath, fileStream, out workbook, out evaluator);
                return ExcelToDataTable(workbook.GetSheetAt(pageIndex), evaluator, firstRowAsHeader);
            }
        }

        public static DataTable ExcelToDataTable(string excelPath, int pageIndex, bool firstRowAsHeader = true)
        {
            using (FileStream fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook;
                IFormulaEvaluator evaluator;
                GetWorkBook(excelPath, fileStream, out workbook, out evaluator);
                if (pageIndex > workbook.NumberOfSheets)
                    throw new Exception("参数PageIndex>Excel页面总数!");
                return ExcelToDataTable(workbook.GetSheetAt(pageIndex), evaluator, firstRowAsHeader);
            }
        }

        private static void GetWorkBook(string excelPath, FileStream fileStream, out IWorkbook workbook, out IFormulaEvaluator evaluator, bool isUseXlsx = false)
        {
            workbook = null;
            evaluator = null;
            if (excelPath.ToLower().EndsWith(".xls") && !isUseXlsx)
            {
                workbook = new HSSFWorkbook(fileStream, true);
                evaluator = new HSSFFormulaEvaluator(workbook);
            }
            else
            {
                workbook = new XSSFWorkbook(fileStream);
                evaluator = new XSSFFormulaEvaluator(workbook);
            }
        }

        private static void GetWorkBook(string excelPath, Stream fileStream, out IWorkbook workbook, out IFormulaEvaluator evaluator, bool isUseXlsx = false)
        {
            workbook = null;
            evaluator = null;
            if (excelPath.ToLower().EndsWith(".xls") && !isUseXlsx)
            {
                workbook = new HSSFWorkbook(fileStream, true);
                evaluator = new HSSFFormulaEvaluator(workbook);
            }
            else
            {
                workbook = new XSSFWorkbook(fileStream);
                evaluator = new XSSFFormulaEvaluator(workbook);
            }
        }

        #region 内部方法
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
                IRow firstRow = sheet.GetRow(0) as IRow;
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
                    IRow row = sheet.GetRow(i) as IRow;
                    DataRow dr = dt.NewRow();
                    FillDataRowByHSSFRow(row, evaluator, ref dr);
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
                        IRow row = sheet.GetRow(i) as IRow;
                        DataRow dr = dt.NewRow();
                        FillDataRowByHSSFRow(row, evaluator, ref dr);
                        dt.Rows.Add(dr);
                    }
                }
                dt.TableName = sheet.SheetName;
                return dt;
            }
        }
        private static void FillDataRowByHSSFRow(IRow row, IFormulaEvaluator evaluator, ref DataRow dr)
        {
            if (row != null)
            {
                for (int j = 0; j < dr.Table.Columns.Count; j++)
                {
                    ICell cell = row.GetCell(j) as ICell;
                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.Blank:
                                dr[j] = DBNull.Value;
                                break;
                            case CellType.Boolean:
                                dr[j] = cell.BooleanCellValue;
                                break;
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                    dr[j] = cell.DateCellValue;
                                else
                                    dr[j] = cell.NumericCellValue;
                                break;
                            case CellType.String:
                                dr[j] = cell.StringCellValue;
                                break;
                            case CellType.Error:
                                dr[j] = cell.ErrorCellValue;
                                break;
                            case CellType.Formula:
                                cell = evaluator.EvaluateInCell(cell) as ICell;
                                dr[j] = cell.ToString();
                                break;
                            default:
                                throw new NotSupportedException(string.Format("Catched unhandle CellType[{0}]", cell.CellType));
                        }
                    }
                }
            }
        }
        private static int GetCellCount(ISheet sheet)
        {
            int firstRowNum = sheet.FirstRowNum;
            int cellCount = 0;
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; ++i)
            {
                IRow row = sheet.GetRow(i) as IRow;
                if (row != null && row.LastCellNum > cellCount)
                {
                    cellCount = row.LastCellNum;
                }
            }
            return cellCount;
        }

        #endregion

        #region 用于Web导出
        /// <summary>
        /// 用于Web导出
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">文件名</param>
        public static void ExportByWeb(DataTable dtSource, string strHeaderText, string strFileName)
        {
            //HttpContext curContext = HttpContext.Current;
            //try
            //{
            //    // 设置编码和附件格式
            //    curContext.Response.ContentType = "application/vnd.ms-excel";
            //    curContext.Response.ContentEncoding = Encoding.UTF8;
            //    curContext.Response.Charset = "";
            //    curContext.Response.AppendHeader("Content-Disposition",
            //        "attachment;filename=" + HttpUtility.UrlEncode(strFileName, Encoding.UTF8));

            //    curContext.Response.BinaryWrite(Export(dtSource, strHeaderText).GetBuffer());
            //    curContext.Response.End();
            //}
            //catch (Exception e)
            //{

            //    throw;
            //}

        }
        #endregion

        #region DataTable导出到Excel的MemoryStream
        /// <summary>
        /// DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        public static MemoryStream Export(DataTable dtSource, string strHeaderText)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();

            #region 右击文件 属性信息
            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "NPOI";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                /* si.Author = "文件作者信息"; //填加xls文件作者信息
                 si.ApplicationName = "创建程序信息"; //填加xls文件创建程序信息
                 si.LastAuthor = "最后保存者信息"; //填加xls文件最后保存者信息
                 si.Comments = "作者信息"; //填加xls文件作者信息
                 si.Title = "标题信息"; //填加xls文件标题信息
                 si.Subject = "主题信息";//填加文件主题信息*/
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }
            #endregion

            ICellStyle dateStyle = workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            //取得列宽
            int[] arrColWidth = new int[dtSource.Columns.Count];
            foreach (DataColumn item in dtSource.Columns)
            {
                arrColWidth[item.Ordinal] = Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
            }
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;
            foreach (DataRow row in dtSource.Rows)
            {
                #region 新建表，填充表头，填充列头，样式
                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = workbook.CreateSheet();
                    }

                    #region 表头及样式
                    {
                        IRow headerRow = sheet.CreateRow(0);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        ICellStyle headStyle = workbook.CreateCellStyle();
                        // headStyle.Alignment = HorizontalAlignment.CENTER;
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        headerRow.GetCell(0).CellStyle = headStyle;
                        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1));
                        //headerRow.Dispose();
                    }
                    #endregion


                    #region 列头及样式
                    {
                        IRow headerRow = sheet.CreateRow(1);
                        ICellStyle headStyle = workbook.CreateCellStyle();
                        // headStyle.Alignment = HorizontalAlignment.CENTER;
                        IFont font = workbook.CreateFont();
                        font.FontHeightInPoints = 10;
                        font.Boldweight = 700;
                        headStyle.SetFont(font);
                        foreach (DataColumn column in dtSource.Columns)
                        {
                            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                            headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                            //设置列宽
                            sheet.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);
                        }
                        //headerRow.Dispose();
                    }
                    #endregion

                    rowIndex = 2;
                }
                #endregion


                #region 填充内容
                IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dtSource.Columns)
                {
                    ICell newCell = dataRow.CreateCell(column.Ordinal);

                    string drValue = row[column].ToString();

                    switch (column.DataType.ToString())
                    {
                        case "System.String"://字符串类型
                            newCell.SetCellValue(drValue);
                            break;
                        case "System.DateTime"://日期类型
                            DateTime dateV;
                            DateTime.TryParse(drValue, out dateV);
                            newCell.SetCellValue(dateV);

                            newCell.CellStyle = dateStyle;//格式化显示
                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(drValue, out boolV);
                            newCell.SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(drValue, out intV);
                            newCell.SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(drValue, out doubV);
                            newCell.SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理
                            newCell.SetCellValue("");
                            break;
                        default:
                            newCell.SetCellValue("");
                            break;
                    }

                }
                #endregion

                rowIndex++;
            }
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;

                //sheet.Dispose();
                //workbook.Dispose();//一般只用写这一个就OK了，他会遍历并释放所有资源，但当前版本有问题所以只释放sheet
                return ms;
            }
        }
        #endregion

    }
}
