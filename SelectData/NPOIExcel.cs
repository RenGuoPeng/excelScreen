
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SelectData
{
    public class NPOIExcel
    {
        /// <summary>
        /// 将excel导入到datatable
        /// </summary>
        /// <param name="filePath">excel路径</param>
        /// <param name="isColumnName">第一行是否是列名</param>
        /// <returns>返回datatable</returns>
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    // 2007版本
                    if (filePath.IndexOf(".xlsx") > 0)
                        workbook = new XSSFWorkbook(fs);
                    // 2003版本
                    else if (filePath.IndexOf(".xls") > 0)
                        workbook = new HSSFWorkbook(fs);

                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowCount = sheet.LastRowNum;//总行数
                            if (rowCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);//第一行
                                int cellCount = firstRow.LastCellNum;//列数

                                //构建datatable的列
                                if (isColumnName)
                                {
                                    startRow = 1;//如果第一行是列名，则从第二行开始读取
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("column" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }

                                //填充行
                                for (int i = startRow; i <= rowCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null) continue;

                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }

        /// <summary>
        /// 写入excel
        /// </summary>
        /// <param name="dt">datatable</param>
        /// <param name="strFile">strFile</param>
        /// <returns></returns>
        public static bool DataTableToExcel(DataTable dt, string strFile)
        {
            bool result = false;
            IWorkbook workbook = null;
            FileStream fs = null;
            IRow row = null;
            ISheet sheet = null;
            ICell cell = null;
            try
            {
                if (dt != null && dt.Rows.Count > 0)
                {
                    workbook = new XSSFWorkbook();//HSSFWorkbook:是操作Excel2003以前（包括2003）的版本，扩展名是.xls  XSSFWorkbook:是操作Excel2007的版本，扩展名是.xlsx
                    sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表
                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数

                    //设置列头
                    row = sheet.CreateRow(0);//excel第一行设为列头
                    for (int c = 0; c < columnCount; c++)
                    {
                        cell = row.CreateCell(c);
                        cell.SetCellValue(dt.Columns[c].ColumnName);
                    }

                    //设置每行每列的单元格,
                    for (int i = 0; i < rowCount; i++)
                    {
                        row = sheet.CreateRow(i + 1);
                        for (int j = 0; j < columnCount; j++)
                        {
                            cell = row.CreateCell(j);//excel第二行开始写入数据
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    using (fs = File.OpenWrite(strFile))
                    {
                        workbook.Write(fs);//向打开的这个xls文件中写入数据
                        result = true;
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                Console.WriteLine(ex.StackTrace + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToTable(string file)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else if (fileExt == ".xlsm") { workbook = new XSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }

                ISheet sheet = workbook.GetSheetAt(0);
               
                //表头
                IRow header = sheet.GetRow(sheet.FirstRowNum);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据
                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }


        /// <summary>
        /// Excel导入成Datable
        /// </summary>
        /// <param name="file">导入路径(包含文件名与扩展名)</param>
        /// <returns></returns>
        public static DataTable ExcelToTable2(string file, string rangeCell = null,string _range="")
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
                if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(fs); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(fs); } else if (fileExt == ".xlsm") { workbook = new XSSFWorkbook(fs); } else { workbook = null; }
                if (workbook == null) { return null; }
                //获得excel工作表1
                ISheet sheet = workbook.GetSheetAt(0);
                string rangeCellString = "";
                if (rangeCell is null)
                {
                    rangeCellString = _range;
                }
                else
                {
                    
                        int MergeIndex = Convert.ToInt32(rangeCell);
                        Point Merge_a, Merge_b;
                        string MergeName = IsMergeCell(sheet, MergeIndex, out Merge_a, out Merge_b);
                        Merge_b.X = sheet.LastRowNum;
                        Merge_a.X = Merge_a.X + 2;
                        string Address_a = GetAddress(Merge_a); 
                        string Address_b = GetAddress(Merge_b);
                        rangeCellString = Address_a + ":" + Address_b;
                    
                }
               
              
                //获取一个范围值 如：A4:B23
                var range2 = rangeCellString;
                var cellRange = CellRangeAddress.ValueOf(range2);
                //头部标题
                IRow header = sheet.GetRow(cellRange.FirstRow);
                List<int> columns = new List<int>();
                for (int i = cellRange.FirstColumn; i <= cellRange.LastColumn; i++)
                {
                    object obj = GetValueType(header.GetCell(i));
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                    }
                    else
                        dt.Columns.Add(new DataColumn(obj.ToString()));
                    columns.Add(i);
                }
                //数据
                for (int i = cellRange.FirstRow+1; i <= cellRange.LastRow; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    int k = 0;
                    foreach (int j in columns)
                    {
                        
                        dr[k] = GetValueType(sheet.GetRow(i).GetCell(j));
                        if (dr[k] != null && dr[k].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                        k++;
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
            }
            return dt;
        }

        /// <summary>
        /// Datable导出成Excel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file">导出路径(包括文件名与扩展名)</param>
        public static void TableToExcel(DataTable dt, string file)
        {
            IWorkbook workbook;
            string fileExt = Path.GetExtension(file).ToLower();
            if (fileExt == ".xlsx") { workbook = new XSSFWorkbook(); } else if (fileExt == ".xls") { workbook = new HSSFWorkbook(); } else { workbook = null; }
            if (workbook == null) { return; }
            ISheet sheet = string.IsNullOrEmpty(dt.TableName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(dt.TableName);

            //表头
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>
        /// 获取当前单元格所在的合并单元格的位置
        /// </summary>
        /// <param name="sheet">sheet表单</param>
        /// <param name="rowIndex">行索引 0开始</param>
        /// <param name="colIndex">列索引 0开始</param>
        /// <param name="start">合并单元格左上角坐标</param>
        /// <param name="end">合并单元格右下角坐标</param>
        /// <returns>返回false表示非合并单元格</returns>
        public static bool IsMergeCell(ISheet sheet, int rowIndex, int colIndex, out Point start, out Point end)
        {
            bool result = false;
            start = new Point(0, 0);
            end = new Point(0, 0);
            if ((rowIndex < 0) || (colIndex < 0)) return result;
            int regionsCount = sheet.NumMergedRegions;
            for (int i = 0; i < regionsCount; i++)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                //sheet.IsMergedRegion(range); 
                if (rowIndex >= range.FirstRow && rowIndex <= range.LastRow && colIndex >= range.FirstColumn && colIndex <= range.LastColumn)
                {
                    start = new Point(range.FirstRow, range.FirstColumn);
                    end = new Point(range.LastRow, range.LastColumn);
                    result = true;
                    break;
                }
            }
            return result;
        }
        /// <summary>
        /// 获取指定【合并单元格】的坐标
        /// </summary>
        /// <param name="sheet">sheet表单</param>
        /// <param name="rowIndex">行索引 0开始</param>
        /// <param name="colIndex">列索引 0开始</param>
        /// <param name="start">合并单元格左上角坐标</param>
        /// <param name="end">合并单元格右下角坐标</param>
        /// <returns>返回false表示非合并单元格</returns>
        public static string IsMergeCell(ISheet sheet, int Mergedindex , out Point start, out Point end)
        {
            CellRangeAddress range = sheet.GetMergedRegion(Mergedindex);
            start = new Point(range.FirstRow, range.FirstColumn);
            end = new Point(range.LastRow, range.LastColumn);
            string cellRange = sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn).ToString();
            
            //if (cellRange.ToString()== "结构 C")
            //{
            //    cellRange = cellRange;
            //}
            //else
            //{
            //    CCWin.MessageBoxEx.Show("选择的索引不是结构C");
            //}
            return cellRange.ToString();
            //start = new Point(0, 0);
            //end = new Point(0, 0);
            //if ((rowIndex < 0) || (colIndex < 0)) return result;
            //int regionsCount = sheet.NumMergedRegions;
            //for (int i = 0; i < regionsCount; i++)
            //{
            //    CellRangeAddress range = sheet.GetMergedRegion(i);
            //    //sheet.IsMergedRegion(range); 
            //    if (rowIndex >= range.FirstRow && rowIndex <= range.LastRow && colIndex >= range.FirstColumn && colIndex <= range.LastColumn)
            //    {
            //        start = new Point(range.FirstRow, range.FirstColumn);
            //        end = new Point(range.LastRow, range.LastColumn);
            //        result = true;
            //        break;
            //    }
            //}

        }
        /// <summary>
        /// 根据坐标获取地址
        /// </summary>
        /// <param name="cell1"></param>
        /// <param name="Address"></param>
       
        /// <returns></returns>
        public static string GetAddress(Point cell1)
        {
            string Celladderss = "";
            string[] ColumnName = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            if (cell1.Y<=26)
            {
                 Celladderss = ColumnName[cell1.Y] + (cell1.X);
            }
            else
            {
                int SecondColumnCount = cell1.Y + 1;
                string[] ColumnNames = new string[SecondColumnCount];
                for (int i = 0; i < ColumnName.Length; i++)
                {
                    ColumnNames[i] = ColumnName[i];
                }
                for (int i = 0; i < ColumnName.Length; i++)
                {
                    for (int j = 0; j < ColumnName.Length; j++)
                    {
                        if (26 * (i+1)+j < SecondColumnCount)
                        {
                            ColumnNames[26 * (i + 1) +j] = ColumnName[i] + ColumnName[j];
                        }
                        else
                        {
                            break;
                        }
                    }
                    if (ColumnNames[SecondColumnCount - 1] != null)
                    {
                        break;
                    }

                }
                 Celladderss = ColumnNames[cell1.Y] + (cell1.X);
            }



            return Celladderss;

        }
        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:
                    return null;
                case CellType.Boolean: //BOOLEAN:
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:
                    return Math.Round(cell.NumericCellValue,3);
                case CellType.String: //STRING:
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:
                default:
                    return "=" + cell.CellFormula;

            }
        }

    }
}

