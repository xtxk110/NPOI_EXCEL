using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using System.Data;
using System.IO;


public class NpoiHelper
{
    private static NpoiHelper _obj = new NpoiHelper();
    private string Extend = string.Empty;//EXCEL的扩展名
    private static bool IsCompatible;//是否生成兼容模板EXCEL文件
    private bool IsStream;//是否以流初始化工作簿
    private Stream stream;
    private NpoiHelper() { }
    /// <summary>
    /// 
    /// </summary>
    /// <param name="isCompatible">是否生成兼容模式EXCEL(XLS)</param>
    private IWorkbook CreateWorkbook(bool isCompatible)
    {
        IWorkbook _instance = null;
        if (isCompatible)
            _instance = new HSSFWorkbook();
        else
            _instance = new XSSFWorkbook();
        return _instance;
    }
    /// <summary>
    /// 
    /// </summary>
    /// <param name="isCompatible">是否生成兼容模式EXCEL(XLS)</param>
    /// <param name="stream">根据文件流生成工作簿</param>
    private IWorkbook CreateWorkbook(bool isCompatible, Stream stream)
    {
        IWorkbook _instance = null;
        if (isCompatible)
            _instance = new HSSFWorkbook(stream);
        else
            _instance = new XSSFWorkbook(stream);
        return _instance;
    }
    /// <summary>
    /// 创建表格头单元格样式
    /// </summary>
    /// <param name="workbook">工作簿</param>
    /// <returns>返回固定单元格样式</returns>
    private ICellStyle GetHeaderCellStyle(IWorkbook workbook)
    {
        ICellStyle style = workbook.CreateCellStyle();
        style.FillPattern = FillPattern.SolidForeground;
        style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
        style.Alignment = HorizontalAlignment.Center;
        style.VerticalAlignment = VerticalAlignment.Center;
        return style;
    }
    
    /// <summary>
    /// 获取NPOI工具类实例
    /// </summary>
    /// <param name="isCompatible">是否生成兼容模式EXCEL(XLS),默认为兼容模式</param>
    /// <returns></returns>
    public static NpoiHelper GetInstance(bool isCompatible=true)
    {
        IsCompatible = isCompatible;    
        if (_obj == null)
            _obj = new NpoiHelper();
        if (IsCompatible)
            _obj.Extend = @".xls";
        else
            _obj.Extend = @".xlsx";
        return _obj;
    }
    #region 导出Excel
    private void _ExportToSheet(DataTable sourceTable,IWorkbook workbook, ISheet sheet, string sheetName, string filePath)
    {
        ICellStyle headCellStyle = GetHeaderCellStyle(workbook);
        IRow headerRow = sheet.CreateRow(0);
        // handling header.
        for (int i = 0; i < sourceTable.Columns.Count; i++)
        {
            DataColumn column = sourceTable.Columns[i];
            ICell headerCell = headerRow.CreateCell(column.Ordinal);
            headerCell.SetCellValue(column.ColumnName);
            headerCell.CellStyle = headCellStyle;
            sheet.SetColumnWidth(i, column.ColumnName.Length * 256 + 1);
        }
        // handling value.
        int rowIndex = 1;

        foreach (DataRow row in sourceTable.Rows)
        {
            IRow dataRow = sheet.CreateRow(rowIndex);
            int colIndex = 0;
            foreach (DataColumn column in sourceTable.Columns)
            {
                ICell cell = dataRow.CreateCell(column.Ordinal);
                string value = (row[column] ?? "").ToString();
                cell.SetCellValue(value);
                int curWidth = sheet.GetColumnWidth(colIndex);
                int newWidth = value.Length * 256 + 1;
                if (curWidth < newWidth)
                    sheet.SetColumnWidth(colIndex, newWidth);
                colIndex++;
            }

            rowIndex++;
        }
       
    }
    /// <summary>
    /// DataTable导出Excel(只针对单一表头)
    /// </summary>
    /// <param name="sourceTable">源数据表</param>
    /// <param name="sheetName">工作表名称(只包含单信工作表时,有效)</param>
    /// <param name="filePath">EXCEL保存的路径（包含文件名,不包含扩展名）(默认保存到当前应用程序的根目录)</param>
    /// <returns></returns>
    public string ExportToExcel(DataTable sourceTable, string sheetName = "result", string filePath = null)
    {
        if (sourceTable.Rows.Count <= 0) return null;

        IWorkbook workbook = CreateWorkbook(IsCompatible);
        if (string.IsNullOrEmpty(filePath))
        {
            filePath = AppDomain.CurrentDomain.BaseDirectory + "excel" + Extend;//当前应用程序的根目录
        }
        else
            filePath = filePath + Extend;
        ISheet sheet = workbook.CreateSheet(sheetName);
        _ExportToSheet(sourceTable,workbook, sheet, sheetName, filePath);
        FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        workbook.Write(fs);
        fs.Dispose();
        sheet = null;
        workbook = null;

        return filePath;
    }
    /// <summary>
    /// DataSet导出EXCEL(只针对单一表头)
    /// </summary>
    /// <param name="sourceSet">源数据DATASET</param>
    /// <param name="filePath">EXCEL保存的路径（包含文件名,不包含扩展名）(默认保存到当前应用程序的根目录)</param>
    /// <returns></returns>
    public string ExportToExcel(DataSet sourceSet, string filePath = null)
    {
        if (sourceSet.Tables.Count <= 0) return null;
        IWorkbook workbook = CreateWorkbook(IsCompatible);
        if (string.IsNullOrEmpty(filePath))
        {
            filePath = AppDomain.CurrentDomain.BaseDirectory + "excel" + Extend;//当前应用程序的根目录
        }
        else
            filePath = filePath + Extend;
        int indexNo = 1;
        foreach (DataTable dt in sourceSet.Tables)
        {
            if (dt.Rows.Count <= 0)
                continue;
            string sheetName = "Sheet" + indexNo;
            ISheet sheet = workbook.CreateSheet(sheetName);
            _ExportToSheet(dt,workbook, sheet, sheetName, filePath);
        }
        FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        workbook.Write(fs);
        fs.Dispose();
        workbook = null;
        return filePath;
    }
    /// <summary>
    /// LIST集合导出EXCEL(只针对单一表头)
    /// </summary>
    /// <typeparam name="T">代表对象不能多层嵌套,只能由简单数据类型组成</typeparam>
    /// <param name="data">T对象集合</param>
    /// <param name="headerNameList"></param>
    /// <param name="sheetName">工作表名称(只包含单信工作表时,有效)</param>
    /// <param name="filePath">EXCEL保存的路径（包含文件名,不包含扩展名）(默认保存到当前应用程序的根目录)</param>
    /// <returns></returns>
    public string ExportToExcel<T>(List<T> data, string sheetName = "result", string filePath = null) where T : class
    {
        if (data.Count <= 0) return null;
        IWorkbook workbook = CreateWorkbook(IsCompatible);
        if (string.IsNullOrEmpty(filePath))
        {
            filePath = AppDomain.CurrentDomain.BaseDirectory + "excel" + Extend;//当前应用程序的根目录
        }
        else
            filePath = filePath + Extend;
        Type t = typeof(T);
        ICellStyle cellStyle = GetHeaderCellStyle(workbook);
        ISheet sheet = workbook.CreateSheet(sheetName);
        IRow headerRow = sheet.CreateRow(0);
        IList<string> headerNameList = new List<string>();
        foreach (var item in t.GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public))
        {
            headerNameList.Add(item.Name);
        }
        for (int i = 0; i < headerNameList.Count; i++)
        {
            ICell cell = headerRow.CreateCell(i);
            cell.SetCellValue(headerNameList[i]);
            cell.CellStyle = cellStyle;
            sheet.SetColumnWidth(i, (headerNameList[i].Length+1) * 256 );
        }        
        int rowIndex = 1;
        foreach (T item in data)
        {
            IRow dataRow = sheet.CreateRow(rowIndex);
            for (int n = 0; n < headerNameList.Count; n++)
            {
                int curWidth = sheet.GetColumnWidth(n);
                string pValue = (t.GetProperty(headerNameList[n]).GetValue(item, null)??"").ToString();
                ICell cell = dataRow.CreateCell(n);
                cell.SetCellValue(pValue);
                int newWidth = (pValue.Length+1) * 256;
                if (curWidth < newWidth)
                    sheet.SetColumnWidth(n, newWidth);
            }
            rowIndex++;
        }
        FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        workbook.Write(fs);
        fs.Dispose();

        sheet = null;
        headerRow = null;
        workbook = null;

        return filePath;
    }
    #endregion
    #region 导入Excel
    /// <summary>
    /// 从工作表中生成DataTable(只针对单一表头)
    /// </summary>
    /// <param name="sheet">工作表</param>
    /// <param name="headerRowIndex">表头索引(单行表头)</param>
    /// <returns></returns>
    private DataTable GetDataTableFromSheet(ISheet sheet, int headerRowIndex)
    {
        DataTable table = new DataTable();

        try
        {
            IRow headerRow = sheet.GetRow(headerRowIndex);//获取表头行(只能时单行表头)
            int cellCount = headerRow.LastCellNum;//获取列数
                                                  //生成表头
            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                ICell singleCell = headerRow.GetCell(i);
                string header = singleCell == null ? string.Empty : singleCell.StringCellValue;
                DataColumn column = new DataColumn(header);
                table.Columns.Add(column);
            }

            //生成数据
            for (int i = (headerRowIndex + 1); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    DataRow dataRow = table.NewRow();
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            dataRow[j] = row.GetCell(j).ToString();
                        }
                    }
                    table.Rows.Add(dataRow);
                }
            }
        }
        catch { }

        return table;
    }
    /// <summary>
    /// 由Excel导入DataTable()(只针对单一表头)
    /// </summary>
    /// <param name="excelFileStream">Excel文件流</param>
    /// <param name="sheetName">Excel工作表名称(可以工作表索引或名称)</param>
    /// <param name="headerRowIndex">Excel表头行索引</param>
    /// <returns>DataTable</returns>
    public DataTable ImportFromExcel(Stream excelFileStream, string sheetName, int headerRowIndex)
    {
        IWorkbook workbook = CreateWorkbook(IsCompatible, excelFileStream);
        ISheet sheet = null;
        int sheetIndex = -1;
        if (int.TryParse(sheetName, out sheetIndex))
        {
            sheet = workbook.GetSheetAt(sheetIndex);
        }
        else
        {
            sheet = workbook.GetSheet(sheetName);
        }
        if (sheet == null)
            return null;
        DataTable table = GetDataTableFromSheet(sheet, headerRowIndex);

        excelFileStream.Close();
        workbook = null;
        sheet = null;
        return table;
    }

    /// <summary>
    /// 由Excel导入DataTable(只针对单一表头)
    /// </summary>
    /// <param name="excelFilePath">Excel文件路径，为物理路径。</param>
    /// <param name="sheetName">Excel工作表名称(可以工作表索引或名称)</param>
    /// <param name="headerRowIndex">Excel表头行索引</param>
    /// <returns>DataTable</returns>
    public DataTable ImportFromExcel(string excelFilePath, string sheetName, int headerRowIndex)
    {
        using (FileStream stream = System.IO.File.OpenRead(excelFilePath))
        {
            return ImportFromExcel(stream, sheetName, headerRowIndex);
        }
    }

    /// <summary>
    /// 由Excel导入DataSet，如果有多个工作表，则导入多个DataTable(只针对单一表头)
    /// </summary>
    /// <param name="excelFileStream">Excel文件流</param>
    /// <param name="headerRowIndex">Excel表头行索引</param>
    /// <param name="isCompatible">是否为兼容模式</param>
    /// <returns>DataSet</returns>
    public DataSet ImportFromExcel(Stream excelFileStream, int headerRowIndex)
    {
        DataSet ds = new DataSet();
        IWorkbook workbook = CreateWorkbook(IsCompatible, excelFileStream);
        for (int i = 0; i < workbook.NumberOfSheets; i++)
        {
            ISheet sheet = workbook.GetSheetAt(i);
            DataTable table = GetDataTableFromSheet(sheet, headerRowIndex);
            ds.Tables.Add(table);
        }

        excelFileStream.Close();
        workbook = null;

        return ds;
    }

    /// <summary>
    /// 由Excel导入DataSet，如果有多个工作表，则导入多个DataTable(只针对单一表头)
    /// </summary>
    /// <param name="excelFilePath">Excel文件路径，为物理路径。</param>
    /// <param name="headerRowIndex">Excel表头行索引</param>
    /// <returns>DataSet</returns>
    public DataSet ImportFromExcel(string excelFilePath, int headerRowIndex)
    {
        using (FileStream stream = System.IO.File.OpenRead(excelFilePath))
        {
            return ImportFromExcel(stream, headerRowIndex);
        }
    }

    #endregion
}
