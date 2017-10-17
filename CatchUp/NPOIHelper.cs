using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace System
{
    using NPOI.HSSF.UserModel;
    using System.Data;
    using System.IO;
    using System.Text;
    using NPOI.HPSF;
    using NPOI.HSSF.Util;
    using NPOI.XSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using ICSharpCode.SharpZipLib.Zip;
    public class NPOIHelper
    {
        public static ICell GetMergedRegionFirstCell(ISheet sheet, int row, int column)
        {
            int sheetMergeCount = sheet.NumMergedRegions;

            for (int i = 0; i < sheetMergeCount; i++)
            {
                var ca = sheet.GetMergedRegion(i);
                int firstColumn = ca.FirstColumn;
                int lastColumn = ca.LastColumn;
                int firstRow = ca.FirstRow;
                int lastRow = ca.LastRow;
                if (row >= firstRow && row <= lastRow)
                {
                    if (column >= firstColumn && column <= lastColumn)
                    {
                        IRow fRow = sheet.GetRow(firstRow);
                        ICell fCell = fRow.GetCell(firstColumn);
                        return fCell;
                    }
                }
            }

            return null;
        }
        public static void SetCellStyle(XSSFWorkbook xssfworkbook, ICell cell)
        {
            ICellStyle style = xssfworkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BottomBorderColor = HSSFColor.Black.Index;
            style.LeftBorderColor = HSSFColor.Black.Index;
            style.RightBorderColor = HSSFColor.Black.Index;
            style.TopBorderColor = HSSFColor.Black.Index;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
  
            //设置表头字体
            IFont font1 = xssfworkbook.CreateFont();
            font1.Boldweight = 700;
            font1.FontName = "微软雅黑";
            style.SetFont(font1);//将设置好的字体样式设置给单元格样式对象。
            cell.CellStyle = style;
        }
        public static void SetMergedCellStyle(XSSFWorkbook xssfworkbook, ISheet sheet, CellRangeAddress region,int fontSize)
        {
            ICellStyle style = xssfworkbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BottomBorderColor = HSSFColor.Black.Index;
            style.LeftBorderColor = HSSFColor.Black.Index;
            style.RightBorderColor = HSSFColor.Black.Index;
            style.TopBorderColor = HSSFColor.Black.Index;

            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

            //设置表头字体
            IFont font1 = xssfworkbook.CreateFont();
            font1.Boldweight = 700;
            font1.FontHeight = fontSize;
            font1.FontName = "微软雅黑";
            style.SetFont(font1);//将设置好的字体样式设置给单元格样式对象。
            
            
            //cell.CellStyle = style;
            for (int i = region.FirstRow; i <= region.LastRow; i++)
            {
                IRow row = sheet.GetRow(i);
                for (int j = region.FirstColumn; j <= region.LastColumn; j++)
                {
                    ICell singleCell = HSSFCellUtil.GetCell(row, (short)j);
                    singleCell.CellStyle = style;
                }
            } 
        }
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="sheet">要合并单元格所在的sheet</param>
        /// <param name="rowstart">开始行的索引</param>
        /// <param name="rowend">结束行的索引</param>
        /// <param name="colstart">开始列的索引</param>
        /// <param name="colend">结束列的索引</param>
        public static void SetCellRangeAddress(ISheet sheet, int rowstart, int rowend, int colstart, int colend)
        {
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowstart, rowend, colstart, colend);
            sheet.AddMergedRegion(cellRangeAddress);
        }

        #region DataTable导出到Excel文件 + static void Export(DataTable dtSource, string strHeaderText, string strFileName)
        /// <summary>
        /// DataTable导出到Excel文件
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// <param name="strFileName">保存位置</param>
        public static void Export(DataTable dtSource, string strHeaderText, string strFileName)
        {
            string[] strFileNameArr = strFileName.Split('.');
            string fileName = strFileNameArr[strFileNameArr.Length - 1].ToLower();
            if (fileName != "xlsx")
            {
                using (MemoryStream ms = Export(dtSource, strHeaderText))
                {
                    using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                    {
                        byte[] data = ms.ToArray();
                        fs.Write(data, 0, data.Length);
                        fs.Flush();
                    }
                }
            }
            else
            {
                TableToExcelForXLSX(dtSource, strFileName);
            }
        }
        #endregion

        #region 追加内容 + static void AppendForXLSX(DataTable dt, string strFileName, int begin)
        /// <summary>
        /// 追加内容
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="strFileName">EXCEL文件物理路径</param>
        /// <param name="begin">追加内容开始索引</param>
        public static void AppendForXLSX(DataTable dt, string strFileName,int begin)
        {
            using (FileStream fs = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
                ISheet sheet = xssfworkbook.GetSheetAt(0);
                
                //数据  
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row1 = sheet.CreateRow(i + begin);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }

                //转为字节数组  
                using (MemoryStream stream = new MemoryStream())
                {
                    xssfworkbook.Write(stream);
                    var buf = stream.ToArray();

                    //追加内容  
                    using (FileStream fs2 = new FileStream(strFileName, FileMode.Open, FileAccess.Write))
                    {
                        fs2.Write(buf, 0, buf.Length);
                        fs2.Flush();
                    }
                }
            }
        }
        #endregion

        #region 追加内容 + static void AppendForXLSX(DataTable dt, string strFileName)
        /// <summary>
        /// 追加内容
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="strFileName">EXCEL文件物理路径</param>
        public static void AppendForXLSX(DataTable dt, string strFileName)
        {
            int totalRowCount;
            AppendForXLSX(dt, strFileName, out totalRowCount);
        } 
        #endregion

        #region 追加内容 + static void AppendForXLSX(DataTable dt, string strFileName, out int totalRowCount)
        /// <summary>
        ///  追加内容
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="strFileName">EXCEL文件物理路径</param>
        /// <param name="totalRowCount">返回数据总行数</param>
        public static void AppendForXLSX(DataTable dt, string strFileName, out int totalRowCount)
        {
            totalRowCount = 0;
            using (FileStream fs = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
                ISheet sheet = xssfworkbook.GetSheetAt(0);
                int begin = sheet.LastRowNum + 1;
                totalRowCount = sheet.LastRowNum + dt.Rows.Count;
                //数据  
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row1 = sheet.CreateRow(i + begin);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }

                //转为字节数组  
                using (MemoryStream stream = new MemoryStream())
                {
                    xssfworkbook.Write(stream);
                    var buf = stream.ToArray();

                    //追加内容  
                    using (FileStream fs2 = new FileStream(strFileName, FileMode.Open, FileAccess.Write))
                    {
                        fs2.Write(buf, 0, buf.Length);
                        fs2.Flush();
                    }
                }
            }
        } 
        #endregion

        #region DataTable导出到Excel的MemoryStream - static MemoryStream Export(DataTable dtSource, string strHeaderText)
        /// <summary>
        /// DataTable导出到Excel的MemoryStream
        /// </summary>
        /// <param name="dtSource">源DataTable</param>
        /// <param name="strHeaderText">表头文本</param>
        /// 
        private static MemoryStream Export(DataTable dtSource, string strHeaderText)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet();

            #region 右击文件 属性信息
            {
                //DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                //dsi.Company = "NPOI";
                //workbook.DocumentSummaryInformation = dsi;

                //SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                //si.Author = "文件作者信息"; //填加xls文件作者信息
                //si.ApplicationName = "创建程序信息"; //填加xls文件创建程序信息
                //si.LastAuthor = "最后保存者信息"; //填加xls文件最后保存者信息
                //si.Comments = "作者信息"; //填加xls文件作者信息
                //si.Title = "标题信息"; //填加xls文件标题信息
                //si.Subject = "主题信息";//填加文件主题信息
                //si.CreateDateTime = DateTime.Now;
                //workbook.SummaryInformation = si;
            }
            #endregion

            //NPOI.SS.UserModel.ICellStyle dateStyle = workbook.CreateCellStyle();
            //NPOI.SS.UserModel.IDataFormat format = workbook.CreateDataFormat();
            //dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

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
                    if (!string.IsNullOrEmpty(strHeaderText))
                    {
                        #region 表头及样式
                        {
                            NPOI.SS.UserModel.IRow headerRow = sheet.CreateRow(0);
                            headerRow.HeightInPoints = 25;
                            headerRow.CreateCell(0).SetCellValue(strHeaderText);

                            NPOI.SS.UserModel.ICellStyle headStyle = workbook.CreateCellStyle();
                            headStyle.Alignment = HorizontalAlignment.Center;
                            NPOI.SS.UserModel.IFont font = workbook.CreateFont();
                            font.FontHeightInPoints = 20;
                            font.Boldweight = 700;
                            headStyle.SetFont(font);
                            headerRow.GetCell(0).CellStyle = headStyle;
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, dtSource.Columns.Count - 1));
                        }
                    }

                        #endregion
                    #region 列头及样式
                    {
                        int index = 0;
                        if (!string.IsNullOrEmpty(strHeaderText))
                        {
                            index = 1;
                        }
                        NPOI.SS.UserModel.IRow headerRow = sheet.CreateRow(index);
                        NPOI.SS.UserModel.ICellStyle headStyle = workbook.CreateCellStyle();
                        headStyle.Alignment = HorizontalAlignment.Center;
                        NPOI.SS.UserModel.IFont font = workbook.CreateFont();
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
                    if (!string.IsNullOrEmpty(strHeaderText))
                    {
                        rowIndex = 2;
                    }
                    rowIndex = 1;
                }
                #endregion
                #region 填充内容
                NPOI.SS.UserModel.IRow dataRow = sheet.CreateRow(rowIndex);
                foreach (DataColumn column in dtSource.Columns)
                {
                    NPOI.SS.UserModel.ICell newCell = dataRow.CreateCell(column.Ordinal);

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

                            //newCell.CellStyle = dateStyle;//格式化显示
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

        #region 文件下载 + static void DownLoadFile(string absoluteFilePath)
        /// <summary>
        /// 文件下载
        /// </summary>
        /// <param name="absoluteFilePath"></param>
        public static void DownLoadFile(string absoluteFilePath)
        {
            string fileExName = Path.GetExtension(absoluteFilePath);
            string name = Path.GetFileName(absoluteFilePath);
            FileInfo DownloadFile = new FileInfo(absoluteFilePath);
            System.Web.HttpContext.Current.Response.Clear();
            System.Web.HttpContext.Current.Response.ClearHeaders();
            System.Web.HttpContext.Current.Response.Buffer = false;
            System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream";//对于Office2007要用这种ContentType
            System.Web.HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=" + System.Web.HttpUtility.UrlEncode(Path.GetFileName(absoluteFilePath), System.Text.Encoding.UTF8));
            System.Web.HttpContext.Current.Response.AppendHeader("Content-Length", DownloadFile.Length.ToString());
            System.Web.HttpContext.Current.Response.WriteFile(DownloadFile.FullName);
            System.Web.HttpContext.Current.Response.Flush();
            if (File.Exists(absoluteFilePath))
            {
                File.Delete(absoluteFilePath);
            }
            System.Web.HttpContext.Current.Response.End();
        }
        #endregion

        #region 多文件打包下载 + static void DownloadMultiFiles(IEnumerable<string> files, string zipFileName)
        /// <summary>
        /// 多文件打包下载
        /// </summary>
        /// <param name="files"></param>
        /// <param name="zipFileName"></param>
        public static void DownloadMultiFiles(List<string> phyPaths, string zipFileName)
        {
            //根据所选文件打包下载
            MemoryStream ms = new MemoryStream();
            byte[] buffer = null;
            using (ZipFile file = ZipFile.Create(ms))
            {
                file.BeginUpdate();
                file.NameTransform = new MyNameTransfom();//通过这个名称格式化器，可以将里面的文件名进行一些处理。默认情况下，会自动根据文件的路径在zip中创建有关的文件夹。

                foreach (string phyPath in phyPaths)
                {
                    file.Add(phyPath);
                }


                file.CommitUpdate();

                buffer = new byte[ms.Length];
                ms.Position = 0;
                ms.Read(buffer, 0, buffer.Length);
            }


            System.Web.HttpContext.Current.Response.AddHeader("content-disposition", "attachment;filename=" + System.Web.HttpUtility.UrlEncode(zipFileName, System.Text.Encoding.UTF8));
            System.Web.HttpContext.Current.Response.BinaryWrite(buffer);
            System.Web.HttpContext.Current.Response.Flush();
            for (int i = 0; i < phyPaths.Count; i++)
            {
                string pPath = phyPaths[i];
                if (File.Exists(pPath))
                {
                    File.Delete(pPath);
                }
            }

            System.Web.HttpContext.Current.Response.End();
        }
        #endregion

        #region 读取excel + static DataTable Import(string strFileName)
        /// <summary>
        /// 读取excel
        /// 默认第一行为标头
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public static DataTable Import(string strFileName, int headIndex)
        {
            string[] strFileNameArr = strFileName.Split('.');
            string fileName = strFileNameArr[strFileNameArr.Length - 1].ToLower();
            string extName = Path.GetExtension(strFileName);
            if (extName != ".xlsx")
            {
                DataTable dt = new DataTable();

                HSSFWorkbook hssfworkbook;
                using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                {
                    hssfworkbook = new HSSFWorkbook(file);
                }
                NPOI.SS.UserModel.ISheet sheet = hssfworkbook.GetSheetAt(0);
                System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                //sheet.FirstRowNum：获取第一行（表头通常是第一行）
                NPOI.SS.UserModel.IRow headerRow = sheet.GetRow(headIndex);
                int cellCount = headerRow.LastCellNum;
                //表头
                for (int j = 0; j < cellCount; j++)
                {
                    NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
                    if (cell == null || cell.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + j.ToString()));
                        //continue;  
                    }
                    else
                    {
                        if (dt.Columns.Contains(cell.ToString()))//说明重复了
                        {
                            dt.Columns.Add(new DataColumn(Guid.NewGuid() + cell.ToString()));
                        }
                        else
                        {
                            dt.Columns.Add(new DataColumn(cell.ToString()));
                        }
                    }
                }
                //数据
                for (int i = (headIndex + 1); i <= sheet.LastRowNum; i++)
                {
                    NPOI.SS.UserModel.IRow row = sheet.GetRow(i);
                    DataRow dataRow = dt.NewRow();

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            dataRow[j] = row.GetCell(j).ToString();
                    }

                    dt.Rows.Add(dataRow);
                }
                return dt;
            }
            else
            {
                return ExcelToTableForXLSX(strFileName, headIndex);
            }
        }
        #endregion

        /// <summary>
        /// 导入
        /// </summary>
        /// <param name="strFileName"></param>
        /// <param name="headIndex">表头开始的行索引</param>
        /// <returns></returns>
        public static DataSet Import2(string strFileName, int[] headIndexArr, int sheetCount)
        {
            string extName = Path.GetExtension(strFileName);
            if (extName != ".xlsx")
            {
                DataSet ds = new DataSet();

                for (int ii = 0; ii < sheetCount; ii++)
                {
                    DataTable dt = new DataTable();
                    int headIndex = headIndexArr[ii];
                    HSSFWorkbook hssfworkbook;
                    using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
                    {
                        hssfworkbook = new HSSFWorkbook(file);
                    }
                    NPOI.SS.UserModel.ISheet sheet = hssfworkbook.GetSheetAt(ii);
                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                    //sheet.FirstRowNum：获取第一行（表头通常是第一行）
                    NPOI.SS.UserModel.IRow headerRow = sheet.GetRow(headIndex);
                    int cellCount = headerRow.LastCellNum;
                    //表头
                    for (int j = 0; j < cellCount; j++)
                    {
                        NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
                        if (cell == null || cell.ToString() == string.Empty)
                        {
                            dt.Columns.Add(new DataColumn("Columns" + j.ToString()));
                            //continue;  
                        }
                        else
                        {
                            if (dt.Columns.Contains(cell.ToString()))//说明重复了
                            {
                                dt.Columns.Add(new DataColumn(Guid.NewGuid() + cell.ToString()));
                            }
                            else
                            {
                                dt.Columns.Add(new DataColumn(cell.ToString()));
                            }
                        }
                    }
                    //数据
                    for (int i = (headIndex + 1); i <= sheet.LastRowNum; i++)
                    {
                        NPOI.SS.UserModel.IRow row = sheet.GetRow(i);
                        DataRow dataRow = dt.NewRow();

                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                                dataRow[j] = row.GetCell(j).ToString();

                        }

                        dt.Rows.Add(dataRow);
                    }
                    ds.Tables.Add(dt);
                }

                return ds;
            }
            else
            {
                return ExcelToTableForXLSX2(strFileName, headIndexArr, sheetCount);
            }
        }


        #region Excel2007
        /// <summary>  
        /// 将Excel文件中的数据读出到DataTable中(xlsx)  
        /// </summary>  
        /// <param name="file"></param>  
        /// <returns></returns>  
        private static DataTable ExcelToTableForXLSX(string file, int headIndex)
        {
            DataTable dt = new DataTable();
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
                ISheet sheet = xssfworkbook.GetSheetAt(0);

                //表头  
                IRow header = sheet.GetRow(headIndex);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXLSX(header.GetCell(i) as XSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        //continue;  
                    }
                    else
                    {
                        if (dt.Columns.Contains(obj.ToString()))//说明重复了
                        {
                            dt.Columns.Add(new DataColumn(Guid.NewGuid() + obj.ToString()));
                        }
                        else
                        {
                            dt.Columns.Add(new DataColumn(obj.ToString()));
                        }
                    }
                    columns.Add(i);
                }
                //数据  
                for (int i = headIndex + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXLSX(sheet.GetRow(i).GetCell(j) as XSSFCell);
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

        

        private static DataSet ExcelToTableForXLSX2(string file, int[] headIndexArr, int sheetCount)
        {

            //FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read);

            DataSet ds = new DataSet();
            for (int ii = 0; ii < sheetCount; ii++)
            {
                DataTable dt = new DataTable();

                //XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);
                int headIndex = headIndexArr[ii];
                XSSFWorkbook xssfworkbook;
                using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    xssfworkbook = new XSSFWorkbook(fs);//FileStream每次读完一次会关闭
                }
                ISheet sheet = xssfworkbook.GetSheetAt(ii);

                //表头  
                IRow header = sheet.GetRow(headIndex);
                List<int> columns = new List<int>();
                for (int i = 0; i < header.LastCellNum; i++)
                {
                    object obj = GetValueTypeForXLSX(header.GetCell(i) as XSSFCell);
                    if (obj == null || obj.ToString() == string.Empty)
                    {
                        dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                        //continue;  
                    }
                    else
                    {
                        if (dt.Columns.Contains(obj.ToString()))//说明重复了
                        {
                            dt.Columns.Add(new DataColumn(Guid.NewGuid() + obj.ToString()));
                        }
                        else
                        {
                            dt.Columns.Add(new DataColumn(obj.ToString()));
                        }
                    }
                    columns.Add(i);
                }
                //数据  
                for (int i = headIndex + 1; i <= sheet.LastRowNum; i++)
                {
                    DataRow dr = dt.NewRow();
                    bool hasValue = false;
                    foreach (int j in columns)
                    {
                        dr[j] = GetValueTypeForXLSX(sheet.GetRow(i).GetCell(j) as XSSFCell);
                        if (dr[j] != null && dr[j].ToString() != string.Empty)
                        {
                            hasValue = true;
                        }
                        else//有可能是合并的格子，所以为空，去上一行单元格的值
                        {
                            if (sheet.GetRow(i).GetCell(j) != null && sheet.GetRow(i).GetCell(j).IsMergedCell)
                            {
                                object firstMergedCell = GetMergedRegionFirstCell(sheet, i, j);
                                if (firstMergedCell != null)
                                {
                                    dr[j] = firstMergedCell;
                                    hasValue = true;
                                }
                            }
                        }
                    }
                    if (hasValue)
                    {
                        dt.Rows.Add(dr);
                    }
                }
                ds.Tables.Add(dt);
            }

            return ds;
        }

        /// <summary>
        /// 将多个DataTable分别分成同一个工作簿的sheet表(只生成一个EXCEL文件)
        /// </summary>
        /// <param name="dts"></param>
        /// <param name="file"></param>
        public static void ExportForXLSX(DataTable[] dts, string file)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            
            for (int i = 0; i < dts.Length; i++)
            {
                ISheet sheet = xssfworkbook.CreateSheet();
                DataTable dt = dts[i];
                //表头  
                IRow row = sheet.CreateRow(0);
                for (int ii = 0; ii < dt.Columns.Count; ii++)
                {
                    ICell cell = row.CreateCell(ii);
                    cell.SetCellValue(dt.Columns[ii].ColumnName);
                }

                //数据  
                for (int ii = 0; ii < dt.Rows.Count; ii++)
                {
                    IRow row1 = sheet.CreateRow(ii + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[ii][j].ToString());
                    }
                }
            }
            

            //转为字节数组  
            using (MemoryStream stream = new MemoryStream())
            {
                xssfworkbook.Write(stream);
                var buf = stream.ToArray();

                //保存为Excel文件  
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }
            }
        }


        /// <summary>  
        /// 将DataTable数据导出到Excel文件中(xlsx)  
        /// </summary>  
        /// <param name="dt"></param>  
        /// <param name="file"></param>  
        private static byte[] TableToExcelForXLSX(DataTable dt, string file)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet();

            //表头  
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
                sheet.SetColumnWidth(i, 4500);
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
            using (MemoryStream stream = new MemoryStream())
            {
                xssfworkbook.Write(stream);
                var buf = stream.ToArray();

                //保存为Excel文件  
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                    return buf;
                }
            }
        }

        /// <summary>  
        /// 获取单元格类型(xlsx)  
        /// </summary>  
        /// <param name="cell"></param>  
        /// <returns></returns>  
        private static object GetValueTypeForXLSX(XSSFCell cell)
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
                    short format = cell.CellStyle.DataFormat;
                    //181:yyyy/MM/dd HH:mm:ss
                    short[] arr = { 14, 17, 20, 21, 22, 31, 32, 57, 58, 177, 178, 179,181};//由于CellType没有DateTime类型，DateTime类型都是Numeric类型，所以要将它跟数字类型区分开来
                    if (arr.Contains(format))
                    {
                        DateTime date = cell.DateCellValue;

                        string re = date.ToString("yyy-MM-dd HH:mm:ss");
                        return re;
                    }
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }
        #endregion
    }
}
