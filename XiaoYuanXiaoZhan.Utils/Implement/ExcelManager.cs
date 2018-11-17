using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using XiaoYuanXiaoZhan.Utils.Const;
using XiaoYuanXiaoZhan.Utils.Interface;
using XiaoYuanXiaoZhan.Utils.Model;

namespace XiaoYuanXiaoZhan.Utils.Implement
{
    public class ExcelManager : IExcelManager
    {
        private ICellStyle _commonCellStyle;
        private IWorkbook _wb;
        private ICellStyle _dateStyle;
        private ICellStyle _doubleStyle;

        public ExcelManager(ExcelTypeEnum excelType = ExcelTypeEnum.XLSX)
        {
            _wb = CreateWorkbook(excelType);
            _commonCellStyle = CreateCommonCellStyle(_wb);
            _dateStyle = GetDateStyle();
            _doubleStyle = GetDouleStyle();
        }

        public IWorkbook Workbook
        {
            get { return _wb; }
        }

        /// <summary>
        /// 创建获取已经存在工作表
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ISheet GetSheet(string sheetName, IWorkbook wb = null)
        {
            if (wb == null)
            {
                wb = _wb;
            }
            ISheet sheet = wb.GetSheet(sheetName);
            if (sheet != null)
            {
                return sheet;
            }

            return wb.CreateSheet(sheetName);
        }

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="excelType"></param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(ExcelTypeEnum excelType = ExcelTypeEnum.XLSX)
        {
            IWorkbook wb = null;
            if (excelType == ExcelTypeEnum.XLS)
            {
                wb = new HSSFWorkbook();
            }
            else
            {
                wb = new XSSFWorkbook();
            }

            return wb;
        }

        /// <summary>
        /// 保存工作簿
        /// </summary>
        /// <param name="wb"></param>
        /// <param name="path"></param>
        public void SaveWorkBook(string path, IWorkbook wb = null)
        {
            if (wb == null)
            {
                wb = _wb;
            }
            FileStream fs = File.OpenWrite(path);
            wb.Write(fs);//向打开的这个Excel文件中写入表单并保存。  
            fs.Close();
            fs = null;
            wb = null;
        }

        /// <summary>
        /// 添加DataTable格式的数据
        /// </summary>
        /// <param name="sheet">需要添加数据的表格</param>
        /// <param name="dt">需要添加的数据</param>
        /// <param name="titleMaps">标题映射</param>
        /// <param name="startRow">开始行</param>
        /// <returns></returns>
        public ExcelAddDataReturnModel AddData(ISheet sheet, DataTable dt, Dictionary<string, string> titleMaps = null, int startRow = 0)
        {
            var result = new ExcelAddDataReturnModel();
            var value = new object();
            ICell cell = null;
            Type type = null;
            var columnWidthList = new Dictionary<int, int>();
            var strValue = string.Empty;

            // 处理标题问题
            var titles = new List<string>();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                titles.Add(dt.Columns[i].ColumnName);
            }

            // 添加标题
            IRow row = SetTitleValue(sheet, titleMaps, startRow, columnWidthList, titles);
            startRow++;

            // 添加内容
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var dtRow = dt.Rows[i];
                row = sheet.CreateRow(startRow);

                for (int j = 0; j < titles.Count; j++)
                {
                    cell = row.CreateCell(j);
                    type = dt.Columns[titles[j]].DataType;
                    value = dtRow[titles[j]];

                    columnWidthList[j] = SetValue(value, cell, type, columnWidthList[j]);
                }

                startRow++;
            }

            result.StartRow = startRow;
            result.ColumnWidthList = columnWidthList;
            return result;
        }

        private static IRow SetTitleValue(ISheet sheet, Dictionary<string, string> titleMaps, int startRow, Dictionary<int, int> columnWidthList, List<string> titles)
        {
            ICell cell = null;
            var strValue = "";
            var row = sheet.CreateRow(startRow);
            for (int i = 0; i < titles.Count; i++)
            {
                cell = row.CreateCell(i);

                if (titleMaps == null || !titleMaps.ContainsKey(titles[i]))
                {
                    strValue = titles[i];
                }
                else
                {
                    strValue = titleMaps[titles[i]];
                }
                cell.SetCellValue(strValue);
                columnWidthList[i] = Encoding.UTF8.GetBytes(strValue).Length;
            }

            return row;
        }

        /// <summary>
        /// 添加列表格式的数据
        /// </summary>
        /// <param name="sheet">需要添加数据的表格</param>
        /// <param name="data">需要添加的数据</param>
        /// <param name="titleMaps">标题映射</param>
        /// <param name="startRow">开始行</param>
        /// <returns></returns>
        public ExcelAddDataReturnModel AddData<T>(ISheet sheet, List<T> data, Dictionary<string, string> titleMaps = null, int startRow = 0)
        {
            var result = new ExcelAddDataReturnModel();
            var value = new object();
            ICell cell = null;
            Type type = null;
            var columnWidthList = new Dictionary<int, int>();
            var strValue = string.Empty;

            // 处理标题问题
            var titles = new List<string>();
            var props = typeof(T).GetProperties().Where(it => it.CanRead && it.PropertyType.IsPublic).ToArray();
            for (int i = 0; i < props.Length; i++)
            {
                titles.Add(props[i].Name);
            }

            // 添加标题
            IRow row = SetTitleValue(sheet, titleMaps, startRow, columnWidthList, titles);
            startRow++;

            // 添加内容
            for (int i = 0; i < data.Count; i++)
            {
                var obj = data[i];
                row = sheet.CreateRow(startRow);

                for (int j = 0; j < titles.Count; j++)
                {
                    cell = row.CreateCell(j);
                    type = props[j].PropertyType;
                    value = props[j].GetValue(obj);
                    columnWidthList[j] = SetValue(value, cell, type, columnWidthList[j]);
                }

                startRow++;
            }

            result.StartRow = startRow;
            result.ColumnWidthList = columnWidthList;
            return result;
        }

        private ICellStyle GetDouleStyle()
        {
            //设置数据显示格式
            IDataFormat dataFormatCustom = _wb.CreateDataFormat();
            var dateStyle = _wb.CreateCellStyle();
            dateStyle.CloneStyleFrom(_commonCellStyle);
            dateStyle.DataFormat = dataFormatCustom.GetFormat("0.00");
            return dateStyle;
        }

        private ICellStyle GetDateStyle()
        {
            //设置数据显示格式
            IDataFormat dataFormatCustom = _wb.CreateDataFormat();
            var dateStyle = _wb.CreateCellStyle();
            dateStyle.CloneStyleFrom(_commonCellStyle);
            dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");
            return dateStyle;
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="cell">单元格</param>
        /// <param name="type">值类型</param>
        /// <param name="colWidth">宽度</param>
        /// <returns></returns>
        private int SetValue(object value, ICell cell, Type type, int colWidth)
        {
            var strValue = string.Empty;
            cell.CellStyle = _commonCellStyle;
            if (type == typeof(int) || type == typeof(long))
            {
                strValue = value.ToString();
                cell.SetCellValue(strValue);
                cell.SetCellType(CellType.Numeric);
            }
            else if (type == typeof(DateTime))
            {
                strValue = DateTime.Parse(value.ToString()).ToString("yyyy-MM-dd HH:mm:ss");
                cell.CellStyle = _dateStyle;
                cell.SetCellValue(strValue);
            }
            else if (type == typeof(double) || type == typeof(float))
            {
                cell.SetCellValue((double)value);
                cell.CellStyle = _doubleStyle;
            }
            else
            {
                strValue = value == null ? "" : value.ToString();
                cell.SetCellValue(strValue);
            }

            var byteCount = Encoding.UTF8.GetBytes(strValue).Length;
            if (byteCount > colWidth)
            {
                colWidth = byteCount;
            }

            return colWidth;
        }

        public ICell CreateCell(IRow row, string value, int cols)
        {
            var cell = row.CreateCell(cols);
            cell.SetCellValue(value);
            return cell;
        }

        public ICell CreateCell(IRow row, string value, int cols, ICellStyle style)
        {
            var cell = row.CreateCell(cols);
            cell.SetCellValue(value);
            cell.CellStyle = style;
            return cell;
        }

        public ICell CreateCell(ISheet sheet, int rowNum, string value, int colNum)
        {
            var row = sheet.CreateRow(rowNum);
            var cell = row.CreateCell(colNum);
            cell.SetCellValue(value);
            return cell;
        }

        /// <summary>
        /// 创建合并单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="regions"></param>
        public void CreateMergedRegion(ISheet sheet, List<CellRangeAddress> regions)
        {
            for (int i = 0; i < regions.Count; i++)
            {
                sheet.AddMergedRegion(regions[i]);
            }
        }

        /// <summary>
        /// 设置列宽度
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="columnWidthDic">列宽数据</param>
        /// <param name="cols">纵列数</param>
        /// <param name="defaultWidth">默认列宽</param>
        public void SetColumnWidth(ISheet sheet, Dictionary<int, int> columnWidthDic, int cols, int? defaultWidth = null, int extWidth = 5)
        {
            for (int i = 0; i < cols; i++)
            {
                if (columnWidthDic.ContainsKey(i))
                {
                    sheet.SetColumnWidth(i, (columnWidthDic[i] + extWidth) * 256);
                }
                else
                {
                    if (defaultWidth != null)
                    {
                        sheet.SetColumnWidth(i, (columnWidthDic[i] + extWidth) * 256);
                    }
                }
            }
        }

        /// <summary>
        /// 创建通用样式
        /// </summary>
        /// <param name="wb"></param>
        /// <returns></returns>
        public ICellStyle CreateCommonCellStyle(IWorkbook wb)
        {
            var style = wb.CreateCellStyle();
            // 设置对齐方式
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.WrapText = false;

            // 设置边框样式
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;

            // 设置字体
            IFont font = wb.CreateFont();
            font.FontName = "宋体";
            font.Color = HSSFColor.Black.Index;
            font.FontHeight = 10;
            //font.Boldweight = (short)FontBoldWeight.Bold;
            style.SetFont(font);

            //style.FillBackgroundColor = HSSFColor.Pink.Index;
            //style.FillPattern = FillPattern.SolidForeground;
            //style.FillForegroundColor = HSSFColor.Black.Index;

            return style;
        }

        /// <summary>
        /// 添加图片
        /// </summary>
        /// <param name="wb">工作簿对象</param>
        /// <param name="sheet">工作表对象</param>
        /// <param name="pics">图片列表</param>
        public void AddPicture(IWorkbook wb, ISheet sheet, List<ExcelPictureModel> pics)
        {
            var pictureIdx = 0;
            var pic = new ExcelPictureModel();
            for (int i = 0; i < pics.Count; i++)
            {
                pic = pics[i];
                pictureIdx = wb.AddPicture(pic.Datas, pic.PicType);
                HSSFPatriarch patriarch = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, pic.Weight, pic.Height, pic.StartCol, pic.StartRow, pic.EndCol, pic.EndRow);
                //##处理照片位置，【图片左上角为（col, row）第row+1行col+1列，右下角为（ col +1, row +1）第 col +1+1行row +1+1列，宽为100，高为50
                HSSFPicture pict = (HSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

                if (pic.IsResize)
                {
                    pict.Resize();
                }
            }
        }
    }
}
