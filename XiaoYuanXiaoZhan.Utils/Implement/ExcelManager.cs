using System;
using System.Collections.Generic;
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

namespace XiaoYuanXiaoZhan.Utils.Implement
{
    public class ExcelManager : IExcelManager
    {
        public ISheet CreateSheet(IWorkbook wb, string sheetName)
        {
            ISheet sheet = wb.GetSheet(sheetName);
            if (sheet != null)
            {
                return sheet;
            }
            return wb.CreateSheet(sheetName);
        }

        public IWorkbook CreateWorkbook( ExcelTypeEnum excelType = ExcelTypeEnum.XLSX)
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

        public void CreateMergedRegion(ISheet sheet, List<CellRangeAddress> regions)
        {
            for (int i = 0; i < regions.Count; i++)
            {
                sheet.AddMergedRegion(regions[i]);
            }
        }

        public void SetColumnWidth(ISheet sheet, Dictionary<int, int> columnWidthDic, int cols, int? defaultWidth)
        {
            for (int i = 0; i < cols; i++)
            {
                if (columnWidthDic.ContainsKey(i))
                {
                    sheet.SetColumnWidth(i, columnWidthDic[i] * 256);
                }
                else
                {
                    if (defaultWidth != null)
                    {
                        sheet.SetColumnWidth(i, defaultWidth.Value * 256);
                    }
                }
            }
        }

        public ICellStyle CreateCommonCellStyle(IWorkbook wb)
        {
            var style = wb.CreateCellStyle();
            // 设置对齐方式
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.WrapText = true;
            
            // 设置边框样式
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;

            // 设置字体
            IFont font = wb.CreateFont();
            font.FontName = "宋体";
            font.Color = HSSFColor.Black.Index;
            font.FontHeight = 20;
            font.Boldweight = (short)FontBoldWeight.Bold;
            style.SetFont(font);

            //style.FillBackgroundColor = HSSFColor.Pink.Index;
            //style.FillPattern = FillPattern.SolidForeground;
            //style.FillForegroundColor = HSSFColor.Black.Index;

            //设置数据显示格式
            //IDataFormat dataFormatCustom = wb.CreateDataFormat();
            //dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");

            return style;
        }
    }
}
