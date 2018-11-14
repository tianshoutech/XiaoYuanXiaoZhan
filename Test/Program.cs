using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using XiaoYuanXiaoZhan.Utils.Implement;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var manager = new ExcelManager();
            var wb = manager.CreateWorkbook();
            var sheet = manager.CreateSheet(wb, "测试");
            var style = manager.CreateCommonCellStyle(wb);
            var colWidthDic = new Dictionary<int, int>();
            var mergeRegionList = new List<CellRangeAddress>();
            var rowNum = 0;
            colWidthDic[0] = 10;
            colWidthDic[1] = 20;
            colWidthDic[2] = 30;
            manager.SetColumnWidth(sheet, colWidthDic, 3, 20);
            var row = sheet.CreateRow(rowNum);
            rowNum++;
            manager.CreateCell(row, "Excel文件导出测试", 0);
            mergeRegionList.Add(new CellRangeAddress(0, 0, 0, 2));

            for (int i = 0; i < 3; i++)
            {
                row = sheet.CreateRow(rowNum);
                manager.CreateCell(row, i.ToString(), 0,style);
                manager.CreateCell(row, "姓名" + i, 1,style);
                manager.CreateCell(row, "序号" + i, 2);
                rowNum++;
            }

            manager.CreateMergedRegion(sheet, mergeRegionList);

            //add picture data to this workbook.
            byte[] bytes = System.IO.File.ReadAllBytes(@"D:\MyProject\NPOIDemo\ShapeImage\image1.jpg");
            int pictureIdx = wb.AddPicture(bytes, PictureType.JPEG);

            // Create the drawing patriarch.  This is the top level container for all shapes. 
            var patriarch = sheet.CreateDrawingPatriarch();

            //add a picture
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 0, 0, 0, 1, 3);
            HSSFPicture pict = patriarch.CreatePicture(anchor, pictureIdx);

            FileStream fs = File.OpenWrite(DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            wb.Write(fs);//向打开的这个Excel文件中写入表单并保存。  
            fs.Close();

            Console.WriteLine("生成完毕");
            Console.ReadKey();
        }
    }
}
