using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text;
using Tests.Model;
using XiaoYuanXiaoZhan.Utils.Implement;
using XiaoYuanXiaoZhan.Utils.Model;

namespace Tests
{
    public class ExcelTest
    {
        public static void GenerateExcel()
        {
            var manager = new ExcelManager();
            var sheet = manager.GetSheet("测试");
            //var personList = new List<Person>();
            //for (int i = 0; i < 10000; i++)
            //{
            //    personList.Add(new Person()
            //    {
            //        Id = i,
            //        Name = "test" + i,
            //        Age = i,
            //        Address = "dkjgjkgjkdjkakdafiokadkghkahkg" + i
            //    });
            //}
            //Stopwatch sw = new Stopwatch();
            //sw.Start();
            //for (int i = 0; i < 100; i++)
            //{
            //    manager.AddData(sheet, personList, null, i * 10000);
            //}
            //manager.SaveWorkBook(wb, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            //sw.Stop();
            //Console.WriteLine(sw.ElapsedMilliseconds);

            var dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name");
            dt.Columns.Add("Age", typeof(double));
            dt.Columns.Add("Birthday", typeof(DateTime));
            dt.Columns.Add("Address");
            var titleMap = new Dictionary<string, string>();
            titleMap.Add("Id", "编号");
            titleMap.Add("Name", "姓名");
            titleMap.Add("Age", "年龄");
            titleMap.Add("Birthday", "生日");
            titleMap.Add("Address", "地址");
            Random rnd = new Random();
            for (int i = 0; i < 10; i++)
            {
                var row = dt.NewRow();
                row["Id"] = i;
                row["Name"] = "test" + i;
                row["Age"] = rnd.Next(1000000) * 1.0 / 1000;
                row["Address"] = "dkjgjkgjkdjkakdafiokadkghkahkg" + i;
                row["Birthday"] = DateTime.Now;
                dt.Rows.Add(row);
            }
            var sw = new Stopwatch();
            var colList = new ExcelAddDataReturnModel();
            sw.Start();
            for (int i = 0; i < 1; i++)
            {
                colList = manager.AddData(sheet, dt, titleMap, i * 10000);
            }
            manager.SetColumnWidth(sheet, colList.ColumnWidthList, 5, null);
            manager.SaveWorkBook(DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            sw.Stop();
            Console.WriteLine(sw.ElapsedMilliseconds);


            //var style = manager.CreateCommonCellStyle(wb);
            //var colWidthDic = new Dictionary<int, int>();
            //var mergeRegionList = new List<CellRangeAddress>();
            //var rowNum = 0;
            //colWidthDic[0] = 10;
            //colWidthDic[1] = 20;
            //colWidthDic[2] = 30;
            //manager.SetColumnWidth(sheet, colWidthDic, 3, 20);
            //var row = sheet.CreateRow(rowNum);
            //rowNum++;
            //manager.CreateCell(row, "Excel文件导出测试", 0);
            //mergeRegionList.Add(new CellRangeAddress(0, 0, 0, 2));

            //for (int i = 0; i < 3; i++)
            //{
            //    row = sheet.CreateRow(rowNum);
            //    manager.CreateCell(row, i.ToString(), 0, style);
            //    manager.CreateCell(row, "姓名" + i, 1, style);
            //    manager.CreateCell(row, "序号" + i, 2);
            //    rowNum++;
            //}

            //manager.CreateMergedRegion(sheet, mergeRegionList);

            ////add picture data to this workbook.
            //byte[] bytes = System.IO.File.ReadAllBytes(@"D:\MyProject\NPOIDemo\ShapeImage\image1.jpg");
            //int pictureIdx = wb.AddPicture(bytes, PictureType.JPEG);

            //// Create the drawing patriarch.  This is the top level container for all shapes. 
            //var patriarch = sheet.CreateDrawingPatriarch();

            ////add a picture
            //HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 0, 0, 0, 1, 3);
            //HSSFPicture pict = patriarch.CreatePicture(anchor, pictureIdx);

            //FileStream fs = File.OpenWrite(DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");
            //wb.Write(fs);//向打开的这个Excel文件中写入表单并保存。  
            //fs.Close();
        }
    }
}
