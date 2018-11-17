using System;
using System.Collections.Generic;
using System.Text;

namespace XiaoYuanXiaoZhan.Utils.Model
{
    /// <summary>
    /// Excel Manger中添加数据时返回的结果
    /// </summary>
    public class ExcelAddDataReturnModel
    {
        public int StartRow { get; set; }
        public Dictionary<int,int> ColumnWidthList { get; set; }
    }
}
