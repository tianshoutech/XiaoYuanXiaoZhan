using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace XiaoYuanXiaoZhan.Utils.Model
{
    /// <summary>
    /// Excel导出图片实体
    /// </summary>
    public class ExcelPictureModel
    {
        /// <summary>
        /// 图片类型
        /// </summary>
        public PictureType PicType { get; set; }
        /// <summary>
        /// 图片放置起始行
        /// </summary>
        public int StartRow { get; set; }
        /// <summary>
        /// 图片放置结束行
        /// </summary>
        public int EndRow { get; set; }
        /// <summary>
        /// 图片放置起始列
        /// </summary>
        public int StartCol { get; set; }
        /// <summary>
        /// 图片放置结束列
        /// </summary>
        public int EndCol { get; set; }
        /// <summary>
        /// 图片高度
        /// </summary>
        public int Height { get; set; }
        /// <summary>
        /// 图片宽度
        /// </summary>
        public int Weight { get; set; }
        /// <summary>
        /// 图片的二进制数据
        /// </summary>
        public byte[] Datas { get; set; }
        /// <summary>
        /// 图片的存放路径
        /// </summary>
        public string Path { get; set; }
        /// <summary>
        /// 图片添加到Excel后是否以原始大小显示
        /// </summary>
        public bool IsResize { get; set; }
    }
}
