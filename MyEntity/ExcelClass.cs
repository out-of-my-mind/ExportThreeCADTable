using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace MyEntity
{
    public static class ExcelClass
    {
        /// <summary>
        /// 行数
        /// </summary>
        public static int colNum { get; set; }

        /// <summary>
        /// 列数
        /// </summary>
        public static int rowNum { get; set; }

        /// <summary>
        /// 行高
        /// </summary>
        public static double rowHeight { get; set; }

        /// <summary>
        /// 行宽
        /// </summary>
        public static double colWidth { get; set; }

        /// <summary>
        /// 表头
        /// </summary>
        public static List<string> headString { get; set; }

        /// <summary>
        /// 选择行，
        /// </summary>
        public static int selectRow { get; set; }

        /// <summary>
        /// 选择列
        /// </summary>
        public static int selectCol { get; set; }
    }
}
