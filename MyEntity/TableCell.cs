using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace MyEntity
{
    public class TableCell
    {
        public int RowIndex { set; get; }
        public int ColumIndex { set; get; }
        public string Value { set; get; }
        public MergeRange MergeRange { set; get; }
        public bool IsMergeRange { set; get; }
    }
    /// <summary>
    /// 合并单元格的信息
    /// </summary>
    public class MergeRange
    {
        /// <summary>
        /// 顶部行索引
        /// </summary>
        public int TopRow;
        /// <summary>
        /// 底部行索引
        /// </summary>
        public int BottomRow;
        /// <summary>
        /// 左边列索引
        /// </summary>
        public int LeftColumn;
        /// <summary>
        /// 右边列索引
        /// </summary>
        public int RightColumn;
    }
}
