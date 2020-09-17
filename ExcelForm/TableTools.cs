using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MyEntity;
using System.Reflection;//后期绑定
using System.Text.RegularExpressions;

namespace Util
{
    /// <summary>C:\Users\tom\Desktop\ExportThreeCADTable\ExcelForm\TableTools.cs
    /// 工具类
    /// </summary>
    public static class MyTools
    {
        public static bool noChangeExcelVar = false; //不修改ExcelClass变量
        /// <summary>
        /// COM接口返回值比较理想，后期绑定。
        /// </summary>
        /// <param name="en"></param>
        /// <returns></returns>
        public static Extents3d getComEntityBoxHQBD(Entity en)
        {
            //out参数来自晓东CAD家园的实例改编。
            Type enti = Type.GetTypeFromHandle(Type.GetTypeHandle(en.AcadObject));
            //使用此方法，可传入out参数，本次传入pt类型，其他类型不知道是否可以。
            object[] argspl1 = new object[2];
            argspl1[0] = new VariantWrapper(0);
            argspl1[1] = new VariantWrapper(0);
            ParameterModifier pmpl1 = new ParameterModifier(2);
            pmpl1[0] = true;
            pmpl1[1] = true;
            ParameterModifier[] modifierspl1 = new ParameterModifier[] { pmpl1 };
            enti.InvokeMember("GetBoundingBox", BindingFlags.InvokeMethod, null, en.AcadObject, argspl1, modifierspl1, null, null);
            Point3d pt1 = new Point3d((double[])argspl1[0]);
            Point3d pt2 = new Point3d((double[])argspl1[1]);
            return new Extents3d(pt1, pt2);
        }
        /// <summary>
        /// 获取对象的矩形中心点坐标，MTEXT特殊一点。
        /// </summary>
        /// <param name="en"></param>
        /// <returns></returns>
        public static Point3d getAreaCenterPoint(this Entity en)
        {
            Point3d resultPt = Point3d.Origin;
            Extents3d box = en.getRecBox();
            return box.MinPoint.MidPoint(box.MaxPoint);
        }
        /// <summary>
        /// 获取多行文字的包围框。
        /// </summary>
        /// <param name="mt"></param>
        /// <returns></returns>
        public static Extents3d GetBoundingBox(this MText mt)
        {
            Point3d pt = mt.Location;

            double zw = mt.ActualWidth;  //实际宽度
            double zh = mt.ActualHeight; //实际高度
            AttachmentPoint dqfs = mt.Attachment;
            double h1, h2, w1, w2;
            if (dqfs == AttachmentPoint.TopLeft || dqfs == AttachmentPoint.TopCenter || dqfs == AttachmentPoint.TopRight) //123
            {
                //竖向在上边
                h1 = -zh;
                h2 = 0.0;
            }
            else if (dqfs == AttachmentPoint.MiddleLeft || dqfs == AttachmentPoint.MiddleCenter || dqfs == AttachmentPoint.MiddleRight)//456
            {
                //竖向在中间
                h1 = (zh / -2.0);
                h2 = zh / 2.0;
            }
            else
            {
                //竖向在下边
                h1 = 0.0;
                h2 = zh;
            }
            if (dqfs == AttachmentPoint.TopCenter || dqfs == AttachmentPoint.MiddleCenter || dqfs == AttachmentPoint.BottomCenter)
            {
                //横向在中间
                w1 = zw / 2.0;//往右
                w2 = zw / -2.0;//往左
            }
            else if (dqfs == AttachmentPoint.TopRight || dqfs == AttachmentPoint.MiddleRight || dqfs == AttachmentPoint.BottomRight)
            {
                //横向在右边
                w1 = 0.0;
                w2 = -zw;
            }
            else
            {
                //横向在左边
                w1 = zw;
                w2 = 0.0;
            }
            Point3d dd1 = pt + new Vector3d(w1, h1, 0);//原点的对角点。
            Point3d dd2 = pt + new Vector3d(w2, h2, 0);

            double minX = Math.Min(dd1.X, dd2.X);
            double maxX = Math.Max(dd1.X, dd2.X);
            double minY = Math.Min(dd1.Y, dd2.Y);
            double maxY = Math.Max(dd1.Y, dd2.Y);

            Point3d p1 = new Point3d(minX, minY, 0);
            Point3d p2 = new Point3d(maxX, maxY, 0);

            Extents3d box = new Extents3d(p1, p2);
            box.TransformBy(Matrix3d.Rotation(mt.Rotation, mt.Normal, pt));//考虑到旋转问题，转换一下矩阵。
            return box;
        }
        /// <summary>
        /// 获取矩形包围框
        /// </summary>
        /// <param name="en"></param>
        /// <returns></returns>
        public static Extents3d getRecBox(this Entity en)
        {
            Extents3d box;
            string str = "";
            if (en is MText)
            {
                //多行文字获取BOX方法
                MText mt = en as MText;
                box = mt.GetBoundingBox();
            }
            else// if (en is DBText)
            {
                //除多行文字外通过COM编程方法获取BOX;
                box = MyTools.getComEntityBoxHQBD(en);
            }
            /*else
            {
                //不是多行文字；单行文字。
                box = en.GeometricExtents;
                if (en is DBText)
                    str += ((DBText)en).Position.ToString();
            }*/
            //System.Windows.Forms.MessageBox.Show(box.ToString()+str);
            return box;
        }
        /// <summary>
        /// 获取Extents3d的高度
        /// </summary>
        /// <param name="ext"></param>
        /// <returns></returns>
        public static double getHeight(this Extents3d ext)
        {
            return ext.MaxPoint.Y - ext.MinPoint.Y;
        }
        /// <summary>
        /// 获取Extents3d的宽度
        /// </summary>
        /// <param name="ext"></param>
        /// <returns></returns>
        public static double getWidth(this Extents3d ext)
        {
            return ext.MaxPoint.X - ext.MinPoint.X;
        }
        /// <summary>
        /// 获取两个点之间的中点
        /// </summary>
        /// <param name="pt1">第一点</param>
        /// <param name="pt2">第二点</param>
        /// <returns>返回两个点之间的中点</returns>
        public static Point3d MidPoint(this Point3d pt1, Point3d pt2)
        {
            Point3d midPoint = new Point3d((pt1.X + pt2.X) / 2.0,
                                        (pt1.Y + pt2.Y) / 2.0,
                                        (pt1.Z + pt2.Z) / 2.0);
            return midPoint;
        }
        /// <summary>
        /// 获取多行文字的真实内容
        /// HJ分析应该是正则表达式。
        /// </summary>
        /// <param name="mtext">多行文字对象</param>
        /// <returns>返回多行文字的真实内容</returns>
        public static string GetText(this MText mtext, string separator = "\\")
        {
            string content = mtext.Contents;//多行文本内容
            //将多行文本按“\\”进行分割
            string[] strs = content.Split(new string[] { @"\\" }, StringSplitOptions.None);
            //指定不区分大小写
            RegexOptions ignoreCase = RegexOptions.IgnoreCase;
            for (int i = 0; i < strs.Length; i++)
            {
                //删除段落缩进格式
                strs[i] = Regex.Replace(strs[i], @"\\pi(.[^;]*);", "", ignoreCase);
                //删除制表符格式
                strs[i] = Regex.Replace(strs[i], @"\\pt(.[^;]*);", "", ignoreCase);
                //删除堆迭格式
                strs[i] = Regex.Replace(strs[i], @"\\S(.[^;]*)(\^|#|\\)(.[^;]*);", @"$1$3", ignoreCase);
                strs[i] = Regex.Replace(strs[i], @"\\S(.[^;]*)(\^|#|\\);", "$1", ignoreCase);
                //删除字体、颜色、字高、字距、倾斜、字宽、对齐格式
                strs[i] = Regex.Replace(strs[i], @"(\\F|\\C|\\H|\\T|\\Q|\\W|\\A)(.[^;]*);", "", ignoreCase);
                //删除下划线、删除线格式
                strs[i] = Regex.Replace(strs[i], @"(\\L|\\O|\\l|\\o)", "", ignoreCase);
                //删除不间断空格格式
                strs[i] = Regex.Replace(strs[i], @"\\~", "", ignoreCase);
                //删除换行符格式
                strs[i] = Regex.Replace(strs[i], @"\\P", "\n", ignoreCase);
                //删除换行符格式(针对Shift+Enter格式)
                //strs[i] = Regex.Replace(strs[i], "\n", "", ignoreCase);
                //删除{}
                strs[i] = Regex.Replace(strs[i], @"({|})", "", ignoreCase);
                //替换回\\,\{,\}字符
                //strs[i] = Regex.Replace(strs[i], @"\x01", @"\", ignoreCase);
                //strs[i] = Regex.Replace(strs[i], @"\x02", @"{", ignoreCase);
                //strs[i] = Regex.Replace(strs[i], @"\x03", @"}", ignoreCase);
            }
            return string.Join(separator, strs);//将文本中的特殊字符去掉后重新连接成一个字符串
        }
        public static void setDataGridView(this DataGridView dataGridView1,int rowNum,int colNum,double rowHeight,double colWidth,List<string> headString)
        {
            MessageBox.Show("设定表格!");
        }
        /// <summary>
        /// 限制输入整数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public static bool inputint(object sender, KeyPressEventArgs e)
        {
            bool b = false;
            int keyCode = (int)e.KeyChar;
            if ((keyCode < 48 || keyCode > 57) && keyCode != 8 && keyCode != 13)
            {
                b = true;
            }
            return b;
        }
        /// <summary>
        /// 限制输入浮点数
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <returns></returns>
        public static bool inputdouble(object sender, KeyPressEventArgs e)
        {
            bool b = false;
            int keyCode = (int)e.KeyChar;
            if ((keyCode < 48 || keyCode > 57) && keyCode != 8)
            {
                if (keyCode == 46)
                {
                    TextBox textbox = (TextBox)sender;
                    if (textbox.Text.IndexOf(".") != -1)
                        b = true;
                }
                else
                    b = true;
            }
            return b;
        }

        public static int Max(int a, int b)
        {
            if (a > b)
                return a;
            else
                return b;
        }
        //选择对象并返回DBObjectCollection
        public static DBObjectCollection HJ_GetObjsBySelectionFilter(this Database db, SelectionFilter filter, OpenMode mode, bool openErased)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //选择符合条件的所有实体
            PromptSelectionResult entSelected = ed.GetSelection(filter);
            if (entSelected.Status != PromptStatus.OK) return null;
            SelectionSet ss = entSelected.Value;

            DBObjectCollection ents = new DBObjectCollection();

            using (Transaction ts = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in ss.GetObjectIds())
                {
                    DBObject obj = ts.GetObject(id, mode, openErased);
                    if (obj != null)
                        ents.Add(obj);
                }
                ts.Commit();//没有这个就不会提交内部的修改
            }
            return ents;
        }
        /// <summary>
        /// 重新定义Table行数列数
        /// </summary>
        /// <param name="table"></param>
        /// <param name="biaoTouRowsCount"></param>
        /// <param name="rowNum"></param>
        /// <param name="colNum"></param>
        /// <param name="rowHeight"></param>
        /// <param name="colWidth"></param>
        public static void ReSize(this Table table, int biaoTouRowsCount, int rowNum, int colNum, double rowHeight, double colWidth)
        {
            //列，插入、删除
            if (colNum > table.NumColumns)
            {
#if cad2007
#else
                if (table.CanInsertColumn(table.NumColumns))
#endif
                {
                    int ksCol = table.NumColumns;
                    table.InsertColumns(table.NumColumns, colWidth, colNum - table.NumColumns);
                    int jsCol = table.NumColumns;
                    if (table.NumRows >= 2)
                    {
                        for (int i = ksCol; i < jsCol; i++)
                        {
                            table.SetTextString(1, i, "列w_" + i);
                        }
                    }
                }
            }
            if (colNum < table.NumColumns)
            {
#if cad2007
#else
                if (table.CanDeleteColumns(table.NumColumns - (table.NumColumns - colNum), table.NumColumns - colNum))
#endif
                {
                    table.DeleteColumns(table.NumColumns - (table.NumColumns - colNum), table.NumColumns - colNum);
                }
            }
            //行，插入、删除
            if (rowNum > table.NumRows - biaoTouRowsCount)
            {
#if cad2007
#else
                 if (table.CanInsertRow(table.NumRows))
#endif
                {
                    table.InsertRows(table.NumRows, rowHeight, rowNum - (table.NumRows - biaoTouRowsCount));
                }
            }
            if (rowNum < table.NumRows - biaoTouRowsCount)
            {
#if cad2007
#else
                 if (table.CanDeleteRows(table.NumRows - (table.NumRows - rowNum - biaoTouRowsCount), table.NumRows - rowNum - biaoTouRowsCount))
#endif
                {
                    table.DeleteRows(table.NumRows - (table.NumRows - rowNum - biaoTouRowsCount), table.NumRows - rowNum - biaoTouRowsCount);
                }
            }
        }
        /// <summary>
        /// 重新定义DataGridView行数列数
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="biaoTouRowsCount"></param>
        /// <param name="rowNum"></param>
        /// <param name="colNum"></param>
        /// <param name="rowHeight"></param>
        /// <param name="colWidth"></param>
        public static void ReSize(this DataGridView dataGridView, int rowNum, int colNum)
        {
            int oldRowNum = dataGridView.Rows.Count;
            int oldColumnNum = dataGridView.Columns.Count;
            int oldSelRow = ExcelClass.selectRow;
            int oldSelCol = ExcelClass.selectCol;
            noChangeExcelVar = true;//
            //列，插入、删除
            while (colNum > dataGridView.Columns.Count)
            {
                int curColNum = dataGridView.Columns.Count;
                dataGridView.Columns.Add("列_" + curColNum.ToString(), "列_" + curColNum.ToString());
                //dataGridView.Columns[dataGridView.Columns.Count - 1].Width = colWidth;//默认列宽
            }
            while (colNum < dataGridView.Columns.Count)
            {
                int curColNum = dataGridView.Columns.Count;
                dataGridView.Columns.RemoveAt(curColNum - 1);
            }
            //行，插入、删除
            if (rowNum > dataGridView.Rows.Count)
            {
                int curRowNum = dataGridView.Rows.Count;//原行数
                int add = rowNum - curRowNum;

                string[] sValue = new string[dataGridView.Columns.Count];
                //为最后一行赋空值，并把原记录保留到sValue[]
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    object obj = dataGridView[i, dataGridView.Rows.Count - 1].Value;
                    if (obj != null)
                        sValue[i] = obj.ToString();
                    else
                        sValue[i] = "";
                    dataGridView[i, dataGridView.Rows.Count - 1].Value = null;
                }
                //增加行
                for (int a = 0; a < add; a++)
                {
                    if (a == 0)
                        dataGridView.Rows.Add(sValue);
                    else
                        dataGridView.Rows.Add();
                }
            }
            if (rowNum < dataGridView.Rows.Count)
            {
                int curRowNum = dataGridView.Rows.Count;
                int cha = curRowNum - rowNum;
                //为最后一行重新赋值
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    dataGridView[i, dataGridView.Rows.Count - 1].Value = dataGridView[i,curRowNum-cha-1].Value;
                }
                //从倒数第二行，倒着删除行
                for (int i = 0; i < cha; i++)
                {
                    dataGridView.Rows.RemoveAt(curRowNum - cha-1);
                }
            }
            //禁止列头排序
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                dataGridView.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            dataGridView.Refresh();
            //判断是扩大了还是缩小了表格，然后设定当前单元格的位置。
            //扩大时位置不变，缩小时保留最大位置。
            if (dataGridView.Columns.Count - 1 < oldSelCol)
                ExcelClass.selectCol = dataGridView.Columns.Count - 1;
            if (dataGridView.Rows.Count - 1 < oldSelRow)
                ExcelClass.selectRow = dataGridView.Rows.Count - 1;
            dataGridView.CurrentCell = dataGridView[ExcelClass.selectCol, ExcelClass.selectRow];

            noChangeExcelVar = false;
        }

    }
}
