using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MyEntity;
using Util;

[assembly: CommandClass(typeof(ExcelForm.Commands))]
namespace ExcelForm
{
    /// <summary>
    /// CAD命令接口类
    /// </summary>
    public class Commands
    {
        /// <summary>
        /// 显示定制表格窗体窗体
        /// </summary>
        [CommandMethod("SCBG",CommandFlags.NoHistory)]
        public static void showForm()
        {
            ExcelForm ef = new ExcelForm();
            acadApp.ShowModelessDialog(ef);
        }
        /// <summary>
        /// 导出真表格
        /// </summary>
        [CommandMethod("DCZBG")]
        public void ExportTrueTable()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                Editor ed = doc.Editor;
                PromptSelectionOptions op = new PromptSelectionOptions();
                op.MessageForAdding = "\n可点选或框选多个表格";
                PromptSelectionResult ents = ed.GetSelection(op);
                if (ents.Status != PromptStatus.OK)
                {
                    return;
                }
                Autodesk.AutoCAD.EditorInput.SelectionSet selectResults = ents.Value;
                if (selectResults == null)
                {//用户没有选中任何
                    return;
                }
                IWorkbook wb = new HSSFWorkbook();
                int sheetIndex = 1;
                foreach (var item in selectResults.GetObjectIds())
                {
                    if (item == ObjectId.Null)
                    {
                        continue;
                    }
                    Entity ent = item.GetObject(OpenMode.ForRead) as Entity;
                    if ((ent as Table) != null)
                    {
                        List<TableCell> tablecellList = new List<TableCell>();
                        List<MergeRange> mergeRangeList = new List<MergeRange>();
                        Table tbl = ent as Table;
                        for (int i = 0; i < tbl.Rows.Count; i++)
                        {
                            for (int j = 0; j < tbl.Columns.Count; j++)
                            {
                                TableCell tableCell = new TableCell();
                                tableCell.RowIndex = i;
                                tableCell.ColumIndex = j;
                                string value = "";
                                if (tbl.Cells[i, j].Value != null)
                                {
                                    value = tbl.Cells[i, j].GetTextString(FormatOption.IgnoreMtextFormat);
                                }
                                MergeRange mergeRange = new MergeRange();
                                CellRange cellRange = tbl.Cells[i, j].GetMergeRange();//合并的信息
                                if (IsMergeRange(cellRange))
                                {
                                    tableCell.IsMergeRange = true;
                                    if (mergeRangeList.Where(o => o.BottomRow == cellRange.BottomRow && o.TopRow == cellRange.TopRow && o.LeftColumn == cellRange.LeftColumn && o.RightColumn == cellRange.RightColumn).Count() == 0)
                                    {
                                        mergeRange.TopRow = cellRange.TopRow;
                                        mergeRange.BottomRow = cellRange.BottomRow;
                                        mergeRange.LeftColumn = cellRange.LeftColumn;
                                        mergeRange.RightColumn = cellRange.RightColumn;
                                        mergeRangeList.Add(mergeRange);
                                    }
                                    else
                                    {
                                        mergeRange = null;
                                    }
                                }
                                else
                                {
                                    tableCell.IsMergeRange = false;
                                    mergeRange = null;
                                }
                                tableCell.Value = value;
                                tableCell.MergeRange = mergeRange;
                                tablecellList.Add(tableCell);
                            }
                        }
                        if (tablecellList.Count <= 0)
                        {
                            continue;
                        }
                        #region 生成文件
                        //创建表  
                        ISheet sh = wb.CreateSheet("sheet" + sheetIndex);
                        int rowindex = tablecellList.Last().RowIndex;
                        int columindex = tablecellList.Last().ColumIndex;
                        for (int i = 0; i <= rowindex; i++)
                        {
                            IRow row = sh.CreateRow(i);
                            for (int j = 0; j <= columindex; j++)
                            {
                                row.CreateCell(j);
                            }
                        }
                        foreach (var itemCell in tablecellList)
                        {
                            if (!(itemCell.IsMergeRange && itemCell.MergeRange == null))
                            {
                                sh.GetRow(itemCell.RowIndex).GetCell(itemCell.ColumIndex).SetCellValue(itemCell.Value);
                            }
                            if (itemCell.MergeRange != null
                                && itemCell.MergeRange.TopRow >= 0
                                && itemCell.MergeRange.BottomRow >= 0
                                && itemCell.MergeRange.LeftColumn >= 0
                                && itemCell.MergeRange.RightColumn >= 0)
                            {
                                sh.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(itemCell.MergeRange.TopRow, itemCell.MergeRange.BottomRow, itemCell.MergeRange.LeftColumn, itemCell.MergeRange.RightColumn));
                            }
                        }
                        #endregion
                        sheetIndex++;
                    }
                }
                SaveFileDialog fileDialog = new SaveFileDialog();
                fileDialog.Filter = "所有文件(*.xls)|*.xls|所有文件(*.xlsx)|*.xlsx"; //设置要选择的文件的类型 
                fileDialog.RestoreDirectory = true;
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    string file = fileDialog.FileName;//返回文件的完整路径 
                    using (FileStream stm = System.IO.File.OpenWrite(file))
                    {
                        wb.Write(stm);
                    }
                    MessageBox.Show("导出数据成功");
                }
            }
        }
        /// <summary>
        /// 导出假表格
        /// </summary>
        [CommandMethod("DCJBG")]
        public void ExportFakeTable()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Transaction transaction = doc.Database.TransactionManager.StartTransaction())
            {
                Editor ed = doc.Editor;

                Polyline3d pl = GetRegion();
                if (null == pl)
                {
                    return;
                }
                Extents3d ext3d = pl.GeometricExtents; //选择区域
                if (ext3d.MinPoint == new Extents3d().MinPoint && ext3d.MaxPoint == new Extents3d().MaxPoint)
                {
                    return;
                }
                //获取区域内所有对象
                PromptSelectionResult ents = ed.SelectCrossingWindow(ext3d.MinPoint, ext3d.MaxPoint);
                if (ents.Status != PromptStatus.OK)
                {
                    return;
                }
                ObjectId[] entsIn = new ObjectId[] { };
                PromptSelectionResult entsRIn = ed.SelectWindow(ext3d.MinPoint, ext3d.MaxPoint);
                if (entsRIn.Status == PromptStatus.OK && entsRIn.Value != null)
                {
                    entsIn = entsRIn.Value.GetObjectIds();
                }

                Autodesk.AutoCAD.EditorInput.SelectionSet selectResults = ents.Value;
                if (selectResults == null)
                {//用户没有选中任何
                    return;
                }
                //ed.WriteMessage("\n相交内："+ selectResults.GetObjectIds().Count()+"包含内"+ entsIn.Count());
                IWorkbook wb = new HSSFWorkbook();
                int sheetIndex = 1;
                Dictionary<Point3d, List<DBText>> dbtextList = new Dictionary<Point3d, List<DBText>>();
                Dictionary<Point3d, List<MText>> mtextList = new Dictionary<Point3d, List<MText>>();
                List<Line> verticalLine = new List<Line>();
                List<Line> horizontalLine = new List<Line>();
                foreach (var item in selectResults.GetObjectIds())
                {
                    if (item == ObjectId.Null)
                    {
                        continue;
                    }
                    Entity ent = item.GetObject(OpenMode.ForRead) as Entity;
                    DealEntity(ent, dbtextList, mtextList, verticalLine, horizontalLine, entsIn);
                }
                if (!(verticalLine.Count > 0 && horizontalLine.Count > 0))
                {
                    return;
                }
                #region 垂直 判断列数
                List<Line> verticalLineClone = new List<Line>();
                verticalLine.ForEach((item) => {
                    verticalLineClone.Add(item.Clone() as Line);
                });
                List<double> xList = new List<double>() { verticalLineClone[0].StartPoint.X };//记录X轴
                for (int i = 0; i < verticalLineClone.Count; i++)
                {
                    if (xList.Where(o => Math.Abs(o - verticalLineClone[i].StartPoint.X) > 0.1).Count() >= xList.Count)
                    {//垂直线 间距 大于0.1的间距，判定为一列
                        xList.Add(verticalLineClone[i].StartPoint.X);
                        verticalLineClone.RemoveAt(i);
                        i--;
                    }
                }

                #endregion
                #region 水平 判断行数  
                List<Line> horizontalLineClone = new List<Line>();
                horizontalLine.ForEach((item) => {
                    horizontalLineClone.Add(item.Clone() as Line);
                });
                List<double> yList = new List<double>() { horizontalLineClone[0].StartPoint.Y };//记录Y轴
                for (int i = 0; i < horizontalLineClone.Count; i++)
                {
                    if (yList.Where(o => Math.Abs(o - horizontalLineClone[i].StartPoint.Y) > 0.1).Count() >= yList.Count)
                    {//水平线 间距 大于0.1的间距，判定为一行
                        yList.Add(horizontalLineClone[i].StartPoint.Y);
                        horizontalLineClone.RemoveAt(i);
                        i--;
                    }
                }
                yList = yList.Distinct().OrderByDescending(o => o).ToList();
                #endregion
                #region 去掉与选择区域边框相交的 线段
                for (int i = 0; i < verticalLine.Count; i++)
                {
                    Point3dCollection pc = new Point3dCollection();
                    verticalLine[i].IntersectWith(pl, Intersect.OnBothOperands, pc, 0, 0);
                    if (pc.Count > 0)
                    {
                        if (isPointInRegion(ext3d, verticalLine[i].StartPoint) || isPointInRegion(ext3d, verticalLine[i].EndPoint))
                        {//水平线的起始点在 选择范围内
                            if (yList.Select(o => Math.Abs(o - verticalLine[i].StartPoint.Y)).Min() > 0.1 && yList.Select(o => Math.Abs(o - verticalLine[i].EndPoint.Y)).Min() < 0.01)
                            {
                                bool isRemove = true;
                                foreach (var item in verticalLine)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    verticalLine[i].IntersectWith(item, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count > 0)
                                    {//有交点就不删除了
                                        isRemove = false;
                                        break;
                                    }
                                }
                                if (isRemove)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    verticalLine[i].IntersectWith(pl, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count <= 1)
                                    {
                                        xList.Remove(verticalLine[i].StartPoint.X);
                                    }
                                }
                            }
                            else if (yList.Select(o => Math.Abs(o - verticalLine[i].StartPoint.Y)).Min() < 0.01 && yList.Select(o => Math.Abs(o - verticalLine[i].EndPoint.Y)).Min() > 0.1)
                            {//x轴采集到的 位置
                                bool isRemove = true;
                                foreach (var item in verticalLine)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    verticalLine[i].IntersectWith(item, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count > 0)
                                    {
                                        isRemove = false;
                                        break;
                                    }
                                }
                                if (isRemove)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    verticalLine[i].IntersectWith(pl, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count <= 1)
                                    {
                                        xList.Remove(verticalLine[i].StartPoint.X);
                                    }
                                }
                            }
                        }
                    }
                }
                for (int i = 0; i < horizontalLine.Count; i++)
                {
                    Point3dCollection pc = new Point3dCollection();
                    horizontalLine[i].IntersectWith(pl, Intersect.OnBothOperands, pc, 0, 0);
                    if (pc.Count > 0)
                    {
                        if (isPointInRegion(ext3d, horizontalLine[i].StartPoint) || isPointInRegion(ext3d, horizontalLine[i].EndPoint))
                        {//水平线的起始点在 选择范围内
                            if (xList.Select(o => Math.Abs(o - horizontalLine[i].StartPoint.X)).Min() > 0.1 && xList.Select(o => Math.Abs(o - horizontalLine[i].EndPoint.X)).Min() < 0.01)
                            {//x轴采集到的 位置
                                bool isRemove = true;
                                foreach(var item in horizontalLine)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    horizontalLine[i].IntersectWith(item, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count > 0)
                                    {
                                        isRemove = false;
                                        break;
                                    }
                                }
                                if(isRemove)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    horizontalLine[i].IntersectWith(pl, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count <= 0)
                                    {
                                        yList.Remove(horizontalLine[i].StartPoint.Y);
                                    }
                                }
                            }
                            else if (xList.Select(o => Math.Abs(o - horizontalLine[i].StartPoint.X)).Min() < 0.01 && xList.Select(o => Math.Abs(o - horizontalLine[i].EndPoint.X)).Min() > 0.1)
                            {//x轴采集到的 位置
                                bool isRemove = true;
                                foreach (var item in horizontalLine)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    horizontalLine[i].IntersectWith(item, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count > 0)
                                    {
                                        isRemove = false;
                                        break;
                                    }
                                }
                                if (isRemove)
                                {
                                    Point3dCollection pcItem = new Point3dCollection();
                                    horizontalLine[i].IntersectWith(pl, Intersect.OnBothOperands, pcItem, 0, 0);
                                    if (pcItem.Count <= 0)
                                    {
                                        yList.Remove(horizontalLine[i].StartPoint.Y);
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
                //xList 从左到右排序
                xList = xList.Distinct().OrderBy(o => o).ToList();
                //yList 从上到下排序
                yList = yList.Distinct().OrderByDescending(o => o).ToList();
                if (!(xList.Count > 1 && yList.Count > 1))
                {
                    MessageBox.Show("未识别到表格，请扩大选择范围");
                    return;
                }
                double xL = Math.Abs(xList.ElementAt(1) - xList.ElementAt(0));
                double xR = Math.Abs(xList.ElementAt(xList.Count - 1) - xList.ElementAt(xList.Count - 2));
                double yT = Math.Abs(yList.ElementAt(0) - yList.ElementAt(1));
                double yB = Math.Abs(yList.ElementAt(yList.Count - 2) - yList.ElementAt(yList.Count - 1));

                Line lineL = new Line(ext3d.MinPoint, new Point3d(ext3d.MinPoint.X, ext3d.MaxPoint.Y, 0));
                Line lineR = new Line(new Point3d(ext3d.MaxPoint.X, ext3d.MinPoint.Y, 0), ext3d.MaxPoint);
                Line lineT = new Line(new Point3d(ext3d.MinPoint.X, ext3d.MaxPoint.Y, 0), ext3d.MaxPoint);
                Line lineB = new Line(ext3d.MinPoint, new Point3d(ext3d.MaxPoint.X, ext3d.MinPoint.Y, 0));
                
                TypedValue[] acTypValAr = new TypedValue[5];
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Operator, "<or"), 0);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "Line"), 1);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 2);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 3);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Operator, "or>"), 4);
                SelectionFilter acSelFtr = new SelectionFilter(acTypValAr);
                List<TableCell> tablecellList = new List<TableCell>();
                List<MergeRange> mergeRangeList = new List<MergeRange>();
                Dictionary<Point3d, MergeRange> dic = new Dictionary<Point3d, MergeRange>();
                try
                {

                    for (int y = 0; y < yList.Count - 1; y++)
                    {
                        for (int x = 0; x < xList.Count - 1; x++)
                        {
                            double xmid1 = (xList[x] + (xList[x + 1] - xList[x]) / 2);//当前单元格中间x轴位置
                            double ymid1 = (yList[y] - (yList[y] - yList[y + 1]) / 2);//当前单元格中间y轴位置
                            Point3d point1 = new Point3d(xmid1, ymid1, 0);
                            Point3d pointx2 = Point3d.Origin;
                            Point3d pointy2 = Point3d.Origin;
                            bool isX = false;//X轴是否有交集
                            if (x + 2 < xList.Count)
                            {
                                #region xList
                                double xmid2 = (xList[x + 1] + (xList[x + 2] - xList[x + 1]) / 2);
                                pointx2 = new Point3d(xmid2, ymid1, 0);
                                PromptSelectionResult crossing = ed.SelectCrossingWindow(point1, pointx2, acSelFtr);
                                if (crossing.Status == PromptStatus.OK)
                                {
                                    SelectionSet ss = crossing.Value;
                                    ObjectId[] objectidcollectionX = ss.GetObjectIds();
                                    List<Line> lines = new List<Line>();
                                    foreach (var objectid in objectidcollectionX)
                                    {
                                        Entity ent = objectid.GetObject(OpenMode.ForRead) as Entity;
                                        if (ent.GeometricExtents.MinPoint == new Extents3d().MinPoint && ent.GeometricExtents.MaxPoint == new Extents3d().MaxPoint)
                                        {
                                            continue;
                                        }
                                        DealEntity(ent, lines, pl, ext3d, lineL, lineR, xL, xR, entsIn);
                                    }
                                    foreach (var o in lines)
                                    {
                                        if (verticalLine.Where(l => l.StartPoint == o.StartPoint && l.EndPoint == o.EndPoint).Count() > 0)
                                        {
                                            isX = true;
                                            break;
                                        }
                                    }
                                }
                                #endregion
                            }
                            else
                            {//最边上的单元格
                                isX = true;
                            }
                            bool isY = false;//Y轴是否有交集
                            if (y + 2 < yList.Count)
                            {
                                #region yList
                                double ymid2 = (yList[y + 1] + (yList[y + 2] - yList[y + 1]) / 2);
                                pointy2 = new Point3d(xmid1, ymid2, 0);
                                PromptSelectionResult crossing = ed.SelectCrossingWindow(point1, pointy2, acSelFtr);
                                if (crossing.Status == PromptStatus.OK)
                                {
                                    SelectionSet ss = crossing.Value;
                                    ObjectId[] objectidcollectionY = ss.GetObjectIds();

                                    List<Line> lines = new List<Line>();
                                    foreach (var objectid in objectidcollectionY)
                                    {
                                        Entity ent = objectid.GetObject(OpenMode.ForRead) as Entity;
                                        if (ent.GeometricExtents.MinPoint == new Extents3d().MinPoint && ent.GeometricExtents.MaxPoint == new Extents3d().MaxPoint)
                                        {
                                            continue;
                                        }
                                        DealEntity(ent, lines, pl, ext3d, lineT, lineB, yT, yB, entsIn);
                                    }
                                    foreach (var o in lines)
                                    {
                                        if (horizontalLine.Where(l => l.StartPoint == o.StartPoint && l.EndPoint == o.EndPoint).Count() > 0)
                                        {
                                            isY = true;
                                            break;
                                        }
                                    }
                                }
                                #endregion
                            }
                            else
                            {//最边上的单元格
                                isY = true;
                            }
                            TableCell tableCell = new TableCell();
                            tableCell.RowIndex = y;
                            tableCell.ColumIndex = x;

                            if (dic.Values.Any(o => y >= o.TopRow && y <= o.BottomRow && x >= o.LeftColumn && x <= o.RightColumn))
                            {//之前统计过的合并单元格
                                tableCell.IsMergeRange = true;
                                tableCell.MergeRange = null;
                                tablecellList.Add(tableCell);
                                continue;
                            }
                            if ((isX && isY))
                            {//非合并的单元格
                                tableCell.IsMergeRange = false;
                                tableCell.MergeRange = null;
                                var tempdbtextList = dbtextList.Where(o => o.Key.X >= xList[x] && o.Key.X <= xList[x + 1] && o.Key.Y >= yList[y + 1] && o.Key.Y <= yList[y]).ToList();
                                foreach (var itemDBText1 in tempdbtextList.GroupBy(o => o.Key.Y).OrderByDescending(o => o.Key))
                                {
                                    foreach (var itemDBText2 in itemDBText1.OrderBy(o => o.Key.X))
                                    {
                                        itemDBText2.Value.ForEach((dbtext) => {
                                            tableCell.Value += dbtext.TextString;
                                        });
                                    }
                                    tableCell.Value += "\n";
                                }
                                var tempmtextList = mtextList.Where(o => o.Key.X >= xList[x] && o.Key.X <= xList[x + 1] && o.Key.Y >= yList[y + 1] && o.Key.Y <= yList[y]).ToList();
                                foreach (var itemMText1 in tempmtextList.GroupBy(o => o.Key.Y).OrderByDescending(o => o.Key))
                                {
                                    foreach (var itemMText2 in itemMText1.OrderBy(o => o.Key.X))
                                    {
                                        itemMText2.Value.ForEach((dbtext) => {
                                            tableCell.Value += dbtext.Text;
                                        });
                                    }
                                    tableCell.Value += "\n";
                                }
                                if (!string.IsNullOrEmpty(tableCell.Value))
                                {
                                    tableCell.Value.TrimEnd('\n');
                                }
                                tablecellList.Add(tableCell);
                            }
                            else
                            {//发生合并
                                MergeRange mr = new MergeRange() { TopRow = y, LeftColumn = x, BottomRow = y, RightColumn = x };
                                if (!isX)
                                {
                                    mr.RightColumn++;
                                    dic.addIsCover(point1, mr, true);
                                    SetMergeRangeX(mr, xList, x + 1, ed, pointx2, acSelFtr, verticalLine, dic, point1);
                                }
                                if (!isY)
                                {
                                    mr.BottomRow++;
                                    dic.addIsCover(point1, mr, true);
                                    SetMergeRangeY(mr, yList, y + 1, ed, pointy2, acSelFtr, horizontalLine, dic, point1);
                                }
                                tableCell.IsMergeRange = true;
                                tableCell.MergeRange = mr;
                                int cNum = (mr.RightColumn - mr.LeftColumn);//列跨度
                                int rNum = (mr.BottomRow - mr.TopRow);//行跨度
                                var tempdbtextList = dbtextList.Where(o => o.Key.X >= xList[x] && o.Key.X <= xList[x + 1 + cNum] && o.Key.Y >= yList[y + 1 + rNum] && o.Key.Y <= yList[y]).ToList();
                                foreach (var itemDBText1 in tempdbtextList.GroupBy(o => o.Key.Y).OrderByDescending(o => o.Key))
                                {
                                    foreach (var itemDBText2 in itemDBText1.OrderBy(o => o.Key.X))
                                    {
                                        itemDBText2.Value.ForEach((dbtext) => {
                                            tableCell.Value += dbtext.TextString;
                                        });
                                    }
                                    tableCell.Value += "\n";
                                }
                                var tempmtextList = mtextList.Where(o => o.Key.X >= xList[x] && o.Key.X <= xList[x + 1 + cNum] && o.Key.Y >= yList[y + 1 + rNum] && o.Key.Y <= yList[y]).ToList();
                                foreach (var itemMText1 in tempmtextList.GroupBy(o => o.Key.Y).OrderByDescending(o => o.Key))
                                {
                                    foreach (var itemMText2 in itemMText1.OrderBy(o => o.Key.X))
                                    {
                                        itemMText2.Value.ForEach((dbtext) => {
                                            tableCell.Value += dbtext.Text;
                                        });
                                    }
                                    tableCell.Value += "\n";
                                }
                                if (!string.IsNullOrEmpty(tableCell.Value))
                                {
                                    tableCell.Value.TrimEnd('\n');
                                }
                                tablecellList.Add(tableCell);
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {

                }
                if (tablecellList.Count <= 0)
                {
                    MessageBox.Show("未识别到表格，请扩大选择范围");
                    return;
                }
                #region 生成文件
                //创建表  
                ISheet sh = wb.CreateSheet("sheet" + sheetIndex);
                int rowindex = tablecellList.Last().RowIndex;
                int columindex = tablecellList.Last().ColumIndex;
                for (int i = 0; i <= rowindex; i++)
                {
                    IRow row = sh.CreateRow(i);
                    for (int j = 0; j <= columindex; j++)
                    {
                        row.CreateCell(j);
                    }
                }
                foreach (var itemCell in tablecellList)
                {
                    if (!(itemCell.IsMergeRange && itemCell.MergeRange == null))
                    {
                        sh.GetRow(itemCell.RowIndex).GetCell(itemCell.ColumIndex).SetCellValue(itemCell.Value);
                    }
                    if (itemCell.MergeRange != null
                        && itemCell.MergeRange.TopRow >= 0
                        && itemCell.MergeRange.BottomRow >= 0
                        && itemCell.MergeRange.LeftColumn >= 0
                        && itemCell.MergeRange.RightColumn >= 0)
                    {
                        sh.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(itemCell.MergeRange.TopRow, itemCell.MergeRange.BottomRow, itemCell.MergeRange.LeftColumn, itemCell.MergeRange.RightColumn));
                    }
                }

                SaveFileDialog fileDialog = new SaveFileDialog();
                fileDialog.Filter = "所有文件(*.xls)|*.xls|所有文件(*.xlsx)|*.xlsx"; //设置要选择的文件的类型 
                fileDialog.RestoreDirectory = true;
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    string file = fileDialog.FileName;//返回文件的完整路径 
                    using (FileStream stm = System.IO.File.OpenWrite(file))
                    {
                        wb.Write(stm);
                    }
                    MessageBox.Show("导出数据成功");
                }
                #endregion

            }
        }

        private void DealEntity(Entity ent, Dictionary<Point3d, List<DBText>> dbtextList, Dictionary<Point3d, List<MText>> mtextList, List<Line> verticalLine, List<Line> horizontalLine, ObjectId[] entsIn = null)
        {
            #region 筛选垂直与水平直线  verticalLine  horizontalLine
            if (ent is Line)
            {
                #region Line
                Line lineH = GetVerticalHorizontalLine(ent as Line, "H");
                if (lineH != null)
                {
                    horizontalLine.Add(lineH);
                }
                else
                {
                    Line lineV = GetVerticalHorizontalLine(ent as Line, "V");
                    if (lineV != null)
                    {
                        verticalLine.Add(lineV);
                    }
                }
                #endregion
            }
            else if (ent is Polyline)
            {
                #region Polyline
                Polyline polyline = (ent as Polyline);

                DBObjectCollection dBObjectCollection = new DBObjectCollection();
                ent.Explode(dBObjectCollection);
                foreach (var item in dBObjectCollection)
                {
                    DealEntity(item as Entity, dbtextList, mtextList, verticalLine, horizontalLine);
                }

                //for (int i = 0; i < polyline.NumberOfVertices - 1; i++)
                //{
                //    Line line = new Line(polyline.GetPoint3dAt(i), polyline.GetPoint3dAt(i + 1));
                //    Line lineH = GetVerticalHorizontalLine(line, "H");
                //    if (lineH != null)
                //    {
                //        horizontalLine.Add(lineH);
                //    }
                //    else
                //    {
                //        Line lineV = GetVerticalHorizontalLine(line, "V");
                //        if (lineV != null)
                //        {
                //            verticalLine.Add(lineV);
                //        }
                //    }
                //}
                #endregion
            }
            #endregion
            #region 筛选文本  dbtextList  mtextList 
            else if (ent is DBText)
            {
                dbtextList.add((ent as DBText).Position, ent as DBText);
            }
            else if (ent is MText)
            {
                mtextList.add((ent as MText).Location, ent as MText);
            }
            #endregion
            else if (ent is BlockReference)
            {
                if (entsIn != null && entsIn.Where(o => o.Handle.ToString().Equals(ent.Handle.ToString())).Count() > 0)
                {//仅处理完全包含在 选择区域内的块参照对象
                    DBObjectCollection dBObjectCollection = new DBObjectCollection();
                    ent.Explode(dBObjectCollection);
                    foreach (var item in dBObjectCollection)
                    {
                        DealEntity(item as Entity, dbtextList, mtextList, verticalLine, horizontalLine);
                    }
                }
                else if (entsIn == null)
                {
                    DBObjectCollection dBObjectCollection = new DBObjectCollection();
                    ent.Explode(dBObjectCollection);
                    foreach (var item in dBObjectCollection)
                    {
                        DealEntity(item as Entity, dbtextList, mtextList, verticalLine, horizontalLine);
                    }
                }
            }
        }
        private void DealEntity(Entity ent, List<Line> lines, Polyline3d pl, Extents3d ext3d, Line line1, Line line2, double d1, double d2, ObjectId[] entsIn = null)
        {
            if (ent is Line)
            {
                Point3dCollection pc = new Point3dCollection();
                ent.IntersectWith(pl, Intersect.OnBothOperands, pc, 0, 0);
                if (pc.Count > 0)
                {
                    if (ent is Line)
                    {//当相交的部分在 选择范围内的长度 小于 两边的单元长度，不计入统计
                        Line lineTemp = ent as Line;
                        if (isPointInRegion(ext3d, lineTemp.StartPoint))
                        {
                            if (line1.Normal.GetAngleTo(lineTemp.Normal) > 0.5 / 180 * Math.PI && (line1.GetClosestPointTo(pc[0], false).DistanceTo(pc[0]) < 0.00001) && d1 > pc[0].DistanceTo(lineTemp.StartPoint))
                            {
                                return;
                            }
                            else if (line2.Normal.GetAngleTo(lineTemp.Normal) > 0.5 / 180 * Math.PI && (line2.GetClosestPointTo(pc[0], false).DistanceTo(pc[0]) < 0.00001) && d2 > pc[0].DistanceTo(lineTemp.StartPoint))
                            {
                                return;
                            }
                        }
                        else
                        {
                            if (line1.Normal.GetAngleTo(lineTemp.Normal) > 0.5 / 180 * Math.PI && (line1.GetClosestPointTo(pc[0], false).DistanceTo(pc[0]) < 0.00001) && d1 > pc[0].DistanceTo(lineTemp.EndPoint))
                            {
                                return;
                            }
                            else if (line2.Normal.GetAngleTo(lineTemp.Normal) > 0.5 / 180 * Math.PI && (line2.GetClosestPointTo(pc[0], false).DistanceTo(pc[0]) < 0.00001) && d2 > pc[0].DistanceTo(lineTemp.EndPoint))
                            {
                                return;
                            }
                        }
                    }
                }
                lines.Add(ent as Line);
            }
            else if (ent is Polyline)
            {
                DBObjectCollection dBObjectCollection = new DBObjectCollection();
                ent.Explode(dBObjectCollection);
                foreach (var item in dBObjectCollection)
                {
                    DealEntity(item as Entity, lines, pl, ext3d, line1, line2, d1, d2);
                }
                //lines.AddRange(GetLineByPolyline(ent as Polyline));
            }
            else if (ent is BlockReference)
            {
                if (entsIn != null && entsIn.Where(o => o.Handle.ToString().Equals(ent.Handle.ToString())).Count() > 0)
                {
                    DBObjectCollection dBObjectCollection = new DBObjectCollection();
                    ent.Explode(dBObjectCollection);
                    foreach (var item in dBObjectCollection)
                    {
                        DealEntity(item as Entity, lines, pl, ext3d, line1, line2, d1, d2);
                    }
                }
                else if (entsIn == null)
                {
                    DBObjectCollection dBObjectCollection = new DBObjectCollection();
                    ent.Explode(dBObjectCollection);
                    foreach (var item in dBObjectCollection)
                    {
                        DealEntity(item as Entity, lines, pl, ext3d, line1, line2, d1, d2);
                    }
                }

            }
        }

        private List<Line> GetLineByPolyline(Polyline polyline)
        {
            List<Line> lineList = new List<Line>();
            for (int i = 0; i < polyline.NumberOfVertices - 1; i++)
            {
                lineList.Add(new Line(polyline.GetPoint3dAt(i), polyline.GetPoint3dAt(i + 1)));
            }
            return lineList;
        }

        private void SetMergeRangeX(MergeRange mr, List<double> xList, int x, Editor ed, Point3d pointx2, SelectionFilter acSelFtr, List<Line> verticalLine, Dictionary<Point3d, MergeRange> dic, Point3d basePoint3d)
        {
            if (xList.Count <= (x + 2))
            {
                return;
            }
            double xmid3 = (xList[x + 1] + (xList[x + 2] - xList[x + 1]) / 2);
            Point3d pointx3 = new Point3d(xmid3, pointx2.Y, 0);
            PromptSelectionResult crossing = ed.SelectCrossingWindow(pointx2, pointx3, acSelFtr);
            bool isX = false;
            if (crossing.Status == PromptStatus.OK)
            {
                SelectionSet ss = crossing.Value;
                ObjectId[] objectidcollectionX = ss.GetObjectIds();
                List<Line> lines = new List<Line>();
                foreach (var objectid in objectidcollectionX)
                {
                    Entity ent = objectid.GetObject(OpenMode.ForRead) as Entity;
                    if (ent is Line)
                    {
                        lines.Add(ent as Line);
                    }
                    else if (ent is Polyline)
                    {
                        lines.AddRange(GetLineByPolyline(ent as Polyline));
                    }
                }
                foreach (var o in lines)
                {
                    if (verticalLine.Where(l => l.StartPoint == o.StartPoint && l.EndPoint == o.EndPoint).Count() > 0)
                    {
                        isX = true;
                        break;
                    }
                }

                //isX = verticalLine.Where(o => objectidcollectionX.Contains(o.ObjectId)).Any();
                if (!isX)
                {
                    mr.RightColumn++;
                    dic.addIsCover(basePoint3d, mr, true);
                    SetMergeRangeX(mr, xList, x + 1, ed, pointx3, acSelFtr, verticalLine, dic, basePoint3d);
                }
            }
        }
        private void SetMergeRangeY(MergeRange mr, List<double> yList, int y, Editor ed, Point3d pointy2, SelectionFilter acSelFtr, List<Line> horizontalLine, Dictionary<Point3d, MergeRange> dic, Point3d basePoint3d)
        {
            if (yList.Count <= (y + 2))
            {
                return;
            }
            double ymid3 = (yList[y + 1] + (yList[y + 2] - yList[y + 1]) / 2);
            Point3d pointy3 = new Point3d(pointy2.X, ymid3, 0);
            PromptSelectionResult crossing = ed.SelectCrossingWindow(pointy2, pointy3, acSelFtr);
            bool isY = false;
            if (crossing.Status == PromptStatus.OK)
            {
                SelectionSet ss = crossing.Value;
                ObjectId[] objectidcollectionY = ss.GetObjectIds();

                List<Line> lines = new List<Line>();
                foreach (var objectid in objectidcollectionY)
                {
                    Entity ent = objectid.GetObject(OpenMode.ForRead) as Entity;
                    if (ent is Line)
                    {
                        lines.Add(ent as Line);
                    }
                    else if (ent is Polyline)
                    {
                        lines.AddRange(GetLineByPolyline(ent as Polyline));
                    }
                }
                foreach (var o in lines)
                {
                    if (horizontalLine.Where(l => l.StartPoint == o.StartPoint && l.EndPoint == o.EndPoint).Count() > 0)
                    {
                        isY = true;
                        break;
                    }
                }

                //isY = horizontalLine.Where(o => objectidcollectionY.Contains(o.ObjectId)).Any();
                if (!isY)
                {
                    mr.BottomRow++;
                    dic.addIsCover(basePoint3d, mr, true);
                    SetMergeRangeY(mr, yList, y + 1, ed, pointy3, acSelFtr, horizontalLine, dic, basePoint3d);
                }
            }
        }

        private Line GetVerticalHorizontalLine(Line line, string type)
        {
            double angle = line.Angle * 180 / Math.PI;
            if (type.Equals("H"))
            {
                if ((angle >= -0.5 && angle <= 0.5) || (angle >= 359.5 && angle <= 360) || (angle >= 179.5 && angle <= 180.5))
                {//水平  转换成从左到右
                    if (line.StartPoint.X > line.EndPoint.X)
                    {
                        if (line.ObjectId != ObjectId.Null && !line.IsWriteEnabled)
                        {
                            line = line.ObjectId.GetObject(OpenMode.ForWrite) as Line;
                        }
                        Line temp = new Line(line.StartPoint, line.EndPoint);
                        line.StartPoint = temp.EndPoint;
                        line.EndPoint = temp.StartPoint;
                        if (line.ObjectId != ObjectId.Null)
                        {
                            line.DowngradeOpen();
                        }
                        return line;
                    }
                    return line;
                }
            }
            else if (type.Equals("V"))
            {
                if ((angle >= 89.5 && angle <= 90.5) || (angle >= 269.5 && angle <= 270.5))
                {//垂直  装换成从上到下
                    if (line.EndPoint.Y > line.StartPoint.Y)
                    {
                        if (line.ObjectId != ObjectId.Null && !line.IsWriteEnabled)
                        {
                            line = line.ObjectId.GetObject(OpenMode.ForWrite) as Line;
                        }
                        Line temp = new Line(line.StartPoint, line.EndPoint);
                        line.StartPoint = temp.EndPoint;
                        line.EndPoint = temp.StartPoint;
                        if (line.ObjectId != ObjectId.Null)
                        {
                            line.DowngradeOpen();
                        }
                        return line;
                    }
                    return line;
                }
            }
            return null;
        }
        private bool IsMergeRange(CellRange cr)
        {
            if (cr != null)
            {
                if (cr.BottomRow == -1 && cr.TopRow == -1 && cr.RightColumn == -1 && cr.LeftColumn == -1)
                {
                    return false;
                }
                return true;
            }
            else
            {
                return false;
            }
        }

        public Polyline3d GetRegion(string firstMessage = "请选择区域", string secondMessage = "请选择区域")
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (DocumentLock acLckDoc = doc.LockDocument())
            {
                Editor editor = doc.Editor;
                Point3d firstPoint = GetPoint(firstMessage);
                if (firstPoint == Point3d.Origin)
                {
                    //用户取消操作
                    return null;
                }
                Point3d secondPoint;
                PromptCornerOptions options = new PromptCornerOptions(secondMessage, firstPoint);
                PromptPointResult i = editor.GetCorner(options);
                if (i.Status == PromptStatus.OK)
                {
                    secondPoint = i.Value;
                }
                else
                {
                    return null;
                }
                Point3dCollection points = new Point3dCollection();
                points.Add(firstPoint);
                points.Add(new Point3d(secondPoint.X, firstPoint.Y, 0));
                points.Add(secondPoint);
                points.Add(new Point3d(firstPoint.X, secondPoint.Y, 0));
                return new Polyline3d(Poly3dType.SimplePoly, points, true);
            }
        }
        public Point3d GetPoint(string mess)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;
            PromptPointOptions prPointOptions = new PromptPointOptions(mess);
            PromptPointResult prPointRes = editor.GetPoint(prPointOptions);
            return prPointRes.Status == PromptStatus.OK ? prPointRes.Value : Point3d.Origin;
        }

        public static Point2d toPoint2d(Point3d pt)
        {
            return new Point2d(pt.X, pt.Y);
        }

        #region 几何计算
        public static bool isPointInRegion(Extents3d ext, Point3d point)
        {//利用向量叉乘判断，逆时针
            return isPointInRegion(ext, toPoint2d(point));
        }
        public static bool isPointInRegion(Extents3d ext, Point2d point)
        {
            int num = 0;
            num += GetClockWise(toPoint2d(ext.MaxPoint), point, new Point2d(ext.MaxPoint.X, ext.MinPoint.Y));
            num += GetClockWise(new Point2d(ext.MinPoint.X, ext.MaxPoint.Y), point, toPoint2d(ext.MaxPoint));
            num += GetClockWise(toPoint2d(ext.MinPoint), point, new Point2d(ext.MinPoint.X, ext.MaxPoint.Y));
            num += GetClockWise(new Point2d(ext.MaxPoint.X, ext.MinPoint.Y), point, toPoint2d(ext.MinPoint));
            return num == 4 ? true : false;
        }
        /// <summary>
        /// 判断矩形B是否与图框A相交
        /// 如果两个矩形相交，那么矩形A B的中心点和矩形的边长是有一定关系的。两个中心点间的距离肯定小于AB边长和的一半
        /// </summary>
        /// <param name="extentsA"></param>
        /// <param name="extentsB"></param>
        /// <returns>1:相交  0:包含  -1:无交集</returns>
        public static int RectangleIsInsert(Extents3d extentsA, Extents3d extentsB)
        {
            double zx = Math.Abs(extentsA.MinPoint.X + extentsA.MaxPoint.X - extentsB.MinPoint.X - extentsB.MaxPoint.X);
            double x = Math.Abs(extentsA.MinPoint.X - extentsA.MaxPoint.X) + Math.Abs(extentsB.MinPoint.X - extentsB.MaxPoint.X);
            double zy = Math.Abs(extentsA.MinPoint.Y + extentsA.MaxPoint.Y - extentsB.MinPoint.Y - extentsB.MaxPoint.Y);
            double y = Math.Abs(extentsA.MinPoint.Y - extentsA.MaxPoint.Y) + Math.Abs(extentsB.MinPoint.Y - extentsB.MaxPoint.Y);
            if (zx <= x && zy <= y)
            {
                if (!(isPointInRegion(extentsA, extentsB.MinPoint) && isPointInRegion(extentsA, extentsB.MaxPoint))
                    && !(isPointInRegion(extentsB, extentsA.MinPoint) && isPointInRegion(extentsB, extentsA.MaxPoint)))
                {
                    return 1;
                }
                return 0;
            }
            else
                return -1;

        }
        /// <summary>判断点在线的内部还是外部    内部或线上返回1，外部返回-1
        /// </summary>
        /// <param name="p1">基准点</param>
        /// <param name="p2">要判断的点</param>
        /// <param name="p3">原点</param>
        /// <returns></returns>
        public static int GetClockWise(Point2d p1, Point2d p2, Point2d p3)
        {
            if ((p1.X - p3.X) * (p2.Y - p3.Y) - (p2.X - p3.X) * (p1.Y - p3.Y) >= 0)
            {//逆时针
                return 1;
            }
            return -1;
        }
        #endregion

    }
}