using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using DotNetARX;

namespace ExcelForm
{
    public class TableClass
    {
        public void setValue(int rowNum,int colNum,int rowHeight,int colWidth,string[] sArr)
        {
            ExcelClass.rowNum = rowNum;
            ExcelClass.colNum = colNum;
            ExcelClass.rowHeight = rowHeight;
            ExcelClass.colWidth = colWidth;
            ExcelClass.headString = sArr.ToList();
        }

        [CommandMethod("dt")]
        public void drawTable()
        {
            try
            {
                //测试结束后 下边4行注释掉
                string[] sArr = new string[] { "表头1", "表头2", "表头3", "表头4" };
                setValue(10, sArr.Length, 2, 12, sArr);
                ExcelClass.selectRow = 5;
                ExcelClass.selectCol = 2;
                //
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                Editor ed = doc.Editor;

                ObjectId currentTableId = createTable(db);//创建表格

                PromptNestedEntityResult entResult = ed.GetNestedEntity("\n请选择文字：");
                while (entResult.Status == PromptStatus.OK)
                {
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        Table currentTable = trans.GetObject(currentTableId,OpenMode.ForRead) as Table;
                        DBObject obj = entResult.ObjectId.GetObject(OpenMode.ForRead);
                        DBText txtobj = obj as DBText;
                        string txt = "";
                        if (txtobj != null)
                            txt = txtobj.TextString;
                        else
                        {
                            MText mtxt = obj as MText;
                            if (mtxt != null)
                                txt = mtxt.Contents;
                            else
                                txt = "";
                        }
                        if (txt.Length != 0)
                        {
                            ed.WriteMessage(txt);
                            //判断是否修改表格行数列数
                            currentTable.UpgradeOpen();//打开编辑
                             //如果选择行>=表格行数或者选择列>=表格列数，根据ExcelClass的行数列数重新定义表格行数列数
                            if (ExcelClass.selectCol >= currentTable.Columns.Count || ExcelClass.selectRow >= currentTable.Rows.Count)
                            {
                                currentTable.SetSize(ExcelClass.rowNum, ExcelClass.colNum);//根据行个数，列个数扩大
                                currentTable.SetRowHeight(ExcelClass.rowHeight);
                            }
                            //对应单元格设定为选择的文字
                            currentTable.Cells[ExcelClass.selectRow, ExcelClass.selectCol].TextString = txt;
                            currentTable.DowngradeOpen();//关闭编辑
                        }
                        else
                            ed.WriteMessage("\n您选择的不是文字，请重新选择!");
                        trans.Commit();
                    }
                    entResult = ed.GetNestedEntity("\n请选择文字：");
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                Application.ShowAlertDialog("发生错误:"+ex.Message);
            }
        }

        private ObjectId createTable(Database db)
        {
            ObjectId objID = ObjectId.Null;
            List<string> btList = ExcelClass.headString as List<string>;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                ObjectId styleId = AddTableStyle("MyTable");//get样式，若未定义则新建
                Table table = new Table();
                table.TableStyle = styleId;
                //Application.ShowAlertDialog(objID.ToString());
                table.Position = Point3d.Origin;
                table.SetSize(ExcelClass.rowNum, ExcelClass.colNum);//设定行/列总数
                table.SetRowHeight(ExcelClass.rowHeight);     //设定行高
                table.SetColumnWidth(ExcelClass.colWidth);    //设定列宽
                table.Cells[0, 0].TextString = "标题行";
                for (int i = 0; i < btList.Count; i++)
                {
                    //设置表头文字
                    table.Cells[1, i].TextString = btList[i];
                }
                objID =db.AddToModelSpace(table);
                trans.Commit();
            }
            return objID;
        }

        //为当前图形添加一个新的表格样式
        public static ObjectId AddTableStyle(string style)
        {
            ObjectId styleId; // 存储表格样式的Id
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 打开表格样式字典
                DBDictionary dict = (DBDictionary)db.TableStyleDictionaryId.GetObject(OpenMode.ForRead);
                if (dict.Contains(style)) // 如果存在指定的表格样式
                    styleId = dict.GetAt(style); // 获取表格样式的Id
                else
                {
                    TableStyle ts = new TableStyle(); // 新建一个表格样式
                    // 设置表格的标题行为灰色
                    //ts.SetBackgroundColor(Color.FromColorIndex(ColorMethod.ByAci, 8), (int)RowType.TitleRow);
                    // 设置表格所有行的外边框的线宽为0.30mm
                    ts.SetGridLineWeight(LineWeight.LineWeight030, (int)GridLineType.OuterGridLines, TableTools.AllRows);
                    // 不加粗表格表头行的底部边框
                    ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalBottom, (int)RowType.HeaderRow);
                    // 不加粗表格数据行的顶部边框
                    ts.SetGridLineWeight(LineWeight.LineWeight000, (int)GridLineType.HorizontalTop, (int)RowType.DataRow);
                    // 设置表格中所有行的文本高度为1
                    ts.SetTextHeight(1, TableTools.AllRows);
                    // 设置表格中所有行的对齐方式为正中
                    ts.SetAlignment(CellAlignment.MiddleCenter, TableTools.AllRows);
                    dict.UpgradeOpen();//切换表格样式字典为写的状态
                    // 将新的表格样式添加到样式字典并获取其Id
                    styleId = dict.SetAt(style, ts);
                    // 将新建的表格样式添加到事务处理中
                    trans.AddNewlyCreatedDBObject(ts, true);
                    trans.Commit();
                }
            }
            return styleId; // 返回表格样式的Id
        }
        
    }
}
