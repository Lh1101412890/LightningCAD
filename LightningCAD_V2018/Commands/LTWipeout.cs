using System.Collections.Generic;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 文字背景遮罩
    /// </summary>
    public class LTWipeout : CommandBase
    {
        [CommandMethod(nameof(LTWipeout))]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Database database = document.Database;
                Editor editor = document.Editor;
                PromptSelectionOptions options = new PromptSelectionOptions()
                {
                    MessageForAdding = "请选择需文字"
                };
                TypedValue[] typedValues = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<or"),
                    new TypedValue((int)DxfCode.Start, "text"),
                    new TypedValue((int)DxfCode.Start, "mtext"),
                    new TypedValue((int)DxfCode.Operator, "or>"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result = editor.GetSelection(options, filter);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                List<DBObject> objects = document.GetObjects(result.Value.GetObjectIds());
                using (Transaction transaction = database.NewTransaction())
                {
                    foreach (var entity in objects)
                    {
                        Point2d min = entity.Bounds.Value.MinPoint.ToPoint2d();
                        Point2d max = entity.Bounds.Value.MaxPoint.ToPoint2d();
                        Point2dCollection pt2dArray = new Point2dCollection
                        {
                            min,
                            new Point2d(max.X, min.Y),
                            max,
                            new Point2d(min.X, max.Y),
                            min//必须闭合
                        };
                        Wipeout wipeout = new Wipeout();
                        wipeout.SetFrom(pt2dArray, new Vector3d(0, 0, 1));
                        document.Drawing(wipeout, nameof(LTWipeout), ColorEnum.White);
                    }
                    BlockTable bt = transaction.GetObject(database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    //绘制排序表，命令DR，上下层关系
                    DrawOrderTable orderTable = transaction.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;
                    orderTable.MoveToTop(new ObjectIdCollection(result.Value.GetObjectIds()));
                    transaction.Commit();
                }
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
    }
}