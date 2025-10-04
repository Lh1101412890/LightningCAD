using System.Linq;

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
    public class LTAlign : CommandBase
    {
        public static void TextAlign(string str)
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Database database = document.Database;
                Editor editor = document.Editor;
                PromptSelectionOptions options = new PromptSelectionOptions()
                {
                    MessageForAdding = "请选择文字"
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
                PromptPointResult promptPointResult = editor.GetPoint("指定对齐位置");
                if (promptPointResult.Status != PromptStatus.OK)
                {
                    return;
                }
                Point3d point = promptPointResult.Value;
                var objects = document.GetObjects(result.Value.GetObjectIds()).Cast<Entity>();
                double x = point.X;
                double y = point.Y;
                int n = 0;
                using (Transaction transaction = database.NewTransaction())
                {
                    foreach (var item in objects)
                    {
                        Point3d max = item.GeometricExtents.MaxPoint;
                        Point3d min = item.GeometricExtents.MinPoint;
                        double mx = 0;
                        double my = 0;
                        switch (str)
                        {
                            case "L":
                                mx = x - min.X;
                                break;
                            case "R":
                                mx = x - max.X;
                                break;
                            case "T":
                                my = y - max.Y;
                                break;
                            case "B":
                                my = y - min.Y;
                                break;
                            //水平
                            case "H":
                                mx = x - (min.X + max.X) / 2;
                                break;
                            //垂直
                            case "V":
                                my = y - (min.Y + max.Y) / 2;
                                break;
                            default:
                                break;
                        }
                        Vector3d vector = new Vector3d(mx, my, 0);
                        if (item is MText mText)
                        {
                            mText.Id.GetObject(OpenMode.ForWrite);
                            mText.Location += vector;
                        }
                        if (item is DBText dBText)
                        {
                            if (dBText.Justify != AttachmentPoint.BaseAlign && dBText.Justify != AttachmentPoint.BaseFit)
                            {
                                dBText.Id.GetObject(OpenMode.ForWrite);
                                switch (dBText.Justify)
                                {
                                    case AttachmentPoint.BaseLeft:
                                        dBText.Position += vector;
                                        break;
                                    default:
                                        dBText.AlignmentPoint += vector;
                                        break;
                                }
                            }
                            else
                            {
                                n++;
                            }
                        }
                        item.Dispose();
                    }
                    transaction.Commit();
                }
                if (n > 0)
                    editor.WriteMessage($"有 {n} 个单行文字对齐属性为对齐或布满，不适用对齐功能\n");
            }
            catch (System.Exception exp)
            {
                exp.Record();
            }
        }
    }

    /// <summary>
    /// 居左对齐
    /// </summary>
    public class LTLeft : CommandBase
    {
        [CommandMethod(nameof(LTLeft), CommandFlags.Redraw)]
        public static void Command()
        {
            LTAlign.TextAlign("L");
        }
    }

    /// <summary>
    /// 居右对齐
    /// </summary>
    public class LTRight : CommandBase
    {
        [CommandMethod(nameof(LTRight), CommandFlags.Redraw)]
        public static void Command()
        {
            LTAlign.TextAlign("R");
        }
    }

    /// <summary>
    /// 居上对齐
    /// </summary>
    public class LTTop : CommandBase
    {
        [CommandMethod(nameof(LTTop), CommandFlags.Redraw)]
        public static void Command()
        {
            LTAlign.TextAlign("T");
        }
    }

    /// <summary>
    /// 居下对齐
    /// </summary>
    public class LTBottomt : CommandBase
    {
        [CommandMethod(nameof(LTBottomt), CommandFlags.Redraw)]
        public static void Command()
        {
            LTAlign.TextAlign("B");
        }
    }

    /// <summary>
    /// 水平居右
    /// </summary>
    public class LTHorizontal : CommandBase
    {
        [CommandMethod(nameof(LTHorizontal), CommandFlags.Redraw)]
        public static void Command()
        {
            LTAlign.TextAlign("H");
        }
    }

    /// <summary>
    /// 垂直居右
    /// </summary>
    public class LTVertical : CommandBase
    {
        [CommandMethod(nameof(LTVertical), CommandFlags.Redraw)]
        public static void Command()
        {
            LTAlign.TextAlign("V");
        }
    }
}