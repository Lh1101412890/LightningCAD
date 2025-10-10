using System;
using System.Collections.Generic;
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
    /// <summary>
    /// （点）高程标注
    /// </summary>
    public class LElevation : CommandBase
    {
        [CommandMethod(nameof(LElevation), CommandFlags.UsePickSet)]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Database database = document.Database;
                Editor editor = document.Editor;

                PromptSelectionOptions prompt1 = new PromptSelectionOptions() { MessageForAdding = "请选择需标注的点" };
                TypedValue[] typedValues =
                {
                    new TypedValue((int)DxfCode.Start,"point"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result1 = editor.GetSelection(prompt1, filter);
                if (result1.Status != PromptStatus.OK)
                {
                    return;
                }

                LSetting.CreatStyles(document);
                ObjectId styleId;
                double height = 3.5;
                using (Transaction transaction = database.NewTransaction())
                {
                    TextStyleTable textStyleTable = (TextStyleTable)transaction.GetObject(database.TextStyleTableId, OpenMode.ForRead);
                    ObjectContextManager manager = database.ObjectContextManager;
                    //获得当前图形的注释比例列表，名为“ACDB_ANNOTATIONSCALES”
                    ObjectContextCollection contexts = manager.GetContextCollection("ACDB_ANNOTATIONSCALES");
                    string current = contexts.CurrentContext.Name.Replace(':', '：');//当前注释比例名称
                    if (textStyleTable.Has("LText-" + current))
                    {
                        styleId = textStyleTable["LText-" + current];//通过比例选取样式
                        height *= double.Parse(current.Split('：').Last()) / double.Parse(current.Split('：').First());
                    }
                    else
                    {
                        styleId = textStyleTable["LText-1：100"];
                        height *= 100;
                    }
                    List<DBPoint> points = document.GetObjects(result1.Value.GetObjectIds()).Cast<DBPoint>().ToList();
                    foreach (var point in points)
                    {
                        DBText dBText = new DBText()
                        {
                            TextStyleId = styleId,
                            TextString = "Z=" + Math.Round(point.Position.Z, 3).ToString() + "m",
                            Position = Point3d.Origin,
                            HorizontalMode = TextHorizontalMode.TextLeft,
                            VerticalMode = TextVerticalMode.TextTop,
                            Height = height,
                            AlignmentPoint = new Point3d(point.Position.X + height * 0.7 / 2, point.Position.Y - height / 4, 0)
                        };
                        document.Drawing(dBText, LayerName.PointElevation, ColorEnum.Magenta);
                        Guid guid = Guid.NewGuid();
                        ResultBuffer values = new ResultBuffer()
                        {
                            new TypedValue(1001, Information.Brand),
                            new TypedValue((int)DxfCode.ExtendedDataAsciiString, $"PointElevation:{guid}"),
                        };
                        document.SetXData(point, values);
                        document.SetXData(dBText, values);
                        dBText.Dispose();
                    }
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