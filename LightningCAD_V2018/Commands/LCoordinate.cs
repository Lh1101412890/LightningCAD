using System;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    public class LCoordinate : CommandBase
    {
        public static void Coordinate(string s)
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Database database = document.Database;
                Editor editor = document.Editor;

                PromptPointResult promptPointResult = editor.GetPoint("请选择标注点");
                if (promptPointResult.Status != PromptStatus.OK)
                {
                    return;
                }

                Point3d start = promptPointResult.Value;

                CoordinateSystem3d ucs = editor.CurrentUserCoordinateSystem.CoordinateSystem3d;
                // Transform from UCS to WCS
                Matrix3d mat = Matrix3d.AlignCoordinateSystem(
                    Point3d.Origin,
                    Vector3d.XAxis,
                    Vector3d.YAxis,
                    Vector3d.ZAxis,
                    ucs.Origin,
                    ucs.Xaxis,
                    ucs.Yaxis,
                    ucs.Zaxis);

                Point3d temp = start.TransformBy(mat);
                if (!start.Equals(temp))
                {
                    start = temp;
                }

                DragLine class1 = new DragLine(start, "指定标注位置");

                PromptResult promptResult = editor.Drag(class1);
                if (promptResult.Status != PromptStatus.OK)
                {
                    return;
                }

                Point3d end = class1.end;

                double x;
                double y;
                if (s == "M")
                {
                    x = Math.Round(start.Y, 3);
                    y = Math.Round(start.X, 3);
                }
                else
                {
                    x = Math.Round(start.Y / 1000, 3);
                    y = Math.Round(start.X / 1000, 3);
                }

                LSetting.CreatStyles(document);
                ObjectId styleId;
                using (Transaction transaction = database.NewTransaction())
                {
                    DBDictionary dBDictionary = transaction.GetObject(database.MLeaderStyleDictionaryId, OpenMode.ForRead) as DBDictionary;
                    ObjectContextManager manager = database.ObjectContextManager;
                    //获得当前图形的注释比例列表，名为“ACDB_ANNOTATIONSCALES”
                    ObjectContextCollection contexts = manager.GetContextCollection("ACDB_ANNOTATIONSCALES");
                    string current = contexts.CurrentContext.Name.Replace(':', '：');//当前注释比例名称
                    try
                    {
                        styleId = (ObjectId)dBDictionary["LMleader-" + current];//通过比例选取样式
                    }
                    catch (System.Exception)
                    {
                        styleId = (ObjectId)dBDictionary["LMleader-1：100"];
                    }
                    transaction.Commit();
                }

                MLeader mLeader = new MLeader()
                {
                    MLeaderStyle = styleId,
                };

                mLeader.AddLeader();
                mLeader.AddLeaderLine(0);
                mLeader.AddFirstVertex(0, start);
                mLeader.AddLastVertex(0, end);

                mLeader.MText = new MText()
                {
                    Contents = "X=" + x + "\\PY=" + y,
                    Rotation = 0,//多行文字旋转角度0，使文字对齐ucs的x轴
                };
                document.Drawing(mLeader, LayerName.Coordinate, ColorEnum.Green);
                string type = s == "M" ? "MCoordinate" : "MmCoordinate";
                ResultBuffer values = new ResultBuffer()
                {
                    new TypedValue(1001, Information.Brand),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, type),
                };
                document.SetXData(mLeader, values);
                values.Dispose();
                //重新设定引线位置，更新显示
                using (Transaction transaction = database.NewTransaction())
                {
                    MLeader mLeader1 = mLeader.Id.GetObject(OpenMode.ForWrite) as MLeader;
                    if (end.X < start.X)
                    {
                        mLeader1.SetDogleg(0, -Vector3d.XAxis);
                        mLeader1.SetVertex(0, 1, end);
                    }
                    mLeader1.TextHeight = 3.5;
                    transaction.Commit();
                }
            }
            catch (System.Exception exp)
            {
                exp.Record();
            }
        }
    }

    /// <summary>
    /// 坐标标注M
    /// </summary>
    public class LMCoordinate : CommandBase
    {
        [CommandMethod(nameof(LMCoordinate))]
        public static void Command()
        {
            LCoordinate.Coordinate("M");
        }
    }

    /// <summary>
    /// 坐标标注Mm
    /// </summary>
    public class LMmCoordinate : CommandBase
    {
        [CommandMethod(nameof(LMmCoordinate))]
        public static void Command()
        {
            LCoordinate.Coordinate("Mm");
        }
    }
}