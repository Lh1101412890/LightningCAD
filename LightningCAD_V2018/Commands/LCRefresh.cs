using System;
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
    /// 坐标更新
    /// </summary>
    public class LCRefresh : CommandBase
    {
        [CommandMethod(nameof(LCRefresh))]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Editor editor = document.Editor;

                SelectionFilter filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.ExtendedDataRegAppName, Information.Brand), });

                PromptSelectionResult result = editor.SelectAll(filter);
                if (result.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("没有坐标标注\n");
                    return;
                }

                var dBObjects = document.GetObjects(result.Value.GetObjectIds()).OfType<MLeader>();

                int n = 0;
                int l = 0;
                using (Transaction transaction = document.Database.NewTransaction())
                {
                    foreach (var mLeader in dBObjects)
                    {
                        Point3d point3d = mLeader.GetFirstVertex(0);
                        double x, y;
                        ResultBuffer resultBuffer = document.GetXData(mLeader);
                        if (resultBuffer == null) { continue; }
                        TypedValue[] values = resultBuffer.AsArray();
                        var data = values[1].Value.ToString();
                        switch (data)
                        {
                            case "MCoordinate":
                                x = Math.Round(point3d.Y, 3);
                                y = Math.Round(point3d.X, 3);
                                break;
                            case "MmCoordinate":
                                x = Math.Round(point3d.Y / 1000, 3);
                                y = Math.Round(point3d.X / 1000, 3);
                                break;
                            default:
                                continue;
                        }

                        string contents = mLeader.MText.Contents;
                        string first = contents.Split("\\P".ToCharArray()).First();
                        string last = contents.Split("\\P".ToCharArray()).Last();
                        if (!first.StartsWith("X=") || !last.StartsWith("Y="))
                        {
                            continue;
                        }
                        string coordinate = "X=" + x + "\\PY=" + y;
                        if (contents != coordinate)
                        {
                            LayerTableRecord layerTableRecord = document.GetObject(mLeader.LayerId) as LayerTableRecord;
                            if (layerTableRecord.IsLocked)
                            {
                                l++;
                                continue;
                            }
                            mLeader.Id.GetObject(OpenMode.ForWrite);
                            mLeader.MText = new MText()
                            {
                                Contents = coordinate,
                                Rotation = 0,//多行文字旋转角度0，使文字对齐ucs的x轴
                            };
                            n++;
                        }
                    }
                    transaction.Commit();
                }
                editor.WriteMessage($"坐标标注共更新 {n} 个，有 {l} 个在锁定图层未更新\n");
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
    }
}