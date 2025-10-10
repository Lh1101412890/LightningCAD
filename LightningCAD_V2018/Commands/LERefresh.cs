using System;
using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 高程更新
    /// </summary>
    public class LERefresh : CommandBase
    {
        [CommandMethod(nameof(LERefresh))]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Database database = document.Database;
                Editor editor = document.Editor;

                TypedValue[] typedValues1 =
                {
                    new TypedValue((int)DxfCode.Operator,"<and"),
                    new TypedValue((int)DxfCode.Start,"Text"),
                    new TypedValue((int)DxfCode.Operator,"and>"),
                };
                SelectionFilter filter1 = new SelectionFilter(typedValues1);
                PromptSelectionResult result1 = editor.SelectAll(filter1);

                SelectionFilter filter2 = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "point") });
                var result2 = editor.SelectAll(filter2);

                if (result1.Status != PromptStatus.OK || result2.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("没有高程标注\n");
                    return;
                }
                List<DBText> dBTexts = document.GetObjects(result1.Value.GetObjectIds()).Cast<DBText>().ToList();
                List<DBPoint> dBPoints = document.GetObjects(result2.Value.GetObjectIds()).Cast<DBPoint>().ToList();

                int n = 0;
                int e = 0;
                int l = 0;
                using (Transaction transaction = database.NewTransaction())
                {
                    foreach (var text in dBTexts)
                    {
                        LayerTableRecord layerTableRecord = document.GetObject(text.LayerId) as LayerTableRecord;
                        if (layerTableRecord.IsLocked)
                        {
                            l++;
                            continue;
                        }
                        ResultBuffer resultBuffer = document.GetXData(text);
                        if (resultBuffer == null) { continue; }
                        TypedValue[] values1 = resultBuffer.AsArray();
                        string data1 = values1[1].Value.ToString();
                        if (data1.StartsWith("PointElevation:"))
                        {
                            string guid1 = data1.Replace("PointElevation:", "");
                            if (dBPoints.Exists(p =>
                            {
                                TypedValue[] values2 = document.GetXData(p).AsArray();
                                string data2 = values2[1].Value.ToString();
                                string guid2 = data2.Replace("PointElevation:", "");
                                return guid1 == guid2;
                            }))
                            {
                                DBPoint dBPoint = dBPoints.First(p =>
                                {
                                    TypedValue[] values2 = document.GetXData(p).AsArray();
                                    string data2 = values2[1].Value.ToString();
                                    string guid2 = data2.Replace("PointElevation:", "");
                                    return guid1 == guid2;
                                });
                                string str = "Z=" + Math.Round(dBPoint.Position.Z, 3).ToString() + "m";
                                if (text.TextString != str)
                                {
                                    text.Id.GetObject(OpenMode.ForWrite);
                                    text.TextString = str;
                                    n++;
                                }
                            }
                            else
                            {
                                e++;
                            }
                        }
                    }
                    transaction.Commit();
                }
                editor.WriteMessage($"高程标注共更新 {n} 个，有 {l} 个在锁定图层未更新，有 {e} 个标注点已失效\n");
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
    }
}