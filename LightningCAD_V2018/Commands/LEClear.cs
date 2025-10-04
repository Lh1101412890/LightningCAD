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
    /// 高程清理
    /// </summary>
    public class LEClear : CommandBase
    {
        [CommandMethod(nameof(LEClear))]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Editor editor = document.Editor;
                TypedValue[] typedValues1 =
                {
                    new TypedValue((int)DxfCode.Operator,"<and"),
                    new TypedValue((int)DxfCode.Start,"Text"),
                    new TypedValue((int)DxfCode.Operator,"and>"),
                };
                SelectionFilter filter1 = new SelectionFilter(typedValues1);
                var result1 = editor.SelectAll(filter1);

                SelectionFilter filter2 = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "point"), });
                var result2 = editor.SelectAll(filter2);
                if (result1.Status != PromptStatus.OK)
                {
                    editor.WriteMessage("没有高程标注\n");
                    return;
                }
                if (result2.Status != PromptStatus.OK)
                {
                    return;
                }
                List<DBText> texts = document.GetObjects(result1.Value.GetObjectIds()).Cast<DBText>().ToList();
                List<DBPoint> dBPoints = document.GetObjects(result2.Value.GetObjectIds()).Cast<DBPoint>().ToList();
                int e = 0;
                int l = 0;
                foreach (var dBText in texts)
                {
                    ResultBuffer resultBuffer = document.GetXData(dBText);
                    if (resultBuffer == null) { continue; }
                    var guid1 = resultBuffer.AsArray()[1].Value.ToString().Replace("PointElevation:", "");
                    if (!dBPoints.Exists(p =>
                    {
                        TypedValue[] values2 = document.GetXData(p).AsArray();
                        string data2 = values2[1].Value.ToString();
                        string guid2 = data2.Replace("PointElevation:", "");
                        return guid1 == guid2;
                    }))
                    {
                        LayerTableRecord layerTableRecord = document.GetObject(dBText.LayerId) as LayerTableRecord;
                        if (layerTableRecord.IsLocked)
                        {
                            l++;
                            continue;
                        }
                        document.Delete(dBText);
                        e++;
                    }
                }
                editor.WriteMessage($"高程标注共删除 {e} 个，有 {l} 个在锁定图层未删除\n");
            }
            catch (System.Exception exp)
            {
                exp.Record();
            }
        }
    }
}