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
    /// 直线转多段线
    /// </summary>
    public class LLTPolyline : CommandBase
    {
        [CommandMethod(nameof(LLTPolyline), CommandFlags.UsePickSet)]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Editor editor = document.Editor;
                PromptSelectionOptions prompt = new PromptSelectionOptions() { MessageForAdding = "请选择直线" };
                TypedValue[] typedValues =
                {
                    new TypedValue((int)DxfCode.Start,"line"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result = editor.GetSelection(prompt, filter);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }

                List<Line> lines = document.GetObjects(result.Value.GetObjectIds()).Cast<Line>().ToList();

                foreach (var line in lines)
                {
                    Polyline polyline = new Polyline()
                    {
                        Closed = false,
                        Color = line.Color
                    };
                    polyline.AddVertexAt(0, line.StartPoint.ToPoint2d(), 0, 0, 0);
                    polyline.AddVertexAt(1, line.EndPoint.ToPoint2d(), 0, 0, 0);
                    document.Drawing(polyline, line.Layer);

                    document.Delete(line);
                }
                editor.WriteMessage($"共转换 {lines.Count} 个\n");
            }
            catch (System.Exception exp)
            {
                exp.Record();
            }
        }
    }
}