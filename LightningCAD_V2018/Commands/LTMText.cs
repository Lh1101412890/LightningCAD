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
    /// 单行文字转多行文字
    /// </summary>
    public class LTMText : CommandBase
    {
        [CommandMethod(nameof(LTMText), CommandFlags.UsePickSet)]
        public static void Command()
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
                new TypedValue((int)DxfCode.Start, "text"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result = editor.GetSelection(options, filter);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                IEnumerable<DBText> textList = document.GetObjects(result.Value.GetObjectIds()).Cast<DBText>();
                using (Transaction transaction = database.NewTransaction())
                {
                    foreach (var dBText in textList)
                    {
                        MText text = new MText
                        {
                            Location = new Point3d(dBText.Position.X, dBText.Position.Y + dBText.Height, 0),
                            Contents = dBText.TextString,
                            TextStyleId = dBText.TextStyleId,
                            TextHeight = dBText.Height,
                        };
                        document.Drawing(text, dBText.Layer);
                        document.Delete(dBText);
                    }
                    transaction.Commit();
                }
            }
            catch (System.Exception exp)
            {
                exp.Record();
            }
        }
    }
}