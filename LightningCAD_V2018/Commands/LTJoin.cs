using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 文字合并
    /// </summary>
    public class LTJoin : CommandBase
    {
        [CommandMethod(nameof(LTJoin))]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Database database = document.Database;
                Editor editor = document.Editor;

                PromptEntityOptions options1 = new PromptEntityOptions("请选择首个单行文字");
                options1.SetRejectMessage("不是单行文字\n");
                options1.AddAllowedClass(typeof(DBText), true);
                PromptEntityResult result1 = editor.GetEntity(options1);
                if (result1.Status != PromptStatus.OK)
                {
                    return;
                }
                editor.WriteMessage("\n");

                DBText dBText = document.GetObject(result1.ObjectId) as DBText;
                Polyline polyline = dBText.GetBoundsRectangle();
                Polyline marker = polyline.GetOffsetCurves(2)[0] as Polyline;
                polyline.Dispose();
                LFlash lFlash = new LFlash(marker, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Highlight);

                PromptEntityOptions options2 = new PromptEntityOptions("请选择后续单行文字");
                options2.SetRejectMessage("不是单行文字\n");
                options2.AddAllowedClass(typeof(DBText), true);
                while (true)
                {
                    PromptEntityResult result2 = editor.GetEntity(options2);
                    if (result2.Status != PromptStatus.OK)
                    {
                        break;
                    }
                    editor.WriteMessage("\n");
                    DBText t = document.GetObject(result2.ObjectId) as DBText;
                    using (Transaction tr = database.NewTransaction())
                    {
                        try
                        {
                            dBText.Id.GetObject(OpenMode.ForWrite);
                            dBText.TextString += t.TextString;
                            document.Delete(t);
                            tr.Commit();
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception exp)
                        {
                            if (exp.ErrorStatus == ErrorStatus.OnLockedLayer)
                            {
                                editor.WriteMessage("图层被锁定!\n");
                            }
                            else
                            {
                                editor.WriteMessage(exp.Message + "\n");
                                exp.Record();
                            }
                            tr.Abort();
                            break;
                        }
                        finally
                        {
                            t.Dispose();
                        }
                    }
                    using (Polyline polyline1 = dBText.GetBoundsRectangle())
                    {
                        lFlash.Update(polyline1.GetOffsetCurves(2)[0] as Polyline);
                    }
                }
                dBText.Dispose();
                lFlash.Delete();
            }
            catch (System.Exception exp)
            {
                exp.Record();
            }
        }
    }
}