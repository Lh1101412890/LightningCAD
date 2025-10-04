using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 关闭所有文件
    /// </summary>
    public class LCFiles : CommandBase
    {
        [CommandMethod(nameof(LCFiles), CommandFlags.Session)]//必须设置成CommandFlags.Session，否则文件不能保存，会提示图形忙
        public static void Command()
        {
            Editor editor = CADApp.DocumentManager.MdiActiveDocument.Editor;
            PromptKeywordOptions prompt = new PromptKeywordOptions("是否保存所有文件和关闭CAD？[保存所有文件：/---关闭CAD(YY)/---不关闭CAD(YN)/不保存：/---关闭CAD(NY)/---不关闭CAD(NN)/放弃(U)]")
            {
                AppendKeywordsToMessage = true,
                AllowNone = false
            };
            prompt.Keywords.Add("YY");
            prompt.Keywords.Add("YN");
            prompt.Keywords.Add("NY");
            prompt.Keywords.Add("NN");
            prompt.Keywords.Add("U");
            PromptResult result = editor.GetKeywords(prompt);
            if (result.Status != PromptStatus.OK || result.StringResult == "U") return;

            //根据选择关闭文件
            switch (result.StringResult)
            {
                case "YY":
                    foreach (Document item in CADApp.DocumentManager)
                    {
                        item.CloseAndSave(item.Name);
                    }
                    CADApp.MainWindow.Close();
                    break;
                case "YN":
                    foreach (Document item in CADApp.DocumentManager)
                    {
                        item.CloseAndSave(item.Name);
                    }
                    break;
                case "NY":
                    foreach (Document item in CADApp.DocumentManager)
                    {
                        item.CloseAndDiscard();
                    }
                    CADApp.MainWindow.Close();
                    break;
                case "NN":
                    foreach (Document item in CADApp.DocumentManager)
                    {
                        item.CloseAndDiscard();
                    }
                    break;
                default:
                    break;
            }
        }
    }
}