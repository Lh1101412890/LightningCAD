using System.Diagnostics;

using Autodesk.AutoCAD.Runtime;

using LightningCAD.LightningExtension;

namespace LightningCAD.Commands
{
    public class LUNotes : CommandBase
    {
        [CommandMethod(nameof(LUNotes))]
        public static void Command()
        {
            Process.Start("Explorer", Information.GetFileInfo("更新说明.pdf").FullName);
        }
    }
}