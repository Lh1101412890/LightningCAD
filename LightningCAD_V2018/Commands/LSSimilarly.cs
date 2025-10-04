using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    public class LSSimilar : CommandBase
    {
        [CommandMethod(nameof(LSSimilar), CommandFlags.Redraw)]
        public static void Command()
        {
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            document.SendStringToExecute("SELECTSIMILAR\n", true, false, false);
        }
    }
}