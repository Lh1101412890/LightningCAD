using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// CAD图库
    /// </summary>
    public class LLibrary : CommandBase
    {
        [CommandMethod(nameof(LLibrary))]
        public static void Command()
        {
            string file = Information.GetFileInfo("Files\\LightningCAD图库.dwg").FullName;
            var documents = CADApp.DocumentManager.Cast<Document>().ToList();
            Document document = documents.Find(d => d.Name == file);
            CADApp.DocumentManager.MdiActiveDocument = document ?? CADApp.DocumentManager.Open(file);
        }
    }
}