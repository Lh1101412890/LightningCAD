using System.Collections.Generic;

using Autodesk.AutoCAD.DatabaseServices;

namespace LightningCAD.Extension
{
    public static class DBOBjectExtension
    {
        public static void DisposeAll(this IEnumerable<DBObject> objects)
        {
            foreach (var item in objects)
            {
                item.Dispose();
            }
        }
    }
}