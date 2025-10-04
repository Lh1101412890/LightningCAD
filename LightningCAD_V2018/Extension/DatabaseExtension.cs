using Autodesk.AutoCAD.DatabaseServices;

namespace LightningCAD.Extension
{
    public static class DatabaseExtension
    {
        public static Transaction NewTransaction(this Database database)
        {
            return database.TransactionManager.StartTransaction();
        }
    }
}