using System.Collections.Generic;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Models
{
    /// <summary>
    /// 柱模型
    /// </summary>
    public class DRColumnModel : DRComponent
    {
        public ObjectId OriginalPolylineId { get; set; }

        public DRColumnModel(Polyline polyline, DBText dBText)
        {
            OriginalPolylineId = polyline.ObjectId;
            DBTextObjectId = dBText.ObjectId;
            Name = dBText.TextString;
            IsValid = polyline.IsClosed();
        }

        private bool IsDrawing = false;
        private List<LFlash> LFlashes { get; set; } = new List<LFlash>();
        public void Drawing(string layer, ColorEnum colorEnum)
        {
            if (IsValid && !IsDrawing)
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Polyline original = document.GetObject(OriginalPolylineId) as Polyline;
                Polyline polyline = original.GetSimplest();
                polyline.ConstantWidth = 70;
                document.Drawing(polyline, layer, colorEnum);
                Point3d start = polyline.Bounds.Value.GetCenter();
                DBText dBText = document.GetObject(DBTextObjectId) as DBText;
                Point3d end = dBText.Bounds.Value.GetCenter();
                Line line = new Line(start, end);
                document.Drawing(line, layer, colorEnum);
                ComponentObjectId = polyline.ObjectId;
                LinklineObjectId = line.ObjectId;
                //创建长度提示文字瞬态
                List<Point2d> point2ds = polyline.GetPointsBasedOnCenter(5);
                for (int i = 0; i < point2ds.Count; i++)
                {
                    Point2d p1 = point2ds[i];
                    Point2d p2 = i == point2ds.Count - 1 ? point2ds[0] : point2ds[i + 1];
                    double dis = p1.GetDistanceTo(p2).Accuracy(5);
                    Point3d center = new Point3d((p1.X + p2.X) / 2 + start.X, (p1.Y + p2.Y) / 2 + start.Y, 0);
                    DBText text = new DBText()
                    {
                        TextStyleId = document.GetLTextObjectId(),
                        TextString = dis.ToString(),
                        Justify = AttachmentPoint.TopCenter,
                        AlignmentPoint = center,
                        Position = center,
                        Height = layer == LayerName.Column ? 80 : 350,//柱平面文字高度80，柱大样文字高度350
                        Layer = layer,
                    };
                    LFlash lFlash = new LFlash(text, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Highlight);
                    LFlashes.Add(lFlash);
                }
                original.Dispose();
                polyline.Dispose();
                dBText.Dispose();
                line.Dispose();
                IsDrawing = true;
            }
        }

        public void Delete()
        {
            if (IsDrawing)
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                document.Delete(ComponentObjectId);
                document.Delete(LinklineObjectId);
                foreach (var l in LFlashes)
                {
                    l.Delete();
                }
                IsDrawing = false;
            }
        }

        public void UpdateName(DBText dBText)
        {
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            using (DocumentLock @lock = document.LockDocument())
            {
                Name = dBText.TextString;
                DBTextObjectId = dBText.ObjectId;
                using (Transaction transaction = dBText.Database.NewTransaction())
                {
                    Point3d center = dBText.Bounds.Value.GetCenter();
                    Line line = LinklineObjectId.GetObject(OpenMode.ForWrite) as Line;
                    line.EndPoint = center;
                    transaction.Commit();
                }
            }
        }
    }
}