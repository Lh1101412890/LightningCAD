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
    /// 墙模型
    /// </summary>
    public class DRWallModel : DRComponent
    {
        public List<ObjectId> OriginalLineIds { get; set; }
        public double Width { get; set; }
        public Point2d StartPoint { get; set; }
        public Point2d EndPoint { get; set; }
        public DRWallModel(Line line)
        {
            IsValid = false;
            OriginalLineIds = new List<ObjectId>
            {
                line.ObjectId,
            };
        }
        public DRWallModel(Line line1, Line line2)
        {
            IsValid = true;
            OriginalLineIds = new List<ObjectId>
            {
                line1.Id,
                line2.Id
            };
            Width = line1.GetDistance(line2);

            line1.GetCenterLine(line2, out Line line);
            StartPoint = line.StartPoint.ToPoint2d();
            EndPoint = line.EndPoint.ToPoint2d();
        }
        private LFlash LFlash { get; set; }

        private bool IsDrawingComponent = false;
        private bool IsDrawingLinkline = false;

        public void SetName(DBText text)
        {
            Name = text.TextString;
            DBTextObjectId = text.ObjectId;
        }
        public void SetName()
        {
            Name = "Q";
            DBTextObjectId = ObjectId.Null;
        }
        public void DrawingComponent()
        {
            if (IsValid && !IsDrawingComponent)
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Polyline polyline = new Polyline();
                polyline.AddVertexAt(0, StartPoint, 0, 0, 0);
                polyline.AddVertexAt(1, EndPoint, 0, 0, 0);
                polyline.ConstantWidth = 50;
                document.Drawing(polyline, LayerName.Wall, ColorEnum.Yellow);
                ComponentObjectId = polyline.ObjectId;
                //长度提示文字瞬态
                DBText dBText = new DBText()
                {
                    TextStyleId = document.GetLTextObjectId(),
                    TextString = Width.Accuracy(0).ToString(),
                    Justify = AttachmentPoint.MiddleCenter,
                    Position = polyline.Bounds.Value.GetCenter(),
                    Height = Width.Accuracy(0),
                    Layer = LayerName.Wall,
                };
                polyline.Dispose();
                LFlash = new LFlash(dBText, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Highlight);
                IsDrawingComponent = true;
            }
        }
        public void DrawingLinkline()
        {
            if (IsValid && !IsDrawingLinkline && DBTextObjectId != ObjectId.Null)
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Polyline polyline = document.GetObject(ComponentObjectId) as Polyline;
                Point3d start = polyline.Bounds.Value.GetCenter();
                DBText dBText = document.GetObject(DBTextObjectId) as DBText;
                Point3d end = dBText.Bounds.Value.GetCenter();
                Line line = new Line(start, end);
                document.Drawing(line, LayerName.Wall, ColorEnum.Yellow);
                LinklineObjectId = line.ObjectId;
                polyline.Dispose();
                dBText.Dispose();
                line.Dispose();
                IsDrawingLinkline = true;
            }
        }
        public void Delete()
        {
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            if (IsDrawingComponent)
            {
                document.Delete(ComponentObjectId);
                LFlash.Delete();
                IsDrawingComponent = false;
            }
            if (IsDrawingLinkline)
            {
                document.Delete(LinklineObjectId);
                IsDrawingLinkline = false;
            }
        }
        public void UpdateName(DBText dBText)
        {
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            switch (dBText)
            {
                case null:
                    Name = "Q";
                    if (LinklineObjectId != ObjectId.Null)
                    {
                        document.Delete(LinklineObjectId);
                        LinklineObjectId = ObjectId.Null;
                    }
                    break;
                default:
                    using (DocumentLock @lock = document.LockDocument())
                    {
                        using (Transaction transaction = dBText.Database.NewTransaction())
                        {
                            Name = dBText.TextString;
                            Point3d center = dBText.Bounds.Value.GetCenter();
                            if (LinklineObjectId != ObjectId.Null)
                            {
                                Line line = LinklineObjectId.GetObject(OpenMode.ForWrite, false, false) as Line;
                                line.EndPoint = center;
                                line.Dispose();
                            }
                            else
                            {
                                Polyline polyline = document.GetObject(ComponentObjectId) as Polyline;
                                Point3d start = polyline.Bounds.Value.GetCenter();
                                Point3d end = dBText.Bounds.Value.GetCenter();
                                Line line = new Line(start, end);
                                document.Drawing(line, LayerName.Wall, ColorEnum.Yellow);
                                LinklineObjectId = line.ObjectId;
                                polyline.Dispose();
                                line.Dispose();
                            }
                            transaction.Commit();
                        }
                    }
                    break;
            }
        }
    }
}