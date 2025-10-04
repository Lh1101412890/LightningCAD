using System;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using LightningCAD.Extension;

namespace LightningCAD.Models.DRBeamModels
{
    public abstract class ComponentBase
    {
        public Line Line { get; set; }
        /// <summary>
        /// 图纸中两条直线形成的构件的实际宽度
        /// </summary>
        public double Width { get; set; }

        public Polyline GetPolyline()
        {
            Polyline polyline = new Polyline();
            double halfwidth = Math.Round(Width / 2);
            Point2d StartPoint = Line.StartPoint.ToPoint2d();
            Point2d EndPoint = Line.EndPoint.ToPoint2d();
            if (Line.IsVertical())
            {
                polyline.AddVertexAt(0, new Point2d(StartPoint.X - halfwidth, StartPoint.Y), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(StartPoint.X + halfwidth, StartPoint.Y), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(EndPoint.X + halfwidth, EndPoint.Y), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(EndPoint.X - halfwidth, EndPoint.Y), 0, 0, 0);
            }
            else if (Line.IsHorizontal())
            {
                polyline.AddVertexAt(0, new Point2d(StartPoint.X, StartPoint.Y - halfwidth), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(StartPoint.X, StartPoint.Y + halfwidth), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(EndPoint.X, EndPoint.Y + halfwidth), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(EndPoint.X, EndPoint.Y - halfwidth), 0, 0, 0);
            }
            else
            {
                double k = StartPoint.GetK(EndPoint);
                double x = halfwidth * Math.Sin(Math.Atan(k));
                double b1 = StartPoint.Y + 1 / k * StartPoint.X;
                double b2 = EndPoint.Y + 1 / k * EndPoint.X;
                double y1 = b1 - 1 / k * (StartPoint.X - x);
                double y2 = b1 - 1 / k * (StartPoint.X + x);
                double y3 = b2 - 1 / k * (EndPoint.X + x);
                double y4 = b2 - 1 / k * (EndPoint.X - x);
                polyline.AddVertexAt(0, new Point2d(StartPoint.X - x, y1), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(StartPoint.X + x, y2), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(EndPoint.X + x, y3), 0, 0, 0);
                polyline.AddVertexAt(0, new Point2d(EndPoint.X - x, y4), 0, 0, 0);
            }
            polyline.Closed = true;
            return polyline;
        }

        public abstract ComponentBase Creat(Line line1, Line line2);
    }
}