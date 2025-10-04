using System;
using System.Collections.Generic;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

namespace LightningCAD.Models.DRBeamModels
{
    public class ComponentBeam : ComponentBase
    {
        public bool HasSupportAtStartPoint(List<Polyline> polylines, double acc = LightningTolerance.Structural)
        {
            return polylines.Exists(p => Line.StartPoint.IsInsideByBounds(p, acc));
        }

        public bool HasSupportAtEndPoint(List<Polyline> polylines, double acc = LightningTolerance.Structural)
        {
            return polylines.Exists(p => Line.EndPoint.IsInsideByBounds(p, acc));
        }

        public ComponentBeam()
        {

        }
        public int Height { get; set; }
        private ComponentBeam(Line line1, Line line2)
        {
            Width = Math.Round(line1.GetDistance(line2));
            Line line = null;

            //斜率k为无穷
            if (line1.IsVertical())
            {
                if (line2.IsVertical())
                {
                    double l1min = Math.Min(line1.StartPoint.Y, line1.EndPoint.Y);
                    double l1max = Math.Max(line1.StartPoint.Y, line1.EndPoint.Y);
                    double l2min = Math.Min(line2.StartPoint.Y, line2.EndPoint.Y);
                    double l2max = Math.Max(line2.StartPoint.Y, line2.EndPoint.Y);
                    double min = (l1min + l2min) / 2;
                    double max = (l1max + l2max) / 2;
                    double x = (line1.StartPoint.X + line1.EndPoint.X + line2.StartPoint.X + line2.EndPoint.X) / 4;
                    line = new Line(new Point3d(x, min, 0), new Point3d(x, max, 0));
                }
            }
            //斜率k为0
            else if (line1.IsHorizontal())
            {
                if (line2.IsHorizontal())
                {
                    double l1min = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double l1max = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double l2min = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double l2max = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    double min = (l1min + l2min) / 2;
                    double max = (l1max + l2max) / 2;
                    double y = (line1.StartPoint.Y + line1.EndPoint.Y + line2.StartPoint.Y + line2.EndPoint.Y) / 4;
                    line = new Line(new Point3d(min, y, 0), new Point3d(max, y, 0));
                }
            }
            else
            {
                double k1 = line1.GetK();
                double k2 = line2.GetK();
                if (Math.Abs(k1 - k2) < 0.2)
                {
                    double b1 = line1.GetB();
                    double b2 = line2.GetB();

                    double k3 = -1 / k1;
                    double b_b = (b1 + b2) / 2;
                    double b_1 = line1.StartPoint.Y - k3 * line1.StartPoint.X;
                    double b_2 = line1.EndPoint.Y - k3 * line1.EndPoint.X;
                    double b_3 = line2.StartPoint.Y - k3 * line2.StartPoint.X;
                    double b_4 = line2.EndPoint.Y - k3 * line2.EndPoint.X;
                    double x1 = (b_1 - b_b) / (k1 - k3);
                    double x2 = (b_2 - b_b) / (k1 - k3);
                    double x3 = (b_3 - b_b) / (k1 - k3);
                    double x4 = (b_4 - b_b) / (k1 - k3);
                    double l1min = Math.Min(x1, x2);
                    double l1max = Math.Max(x1, x2);
                    double l2min = Math.Min(x3, x4);
                    double l2max = Math.Max(x3, x4);
                    double min = (l1min + l2min) / 2;
                    double max = (l1max + l2max) / 2;
                    line = new Line(new Point3d(min, k1 * min + b_b, 0), new Point3d(max, k1 * max + b_b, 0));
                }
            }
            Line = line;
            Height = 50;
        }

        /// <summary>
        /// 设置梁构件的原位标注
        /// </summary>
        /// <param name="str">需要是尺寸或标高原位标注，格式：“200x400”</param>
        public void SetOrthotopic(string str)
        {
            if (str.Contains("x"))
            {
                if (int.TryParse(str.Split('x')[0], out int width))
                {
                    Width = width;
                }
                if (int.TryParse(str.Split('x')[1], out int height))
                {
                    Height = height;
                }
            }
        }

        public override ComponentBase Creat(Line line1, Line line2)
        {
            return new ComponentBeam(line1, line2);
        }
    }
}