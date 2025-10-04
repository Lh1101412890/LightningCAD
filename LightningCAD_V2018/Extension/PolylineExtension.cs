using System;
using System.Collections.Generic;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using Lightning.Extension;

using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static class PolylineExtension
    {
        /// <summary>
        /// 去掉polyline多余的顶点求得最简化的封闭polyline
        /// </summary>
        /// <param name="polyline"></param>
        /// <returns></returns>
        public static Polyline GetSimplest(this Polyline polyline)
        {
            List<Point2d> points = new List<Point2d>();
            //提取点
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                points.Add(polyline.GetPoint2dAt(i));
            }
            //去掉重复的点
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = i + 1; j < points.Count; j++)
                {
                    if (points[i].GetDistanceTo(points[j]) < 5)
                    {
                        points.RemoveAt(j);
                    }
                }
            }
            //去掉共线的点
            for (int i = 0; i < points.Count;)
            {
                if (points.Count < 4)
                {
                    return null;
                }
                if (i == 0)
                {
                    double k1 = points[points.Count - 1].GetK(points[0]);
                    double k2 = points[0].GetK(points[1]);
                    if (k1 == k2)
                    {
                        points.RemoveAt(0);
                    }
                    else
                    {
                        i++;
                    }
                }
                else if (i == points.Count - 1)
                {
                    double k1 = points[points.Count - 2].GetK(points[points.Count - 1]);
                    double k2 = points[points.Count - 1].GetK(points[0]);
                    if (k1 == k2)
                    {
                        points.RemoveAt(points.Count - 1);
                    }
                    else
                    {
                        i++;
                    }
                }
                else
                {
                    double k1 = points[i - 1].GetK(points[i]);
                    double k2 = points[i].GetK(points[i + 1]);
                    if (k1 == k2)
                    {
                        points.RemoveAt(i);
                    }
                    else
                    {
                        i++;
                    }
                }
            }
            Polyline polyline1 = new Polyline()
            {
                Closed = true
            };
            for (int i = 0; i < points.Count; i++)
            {
                polyline1.AddVertexAt(i, points[i], 0, 0, 0);
            }
            return polyline1;
        }

        /// <summary>
        /// 获取多段线的顶点（基于多段线中心坐标）
        /// </summary>
        /// <param name="polyline"></param>
        /// <param name="acc">数据精度</param>
        /// <returns></returns>
        public static List<Point2d> GetPointsBasedOnCenter(this Polyline polyline, int acc = LightningTolerance.Precision)
        {
            List<Point2d> points = new List<Point2d>();
            Point3d center = polyline.GetCenter();
            double x0 = center.X;
            double y0 = center.Y;
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                Point2d point2d = polyline.GetPoint2dAt(i);
                double x1 = point2d.X;
                double y1 = point2d.Y;
                double x = (x1 - x0).Accuracy(acc);
                double y = (y1 - y0).Accuracy(acc);
                points.Add(new Point2d(x, y));
            }
            return points;
        }

        /// <summary>
        /// 比较两个多段线是否一致
        /// </summary>
        /// <param name="polyline1"></param>
        /// <param name="polyline2"></param>
        /// <param name="acc">数据精度</param>
        /// <returns></returns>
        public static bool ComparePolyline(this Polyline polyline1, Polyline polyline2, int acc)
        {
            double width1 = polyline1.Bounds.Value.GetWidth();
            double height1 = polyline1.Bounds.Value.GetHeight();
            double width2 = polyline2.Bounds.Value.GetWidth();
            double height2 = polyline2.Bounds.Value.GetHeight();
            if (Math.Abs(width1 - width2) >= acc || Math.Abs(height1 - height2) >= acc || polyline1.NumberOfVertices != polyline2.NumberOfVertices)
            {
                return false;
            }
            List<Point2d> points1 = polyline1.GetPointsBasedOnCenter(acc);
            List<Point2d> points2 = polyline2.GetPointsBasedOnCenter(acc);
            bool result = true;
            foreach (var item in points1)
            {
                if (!points2.Exists(p => item.X.Accuracy(acc) == p.X.Accuracy(acc) && item.Y.Accuracy(acc) == p.Y.Accuracy(acc)))
                {
                    result = false;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// 变换多段线
        /// </summary>
        /// <param name="polyline"></param>
        /// <param name="isMirrored">是否先镜像</param>
        /// <param name="angle">旋转角度</param>
        /// <returns></returns>
        public static Polyline Transform(this Polyline polyline, bool isMirrored, double angle)
        {
            Polyline polyline1;
            double radians = angle * (Math.PI / 180.0);
            if (isMirrored)
            {
                Plane plane = new Plane(polyline.Bounds.Value.GetCenter(), Vector3d.XAxis);
                Polyline polyline2 = polyline.GetTransformedCopy(Matrix3d.Mirroring(plane)) as Polyline;
                polyline1 = polyline2.GetTransformedCopy(Matrix3d.Rotation(radians, Vector3d.ZAxis, polyline.Bounds.Value.GetCenter())) as Polyline;
                plane.Dispose();
                polyline2.Dispose();
            }
            else
            {
                polyline1 = polyline.GetTransformedCopy(Matrix3d.Rotation(radians, Vector3d.ZAxis, polyline.Bounds.Value.GetCenter())) as Polyline;
            }
            return polyline1;
        }

        /// <summary>
        /// 获取多段线的原始角度
        /// </summary>
        /// <param name="polyline"></param>
        /// <returns></returns>
        public static int GetOriginalAngle(this Polyline polyline)
        {
            Point2d point1 = polyline.GetPoint2dAt(0);
            Point2d point2 = polyline.GetPoint2dAt(1);
            if (point1.X == point2.X)
            {
                return 0;
            }
            double k = Math.Abs((point2.Y - point1.Y) / (point2.X - point1.X));
            if (k > Math.Tan(89 * (Math.PI / 180.0)))
            {
                return 0;
            }
            int angle = (int)(Math.Atan(k) * (180.0 / Math.PI)).Accuracy(0);
            return angle;
        }

        /// <summary>
        /// 获取直线
        /// </summary>
        /// <param name="polyline"></param>
        /// <returns></returns>
        public static List<Line> GetLines(this Polyline polyline)
        {
            List<Line> lines = new List<Line>();
            int n = polyline.Closed ? polyline.NumberOfVertices : polyline.NumberOfVertices - 1;
            for (int i = 0; i < n; i++)
            {
                Line line = i == polyline.NumberOfVertices - 1
                    ? new Line(polyline.GetPoint2dAt(i).ToPoint3d(0), polyline.GetPoint2dAt(0).ToPoint3d(0))
                    : new Line(polyline.GetPoint2dAt(i).ToPoint3d(0), polyline.GetPoint2dAt(i + 1).ToPoint3d(0));
                lines.Add(line);
            }
            return lines;
        }
    }
}