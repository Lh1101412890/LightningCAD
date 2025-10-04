using System;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static class Point2dExtension
    {
        /// <summary>
        /// 定义两个 Point2d 对象的加法
        /// </summary>
        /// <param name="point1">第一个 Point2d 对象</param>
        /// <param name="point2">第二个 Point2d 对象</param>
        /// <returns>相加后的 Point2d 对象</returns>
        public static Point2d Add(this Point2d point1, Point2d point2)
        {
            return new Point2d(point1.X + point2.X, point1.Y + point2.Y);
        }

        /// <summary>
        /// 定义两个 Point2d 对象的减法，point1 - point2
        /// </summary>
        /// <param name="point1">第一个 Point2d 对象</param>
        /// <param name="point2">第二个 Point2d 对象</param>
        /// <returns>相减后的 Point2d 对象</returns>
        public static Point2d Reduce(this Point2d point1, Point2d point2)
        {
            return new Point2d(point1.X - point2.X, point1.Y - point2.Y);
        }

        /// <summary>
        /// 转化为Point3d
        /// </summary>
        /// <param name="point"></param>
        /// <param name="z">高程</param>
        /// <returns></returns>
        public static Point3d ToPoint3d(this Point2d point, double z)
        {
            return new Point3d(point.X, point.Y, z);
        }

        /// <summary>
        /// 到指定点的斜率
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public static double GetK(this Point2d start, Point2d end)
        {
            double k = (end.Y - start.Y) / (end.X - start.X);
            //直线角度≥89°
            if (Math.Abs(k) >= Math.Tan(89 * (Math.PI / 180)))
            {
                return k > 0 ? double.PositiveInfinity : double.NegativeInfinity;
            }
            //直线角度≤1°
            return Math.Abs(k) <= Math.Tan(1 * (Math.PI / 180)) ? 0 : k;
        }

        /// <summary>
        /// 通过Bounds判断点是否在实体范围内
        /// </summary>
        /// <param name="point"></param>
        /// <param name="entity"></param>
        /// <param name="tolerance">允许误差，如果小于0会被设为0</param>
        /// <returns></returns>
        public static bool IsInsideByBounds(this Point2d point, Entity entity, double tolerance = LightningTolerance.Global)
        {
            Point3d max = entity.Bounds.Value.MaxPoint;
            Point3d min = entity.Bounds.Value.MinPoint;
            if (tolerance < 0)
            {
                tolerance = 0;
            }
            return min.X - tolerance <= point.X && point.X <= max.X + tolerance
                && min.Y - tolerance <= point.Y && point.Y <= max.Y + tolerance;
        }

        /// <summary>
        /// 判断点是否在polyline内部，该方法比较耗费性能
        /// </summary>
        /// <param name="point"></param>
        /// <param name="polyline"></param>
        /// <returns>点在多段线内部返回2,如果在边上返回1，否则返回0；若多段线不闭合返回-1</returns>
        public static int IsInsideByCompute(this Point2d point, Polyline polyline)
        {
            if (!polyline.Closed)
            {
                return -1;
            }
            double angle = 0;
            for (int i = 0; i < polyline.NumberOfVertices; i++)
            {
                Point2d point1 = polyline.GetPoint2dAt(i);
                Point2d point2 = i == polyline.NumberOfVertices - 1
                    ? polyline.GetPoint2dAt(0)
                    : polyline.GetPoint2dAt(i + 1);

                double a = Math.Sqrt(Math.Pow(point1.X - point2.X, 2) + Math.Pow(point1.Y - point2.Y, 2));
                double b = Math.Sqrt(Math.Pow(point.X - point2.X, 2) + Math.Pow(point.Y - point2.Y, 2));
                double c = Math.Sqrt(Math.Pow(point.X - point1.X, 2) + Math.Pow(point.Y - point1.Y, 2));

                // 角度
                double A = Math.Acos((b * b + c * c - a * a) / (2 * b * c)) * 180 / Math.PI;

                // 判断是逆时针还是顺时针旋转
                double cross = (point1.X - point.X) * (point2.Y - point.Y) - (point1.Y - point.Y) * (point2.X - point.X);

                angle += cross >= 0 ? A : -A;
            }

            // 如果点在外部，角度和为0
            return Math.Abs(angle) < 1 ? 0 : Math.Abs(angle) < 359 ? 1 : 2;
        }

        /// <summary>
        /// 判断点是否在polyline内部，该方法内部调用两次无误差的IsInsideByCompute方法，所以更耗费性能
        /// </summary>
        /// <param name="point"></param>
        /// <param name="polyline"></param>
        /// <param name="tolerance">允许误差</param>
        /// <returns>点在多段线内部返回2,如果在边上返回1，否则返回0；若多段线不闭合返回-1</returns>
        public static int IsInsideByCompute(this Point2d point, Polyline polyline, double tolerance)
        {
            if (!polyline.Closed)
            {
                return -1;
            }
            Polyline polyline1;
            if (tolerance > 0)
            {
                Line line = new Line(polyline.GetPoint2dAt(0).ToPoint3d(0), polyline.GetPoint2dAt(1).ToPoint3d(0));
                Line line1 = line.GetOffsetCurves(tolerance)[0] as Line;
                Point3d test = line1.GetMidpoint();
                line.Dispose();
                line1.Dispose();
                if (test.ToPoint2d().IsInsideByCompute(polyline) == 2)
                {
                    tolerance *= -1;
                }
                polyline1 = polyline.GetOffsetCurves(tolerance)[0] as Polyline;
            }
            else
            {
                polyline1 = polyline;
            }

            int result = point.IsInsideByCompute(polyline1);
            if (tolerance > 0)
            {
                polyline1.Dispose();
            }

            return result;
        }
    }
}