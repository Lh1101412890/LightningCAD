using System;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static class LineExtension
    {
        public static double GetK(this Line line)
        {
            return (line.EndPoint.Y - line.StartPoint.Y) / (line.EndPoint.X - line.StartPoint.X);
        }

        public static double GetB(this Line line)
        {
            return line.StartPoint.Y - line.GetK() * line.StartPoint.X;
        }

        /// <summary>
        /// 获取两直线之间的垂直距离
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns>不平行则返回-1</returns>
        public static double GetDistance(this Line line1, Line line2)
        {
            if (line1.IsVertical() && line2.IsVertical())
            {
                double x1 = line1.StartPoint.X;
                double x2 = line1.EndPoint.X;
                double x3 = line2.StartPoint.X;
                double x4 = line2.EndPoint.X;
                return Math.Abs((x3 + x4) / 2 - (x1 + x2) / 2);
            }
            else
            {
                double k1 = line1.GetK();
                double b1 = line1.GetB();
                double k2 = line2.GetK();
                double b2 = line2.GetB();
                return Math.Abs(k1 - k2) < LightningTolerance.Global ? Math.Abs(b1 - b2) / Math.Sqrt(k1 * k2 + 1) : -1;
            }
        }

        /// <summary>
        /// 判断直线是否是水平的，允许误差1°
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public static bool IsHorizontal(this Line line)
        {
            //1°直线的斜率
            return Math.Abs(line.GetK()) <= Math.Tan(1 * (Math.PI / 180));
        }

        /// <summary>
        /// 判断直线是否是垂直的，允许误差1°
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public static bool IsVertical(this Line line)
        {
            //89°直线斜率
            return Math.Abs(line.GetK()) >= Math.Tan(89 * (Math.PI / 180));
        }

        /// <summary>
        /// 获取直线中点
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public static Point3d GetMidpoint(this Line line)
        {
            return new Point3d((line.StartPoint.X + line.EndPoint.X) / 2, (line.StartPoint.Y + line.EndPoint.Y) / 2, (line.StartPoint.Z + line.EndPoint.Z) / 2);
        }

        /// <summary>
        /// 判断两条直线是否平行有交集
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static bool IsParallelUnion(this Line line1, Line line2)
        {
            if (line1.IsVertical())
            {
                if (line2.IsVertical())
                {
                    double y1 = Math.Min(line1.StartPoint.Y, line1.EndPoint.Y);
                    double y2 = Math.Max(line1.StartPoint.Y, line1.EndPoint.Y);
                    double y3 = Math.Min(line2.StartPoint.Y, line2.EndPoint.Y);
                    double y4 = Math.Max(line2.StartPoint.Y, line2.EndPoint.Y);
                    if ((y1 <= y3 && y3 <= y2)
                        || (y1 <= y4 && y4 <= y2)
                        || (y3 <= y1 && y2 <= y4))
                    {
                        return true;
                    }
                }
            }
            else if (line1.IsHorizontal())
            {
                if (line2.IsHorizontal())
                {
                    double x1 = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double x2 = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double x3 = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double x4 = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    if ((x1 <= x3 && x3 <= x2)
                        || (x1 <= x4 && x4 <= x2)
                        || (x3 <= x1 && x2 <= x4))
                    {
                        return true;
                    }
                }
            }
            else
            {
                double k1 = line1.GetK();
                double length = line1.GetDistance(line2);
                if (length >= 0)
                {
                    double b1 = line1.GetB();
                    double b2 = line2.GetB();
                    double x1 = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double x2 = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double x3 = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double x4 = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    double b = Math.Abs(b1 - b2);
                    double c = Math.Sqrt(b * b - length * length);
                    double p = 0.5 * (b + length + c);
                    double s = Math.Sqrt(p * (p - b) * (p - length) * (p - c));
                    double x = s * 2 / b;
                    if ((k1 > 0 && b1 > b2) || (k1 < 0 && b1 < b2))
                    {
                        x3 -= 2 * x;
                        x4 -= 2 * x;
                    }
                    else
                    {
                        x3 += 2 * x;
                        x4 += 2 * x;
                    }
                    if ((x1 <= x3 && x3 <= x2)
                        || (x1 <= x4 && x4 <= x2)
                        || (x3 <= x1 && x2 <= x4))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 判断两条直线是否可以延长
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static bool TryExtend(this Line line1, Line line2, out Line line, double acc = LightningTolerance.Global)
        {
            line = null;
            if (line1.IsHorizontal())
            {
                if (line2.IsHorizontal() && Math.Abs(line1.GetDistance(line2)) < acc)
                {
                    double l1min = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double l1max = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double l2min = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double l2max = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    if ((l1min - acc <= l2min && l2min <= l1max + acc) || (l1min - acc <= l2max && l2max <= l1max + acc) || (l2min - acc <= l1min && l1max <= l2max + acc))
                    {
                        double min = Math.Min(l1min, l2min);
                        double max = Math.Max(l1max, l2max);
                        double y = line1.StartPoint.Y;
                        line = new Line(new Point3d(min, y, 0), new Point3d(max, y, 0));
                        return true;
                    }
                }
            }
            else if (line1.IsVertical())
            {
                if (line2.IsVertical() && Math.Abs(line1.GetDistance(line2)) < acc)
                {
                    double l1min = Math.Min(line1.StartPoint.Y, line1.EndPoint.Y);
                    double l1max = Math.Max(line1.StartPoint.Y, line1.EndPoint.Y);
                    double l2min = Math.Min(line2.StartPoint.Y, line2.EndPoint.Y);
                    double l2max = Math.Max(line2.StartPoint.Y, line2.EndPoint.Y);
                    if ((l1min - acc <= l2min && l2min <= l1max + acc) || (l1min - acc <= l2max && l2max <= l1max + acc) || (l2min - acc <= l1min && l1max <= l2max + acc))
                    {
                        double min = Math.Min(l1min, l2min);
                        double max = Math.Max(l1max, l2max);
                        double x = line1.StartPoint.X;
                        line = new Line(new Point3d(x, min, 0), new Point3d(x, max, 0));
                        return true;
                    }
                }
            }
            else
            {
                double k1 = line1.GetK();
                double b1 = line1.GetB();
                double k2 = line2.GetK();
                double b2 = line2.GetB();
                if ((Math.Abs(k1 - k2) < 0.01) && Math.Abs(b1 - b2) < acc)
                {
                    double l1min = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double l1max = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double l2min = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double l2max = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    if ((l1min - acc <= l2min && l2min <= l1max + acc) || (l1min - acc <= l2max && l2max <= l1max + acc) || (l2min - acc <= l1min && l1max <= l2max + acc))
                    {
                        double min = Math.Min(l1min, l2min);
                        double max = Math.Max(l1max, l2max);
                        line = new Line(new Point3d(min, k1 * min + b1, 0), new Point3d(max, k1 * max + b1, 0));
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 获取两条直线的中心线
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <param name="line"></param>
        /// <returns></returns>
        public static bool GetCenterLine(this Line line1, Line line2, out Line line)
        {
            line = null;
            double max;
            double min;
            //斜率k为无穷
            if (line1.IsVertical())
            {
                if (!line2.IsVertical())
                {
                    return false;
                }
                max = Math.Max(Math.Max(Math.Max(line1.StartPoint.Y, line1.EndPoint.Y), line2.StartPoint.Y), line2.EndPoint.Y);
                min = Math.Min(Math.Min(Math.Min(line1.StartPoint.Y, line1.EndPoint.Y), line2.StartPoint.Y), line2.EndPoint.Y);
                double x = (line1.StartPoint.X + line1.EndPoint.X + line2.StartPoint.X + line2.EndPoint.X) / 4;
                line = new Line(new Point3d(x, min, 0), new Point3d(x, max, 0));
            }
            //斜率k为0
            else if (line1.IsHorizontal())
            {
                if (!line2.IsHorizontal())
                {
                    return false;
                }
                max = Math.Max(Math.Max(Math.Max(line1.StartPoint.X, line1.EndPoint.X), line2.StartPoint.X), line2.EndPoint.X);
                min = Math.Min(Math.Min(Math.Min(line1.StartPoint.X, line1.EndPoint.X), line2.StartPoint.X), line2.EndPoint.X);
                double y = (line1.StartPoint.Y + line1.EndPoint.Y + line2.StartPoint.Y + line2.EndPoint.Y) / 4;
                line = new Line(new Point3d(min, y, 0), new Point3d(max, y, 0));
            }
            else
            {
                double k1 = line1.GetK();
                double k2 = line2.GetK();
                if (Math.Abs(k1 - k2) >= 0.2)
                {
                    return false;
                }
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

                min = Math.Min(Math.Min(Math.Min(x1, x2), x3), x4);
                max = Math.Max(Math.Max(Math.Max(x1, x2), x3), x4);
                line = new Line(new Point3d(min, k1 * min + b_b, 0), new Point3d(max, k1 * max + b_b, 0));
            }
            return true;
        }

    }
}