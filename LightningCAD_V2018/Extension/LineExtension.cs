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
    }
}