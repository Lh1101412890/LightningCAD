using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static class Point3dExtension
    {
        /// <summary>
        /// 定义两个 Point3d 对象的加法
        /// </summary>
        /// <param name="point1">第一个 Point3d 对象</param>
        /// <param name="point2">第二个 Point3d 对象</param>
        /// <returns>相加后的 Point3d 对象</returns>
        public static Point3d Add(this Point3d point1, Point3d point2)
        {
            return new Point3d(point1.X + point2.X, point1.Y + point2.Y, point1.Z + point2.Z);
        }

        /// <summary>
        /// 定义两个 Point3d 对象的减法，point1 - point2
        /// </summary>
        /// <param name="point1">第一个 Point3d 对象</param>
        /// <param name="point2">第二个 Point3d 对象</param>
        /// <returns>相减后的 Point3d 对象</returns>
        public static Point3d Reduce(this Point3d point1, Point3d point2)
        {
            return new Point3d(point1.X - point2.X, point1.Y - point2.Y, point1.Z - point2.Z);
        }

        /// <summary>
        /// 3d点转换为XY平面的2d点
        /// </summary>
        /// <param name="point"></param>
        /// <returns></returns>
        public static Point2d ToPoint2d(this Point3d point)
        {
            return new Point2d(point.X, point.Y);
        }

        /// <summary>
        /// 通过Bounds判断点是否在实体范围内
        /// </summary>
        /// <param name="point"></param>
        /// <param name="entity"></param>
        /// <param name="tolerance">允许误差，如果小于0会被设为0</param>
        /// <returns></returns>
        public static bool IsInsideByBounds(this Point3d point, Entity entity, double tolerance = LightningTolerance.Global)
        {
            Point3d max = entity.Bounds.Value.MaxPoint;
            Point3d min = entity.Bounds.Value.MinPoint;
            if (tolerance < 0)
            {
                tolerance = 0;
            }
            return min.X - tolerance <= point.X && point.X <= max.X + tolerance
                && min.Y - tolerance <= point.Y && point.Y <= max.Y + tolerance
                && min.Z - tolerance <= point.Z && point.Z <= max.Z + tolerance;
        }

        /// <summary>
        /// 比较两个 Point3d 对象是否相等
        /// </summary>
        /// <param name="point1"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="z"></param>
        /// <returns></returns>
        public static bool IsEqualTo(this Point3d point1, double x, double y, double z)
        {
            return point1.X == x && point1.Y == y && point1.Z == z;
        }
    }
}