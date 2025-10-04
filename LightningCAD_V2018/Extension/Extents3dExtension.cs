using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LightningCAD.Extension
{
    public static class Extents3dExtension
    {
        /// <summary>
        /// 获取中心点
        /// </summary>
        /// <param name="extents3d"></param>
        /// <returns></returns>
        public static Point3d GetCenter(this Extents3d extents3d)
        {
            return new Point3d((double)((extents3d.MaxPoint.X + extents3d.MinPoint.X) / 2), (double)((extents3d.MaxPoint.Y + extents3d.MinPoint.Y) / 2), (double)((extents3d.MaxPoint.Z + extents3d.MinPoint.Z) / 2));
        }

        /// <summary>
        /// 获取宽度
        /// </summary>
        /// <param name="extents3d"></param>
        /// <returns></returns>
        public static double GetWidth(this Extents3d extents3d)
        {
            return extents3d.MaxPoint.X - extents3d.MinPoint.X;
        }

        /// <summary>
        /// 获取高度
        /// </summary>
        /// <param name="extents3d"></param>
        /// <returns></returns>
        public static double GetHeight(this Extents3d extents3d)
        {
            return extents3d.MaxPoint.Y - extents3d.MinPoint.Y;
        }

        /// <summary>
        /// 获取深度
        /// </summary>
        /// <param name="extents3d"></param>
        /// <returns></returns>
        public static double GetDepth(this Extents3d extents3d)
        {
            return extents3d.MaxPoint.Z - extents3d.MinPoint.Z;

        }

        /// <summary>
        /// 转化为XY平面的Extents2d
        /// </summary>
        /// <param name="extents3d"></param>
        /// <returns></returns>
        public static Extents2d ToExtents2d(this Extents3d extents3d)
        {
            return new Extents2d(extents3d.MinPoint.ToPoint2d(), extents3d.MaxPoint.ToPoint2d());
        }
    }
}
