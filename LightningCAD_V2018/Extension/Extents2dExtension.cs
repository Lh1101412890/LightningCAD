using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace LightningCAD.Extension
{
    public static class Extents2dExtension
    {
        /// <summary>
        /// 利用extents2D创建一个多段线（矩形）
        /// </summary>
        /// <param name="extents2d"></param>
        /// <returns></returns>
        public static Polyline ToPolyline(this Extents2d extents2d)
        {
            Polyline polyline = new Polyline()
            {
                Closed = true
            };
            polyline.AddVertexAt(0, extents2d.GetLeftTop(), 0, 0, 0);
            polyline.AddVertexAt(1, extents2d.MaxPoint, 0, 0, 0);
            polyline.AddVertexAt(2, extents2d.GetRightBottom(), 0, 0, 0);
            polyline.AddVertexAt(3, extents2d.MinPoint, 0, 0, 0);
            return polyline;
        }

        /// <summary>
        /// 获取中心点
        /// </summary>
        /// <param name="extents2d"></param>
        /// <returns></returns>
        public static Point2d GetCenter(this Extents2d extents2d)
        {
            return new Point2d((double)((extents2d.MaxPoint.X + extents2d.MinPoint.X) / 2), (double)((extents2d.MaxPoint.Y + extents2d.MinPoint.Y) / 2));
        }

        /// <summary>
        /// 获取左上角点
        /// </summary>
        /// <param name="extents2d"></param>
        /// <returns></returns>
        public static Point2d GetLeftTop(this Extents2d extents2d)
        {
            return new Point2d((double)extents2d.MinPoint.X, (double)extents2d.MaxPoint.Y);
        }

        /// <summary>
        /// 获取右下角点
        /// </summary>
        /// <param name="extents2d"></param>
        /// <returns></returns>
        public static Point2d GetRightBottom(this Extents2d extents2d)
        {
            return new Point2d((double)extents2d.MaxPoint.X, (double)extents2d.MinPoint.Y);
        }

        /// <summary>
        /// 获取宽度
        /// </summary>
        /// <param name="extents2d"></param>
        /// <returns></returns>
        public static double GetWidth(this Extents2d extents2d)
        {
            return extents2d.MaxPoint.X - extents2d.MinPoint.X;
        }

        /// <summary>
        /// 获取高度
        /// </summary>
        /// <param name="extents2d"></param>
        /// <returns></returns>
        public static double GetHeight(this Extents2d extents2d)
        {
            return extents2d.MaxPoint.Y - extents2d.MinPoint.Y;
        }
    }
}