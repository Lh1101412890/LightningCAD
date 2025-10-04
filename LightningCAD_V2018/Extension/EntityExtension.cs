using System.Collections.Generic;
using System.Linq;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static partial class EntityExtension
    {
        /// <summary>
        /// 获取标记模型空间中的单个实体的矩形
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="offset">距离Entity的Bounds的边缘距离</param>
        /// <param name="colorEnum">标记矩形颜色</param>
        /// <param name="width">标记矩形线宽度</param>
        /// <returns></returns>
        public static Polyline GetBoundsRectangle(this Entity entity)
        {
            return entity.Bounds.Value.ToExtents2d().ToPolyline();
        }

        /// <summary>
        /// 获取中心点
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public static Point3d GetCenter(this Entity entity)
        {
            return entity.Bounds.Value.GetCenter();
        }

        /// <summary>
        /// 到目标target的距离（中心到中心）
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="target">目标实体</param>
        /// <returns></returns>
        public static double GetDistance(this Entity entity, Entity target)
        {
            return entity.GetCenter().DistanceTo(target.GetCenter());
        }

        /// <summary>
        /// 在entities中寻找最近的Entity
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="entities">搜索对象</param>
        /// <returns>最近的entity</returns>
        public static Entity GetNearest(this Entity entity, IEnumerable<Entity> entities)
        {
            return entities.Aggregate((nearest, others) => others.GetDistance(entity) < nearest.GetDistance(entity) ? others : nearest);
        }

        /// <summary>
        /// 通过Bounds判断entity是否在target内部
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="target">目标</param>
        /// <param name="tolerance">允许误差，如果小于0会被设为0</param>
        /// <returns></returns>
        public static bool IsInsideByBounds(this Entity entity, Entity target, double tolerance = LightningTolerance.Global)
        {
            double tMinX = target.Bounds.Value.MinPoint.X;
            double tMinY = target.Bounds.Value.MinPoint.Y;
            double tMinZ = target.Bounds.Value.MinPoint.Z;

            double tMaxX = target.Bounds.Value.MaxPoint.X;
            double tMaxY = target.Bounds.Value.MaxPoint.Y;
            double tMaxZ = target.Bounds.Value.MaxPoint.Z;

            double eMinX = entity.Bounds.Value.MinPoint.X;
            double eMinY = entity.Bounds.Value.MinPoint.Y;
            double eMinZ = entity.Bounds.Value.MinPoint.Z;

            double eMaxX = entity.Bounds.Value.MaxPoint.X;
            double eMaxY = entity.Bounds.Value.MaxPoint.Y;
            double eMaxZ = entity.Bounds.Value.MaxPoint.Z;

            if (tolerance < 0)
            {
                tolerance = 0;
            }

            return tMinX - tolerance <= eMinX && eMaxX <= tMaxX + tolerance && tMinY - tolerance <= eMinY && eMaxY <= tMaxY + tolerance && tMinZ - tolerance <= eMinZ && eMaxZ <= tMaxZ + tolerance;
        }
    }
}