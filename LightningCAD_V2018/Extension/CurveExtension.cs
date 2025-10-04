using Autodesk.AutoCAD.DatabaseServices;

using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static class CurveExtension
    {
        /// <summary>
        /// 判断曲线是否闭合
        /// </summary>
        /// <param name="curve"></param>
        /// <returns></returns>
        public static bool IsClosed(this Curve curve, double acc = LightningTolerance.Global)
        {
            return curve.Closed || curve.StartPoint.DistanceTo(curve.EndPoint) <= acc;
        }
    }
}