using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

namespace LightningCAD.Models
{
    /// <summary>
    /// 拖拽矩形
    /// </summary>
    public class DragRectangle : EntityJig
    {
        public Point3d point2;
        private Point3d point1;
        private readonly string message;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="point1">基准点</param>
        /// <param name="message">提示信息</param>
        /// <param name="color">显示颜色</param>
        public DragRectangle(Point3d point1, string message, ColorEnum color = ColorEnum.Yellow) : base(new Polyline())
        {
            Polyline polyline = Entity as Polyline;
            this.point1 = point1;
            this.message = message;
            polyline.Color = color.ToColor();
            polyline.Closed = true;
            polyline.AddVertexAt(0, point1.ToPoint2d(), 0, 0, 0);
            polyline.AddVertexAt(1, point1.ToPoint2d(), 0, 0, 0);
            polyline.AddVertexAt(2, point1.ToPoint2d(), 0, 0, 0);
            polyline.AddVertexAt(3, point1.ToPoint2d(), 0, 0, 0);
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions options = new JigPromptPointOptions(message)
            {
                UserInputControls = UserInputControls.Accept3dCoordinates
            };
            PromptPointResult promptPointResult = prompts.AcquirePoint(options);
            point2 = promptPointResult.Value;
            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            if (point1 != point2)
            {
                ((Polyline)Entity).RemoveVertexAt(3);
                ((Polyline)Entity).RemoveVertexAt(2);
                ((Polyline)Entity).RemoveVertexAt(1);
                ((Polyline)Entity).AddVertexAt(1, new Point2d(point2.X, point1.Y), 0, 0, 0);
                ((Polyline)Entity).AddVertexAt(2, point2.ToPoint2d(), 0, 0, 0);
                ((Polyline)Entity).AddVertexAt(3, new Point2d(point1.X, point2.Y), 0, 0, 0);
            }
            return true;
        }
    }
}