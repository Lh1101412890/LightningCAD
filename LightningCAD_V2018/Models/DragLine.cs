using System;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Models
{
    /// <summary>
    /// 拖拽线
    /// </summary>
    public class DragLine : EntityJig
    {
        public Point3d end;
        private Point3d start;
        private readonly string message;
        public DragLine(Point3d start, string message) : base(new Line())
        {
            Line line = Entity as Line;
            this.start = start;
            this.message = message;
            line.StartPoint = this.start;
            line.EndPoint = this.start;
        }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            JigPromptPointOptions options = new JigPromptPointOptions(message)
            {
                UserInputControls = UserInputControls.Accept3dCoordinates
            };
            PromptPointResult promptPointResult = prompts.AcquirePoint(options);
            Database database = CADApp.DocumentManager.MdiActiveDocument.Database;
            if (database.Orthomode)
            {
                //相应正交模式
                Point3d value = promptPointResult.Value;
                end = Math.Abs(value.X - start.X) < Math.Abs(value.Y - start.Y)
                    ? new Point3d(start.X, promptPointResult.Value.Y, promptPointResult.Value.Z)
                    : new Point3d(promptPointResult.Value.X, start.Y, promptPointResult.Value.Z);
            }
            else
            {
                end = promptPointResult.Value;
            }
            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            if (start != end)
            {
                ((Line)Entity).EndPoint = end;
            }
            return true;
        }
    }
}
