using System;

using Autodesk.AutoCAD.DatabaseServices;

using LightningCAD.Extension;

namespace LightningCAD.Models.DRBeamModels
{
    public class ComponentWall : ComponentBase
    {
        public ComponentWall()
        {

        }
        private ComponentWall(Line line1, Line line2)
        {
            Width = Math.Round(line1.GetDistance(line2));
            line1.GetCenterLine(line2, out Line line);
            Line = line;
        }

        public override ComponentBase Creat(Line line1, Line line2)
        {
            return new ComponentWall(line1, line2);
        }
    }
}