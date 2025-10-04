using System.Windows.Forms;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace LightningCAD.Models
{
    public class LFlash : Transient
    {
        private Entity Entity;
        public LFlash(Entity entity, TransientDrawingMode mode)
        {
            Entity = entity;
            Transient.CapturedDrawable = this;
            TransientManager.CurrentTransientManager.AddTransient(this, mode, 0, new IntegerCollection());
        }
        protected override int SubSetAttributes(DrawableTraits traits)
        {
            traits.FillType = FillType.FillAlways;
            return (int)DrawableAttributes.IsAnEntity;
        }

        protected override void SubViewportDraw(ViewportDraw vd)
        {
            vd.Geometry.Draw(Entity);
            System.Drawing.Point point = Cursor.Position;
            Cursor.Position = new System.Drawing.Point(point.X, point.Y);
        }

        protected override bool SubWorldDraw(WorldDraw wd)
        {
            return false;
        }
        public void Update(Entity entity)
        {
            Entity.Dispose();
            Entity = entity;
            TransientManager.CurrentTransientManager.UpdateTransient(this, new IntegerCollection());
        }

        public void Delete()
        {
            TransientManager.CurrentTransientManager.EraseTransient(this, new IntegerCollection());
            Entity.Dispose();
            Dispose();
        }
    }
}