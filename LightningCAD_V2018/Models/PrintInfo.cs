using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

using LightningCAD.Extension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Models
{
    /// <summary>
    /// 打印信息模型
    /// </summary>
    public class PrintInfo : ILFlashModel
    {
        public PrintInfo(int n, Extents2d extents)
        {
            Number = n;
            Extents2d = extents;
            Size = PaperSize.A4;
            Direction = extents.GetHeight() > extents.GetWidth() ? PrintDirection.纵向 : PrintDirection.横向;
            LineWeight = false;
            Transparent = true;
            Style = PrintStyle.黑白;
            Note = "";
            HasFlash = false;
        }
        public void CreatFlash()
        {
            HasFlash = true;
            Polyline polyline = Extents2d.ToPolyline();
            polyline.Layer = "0";
            PolylineFlash = new LFlash(polyline, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Highlight);
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            polyline.GetPoint3dAt(3);
            DBText dBText = new DBText()
            {
                TextStyleId = document.GetLTextObjectId(),
                TextString = Number.ToString() + "," + Size + "," + Direction,
                Position = polyline.GetPoint3dAt(3),
                Height = (Extents2d.MaxPoint.Y - Extents2d.MinPoint.Y) * 0.12,
                Layer = "0",
            };
            TextFlash = new LFlash(dBText, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Highlight);
        }
        public void UpdateFlash()
        {
            Polyline polyline = Extents2d.ToPolyline();
            polyline.Layer = "0";
            PolylineFlash?.Update(polyline);
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            DBText dBText = new DBText()
            {
                TextStyleId = document.GetLTextObjectId(),
                TextString = Number.ToString() + "," + Size + "," + Direction,
                Position = polyline.GetPoint3dAt(3),
                Height = (Extents2d.MaxPoint.Y - Extents2d.MinPoint.Y) * 0.12,
                Layer = "0",
            };
            TextFlash?.Update(dBText);
        }
        public void DeleteFlash()
        {
            HasFlash = false;
            TextFlash?.Delete();
            PolylineFlash?.Delete();
        }

        /// <summary>
        /// 序号
        /// </summary>
        public int Number { get; set; }
        /// <summary>
        /// 打印纸张尺寸
        /// </summary>
        public PaperSize Size { get; set; }
        /// <summary>
        /// 打印方向，true代表纵向，false代表横向
        /// </summary>
        public PrintDirection Direction { get; set; }
        /// <summary>
        /// 是否打印线宽
        /// </summary>
        public bool LineWeight { get; set; }
        /// <summary>
        /// 是否透明度打印
        /// </summary>
        public bool Transparent { get; set; }
        /// <summary>
        /// 打印样式
        /// </summary>
        public PrintStyle Style { get; set; }
        /// <summary>
        /// 打印范围
        /// </summary>
        public Extents2d Extents2d { get; set; }
        /// <summary>
        /// 备注
        /// </summary>
        public string Note
        {
            get => note;
            set => note = value.Replace('\\', '_').Replace('/', '_').Replace(':', '_').Replace('*', '_').Replace('?', '_').Replace('"', '_').Replace('<', '_').Replace('>', '_').Replace('|', '_');
        }
        private string note;
        public bool HasFlash { get; set; }
        public LFlash TextFlash { get; set; }
        public LFlash PolylineFlash { get; set; }
    }

    /// <summary>
    /// 打印尺寸
    /// </summary>
    public enum PaperSize
    {
        A4,
        A3,
        A2,
        A1,
        A0,
    }

    /// <summary>
    /// 打印方向
    /// </summary>
    public enum PrintDirection
    {
        纵向,
        横向,
    }

    /// <summary>
    /// 打印样式
    /// </summary>
    public enum PrintStyle
    {
        黑白,
        彩色
    }

}
