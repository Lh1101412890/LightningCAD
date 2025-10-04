using System;
using System.Windows;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace LightningCAD.Views
{
    /// <summary>
    /// BreakLineView.xaml 的交互逻辑
    /// </summary>
    public partial class LBLineView : ViewBase
    {
        private const string group = "BreakLine";
        private const string symbolKey = "Symbol";
        private const string extendKey = "Extend";
        public LBLineView()
        {
            InitializeComponent();
            Loaded += LBLineView_Loaded;
            Closed += LBLineView_Closed;
        }

        private void LBLineView_Loaded(object sender, RoutedEventArgs e) => ReadSet();

        private void LBLineView_Closed(object sender, EventArgs e) => SaveSet();

        private void DrawingButton_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            Editor editor = document.Editor;
            bool v = double.TryParse(symbol.Text, out double ratio);
            bool h = double.TryParse(extend.Text, out double lengthRatio);
            if (v is false || ratio <= 0 || ratio >= 1)
            {
                editor.WriteMessage("符号比例只能为(0,1)之间的数！\n");
                return;
            }
            if (h is false || lengthRatio < 0)
            {
                editor.WriteMessage("延长比例必须≥0！\n");
                return;
            }
            PromptPointResult result = editor.GetPoint("折断线起点");
            if (result.Status != PromptStatus.OK) return;
            Point3d start = result.Value;

            DragLine dragLine = new DragLine(start, "折断线终点");
            PromptResult promptResult = editor.Drag(dragLine);
            if (promptResult.Status != PromptStatus.OK) return;
            Point3d end = dragLine.end;

            Point2d point0 = start.ToPoint2d();
            Point2d point5 = end.ToPoint2d();
            if (point0.GetDistanceTo(point5) == 0) return;

            Polyline polyline = new Polyline()
            {
                Closed = false
            };

            Point2d point1;
            Point2d point2;
            Point2d point3;
            Point2d point4;
            double x0 = point0.X;
            double y0 = point0.Y;
            double x5 = point5.X;
            double y5 = point5.Y;

            //垂直或近似垂直
            if (Math.Abs(x0 - x5) < Math.Abs(y0 - y5) * 0.018)
            {
                double length = Math.Abs(y0 - y5);
                double height = length * ratio;
                //垂直向下
                if (y0 - y5 > 0)
                {
                    point1 = new Point2d(x0, y0 - length / 2 + length * ratio / 2);
                    point2 = new Point2d(x0 - height, point1.Y - length * ratio / 4);
                    point3 = new Point2d(x0 + height, point2.Y - length * ratio / 2);
                    point4 = new Point2d(x0, point3.Y - length * ratio / 4);
                }
                //垂直向上       
                else
                {
                    point1 = new Point2d(x0, y0 + length / 2 - length * ratio / 2);
                    point2 = new Point2d(x0 + height, point1.Y + length * ratio / 4);
                    point3 = new Point2d(x0 - height, point2.Y + length * ratio / 2);
                    point4 = new Point2d(x0, point3.Y + length * ratio / 4);
                }
                point5 = new Point2d(x0, y5);
            }
            //水平或近似水平
            else if (Math.Abs(y0 - y5) < Math.Abs(x0 - x5) * 0.018)
            {
                double length = Math.Abs(x0 - x5);
                double height = length * ratio;
                //水平向左
                if (x0 - x5 > 0)
                {
                    point1 = new Point2d(x0 - length / 2 + length * ratio / 2, y0);
                    point2 = new Point2d(point1.X - length * ratio / 4, y0 - height);
                    point3 = new Point2d(point2.X - length * ratio / 2, y0 + height);
                    point4 = new Point2d(point3.X - length * ratio / 4, y0);
                }
                //水平向右
                else
                {
                    point1 = new Point2d(x0 + length / 2 - length * ratio / 2, y0);
                    point2 = new Point2d(point1.X + length * ratio / 4, y0 + height);
                    point3 = new Point2d(point2.X + length * ratio / 2, y0 - height);
                    point4 = new Point2d(point3.X + length * ratio / 4, y0);
                }
                point5 = new Point2d(x5, y0);
            }
            //倾斜
            else
            {
                double x_x = Math.Abs(x5 - x0);
                double y_y = Math.Abs(y5 - y0);
                double length = Math.Sqrt(Math.Pow(x_x, 2) + Math.Pow(y_y, 2));//两点直线长度
                double sin_a = Math.Round(y_y / length, 10);
                double cos_a = Math.Round(x_x / length, 10);
                double asin_a = Math.Asin(sin_a) * 180 / Math.PI;
                double atan_a = Math.Atan(0.25) * 180 / Math.PI;
                double cos_c = Math.Round(Math.Cos((90 - asin_a - atan_a) / 180 * Math.PI), 3);
                double length_1_4 = Math.Sqrt(Math.Pow(ratio * length, 2) + Math.Pow(0.25 * ratio * length, 2)) * cos_c;
                double length_1_2 = 0.5 * ratio * length * cos_a;

                //中间直线：y=kx+a
                //上方直线：y=kx+a+b
                //下方直线：y=kx+a-b
                double k = (y5 - y0) / (x5 - x0);
                double a = y0 - k * x0;
                double b = length * ratio / cos_a;

                double x1;
                double y1;
                double x2;
                double y2;
                double x3;
                double y3;
                double x4;
                double y4;

                //原点向第1象限
                if (x5 > x0 && y5 > y0)
                {
                    x1 = x0 + 0.5 * x_x - length_1_2;
                    y1 = k * x1 + a;
                    x2 = x0 + 0.5 * x_x - length_1_4;
                    y2 = k * x2 + a + b;
                    x3 = x0 + 0.5 * x_x + length_1_4;
                    y3 = k * x3 + a - b;
                    x4 = x0 + 0.5 * x_x + length_1_2;
                    y4 = k * x4 + a;
                }
                //原点向第2象限
                else if (x5 < x0 && y5 > y0)
                {
                    x1 = x0 - 0.5 * x_x + length_1_2;
                    y1 = k * x1 + a;
                    x2 = x0 - 0.5 * x_x + length_1_4;
                    y2 = k * x2 + a + b;
                    x3 = x0 - 0.5 * x_x - length_1_4;
                    y3 = k * x3 + a - b;
                    x4 = x0 - 0.5 * x_x - length_1_2;
                    y4 = k * x4 + a;
                }
                //原点向第3象限
                else if (x5 < x0 && y5 < y0)
                {
                    x1 = x0 - 0.5 * x_x + length_1_2;
                    y1 = k * x1 + a;
                    x2 = x0 - 0.5 * x_x + length_1_4;
                    y2 = k * x2 + a - b;
                    x3 = x0 - 0.5 * x_x - length_1_4;
                    y3 = k * x3 + a + b;
                    x4 = x0 - 0.5 * x_x - length_1_2;
                    y4 = k * x4 + a;
                }
                //原点向第4象限
                else
                {
                    x1 = x0 + 0.5 * x_x - length_1_2;
                    y1 = k * x1 + a;
                    x2 = x0 + 0.5 * x_x - length_1_4;
                    y2 = k * x2 + a - b;
                    x3 = x0 + 0.5 * x_x + length_1_4;
                    y3 = k * x3 + a + b;
                    x4 = x0 + 0.5 * x_x + length_1_2;
                    y4 = k * x4 + a;
                }
                point1 = new Point2d(x1, y1);
                point2 = new Point2d(x2, y2);
                point3 = new Point2d(x3, y3);
                point4 = new Point2d(x4, y4);
            }

            polyline.AddVertexAt(0, point0, 0, 0, 0);
            polyline.AddVertexAt(1, point1, 0, 0, 0);
            polyline.AddVertexAt(2, point2, 0, 0, 0);
            polyline.AddVertexAt(3, point3, 0, 0, 0);
            polyline.AddVertexAt(4, point4, 0, 0, 0);
            polyline.AddVertexAt(5, point5, 0, 0, 0);

            if (h)
            {
                Matrix3d scaliing = Matrix3d.Scaling(lengthRatio, new Point3d((x0 + x5) / 2, (y0 + y5) / 2, 0));
                polyline = polyline.GetTransformedCopy(scaliing) as Polyline;
            }

            document.Drawing(polyline);
        }

        /// <summary>
        /// 保存设置
        /// </summary>
        private void SaveSet()
        {
            Information.God.SetValue(group, symbolKey, symbol.Text);
            Information.God.SetValue(group, extendKey, extend.Text);
        }

        /// <summary>
        /// 读取设置
        /// </summary>
        private void ReadSet()
        {
            Information.God.GetValue(group, symbolKey, out object symbolValue);
            Information.God.GetValue(group, extendKey, out object extendValue);
            symbol.Text = symbolValue?.ToString() ?? "0.02";
            extend.Text = extendValue?.ToString() ?? "1.1";
        }
    }
}