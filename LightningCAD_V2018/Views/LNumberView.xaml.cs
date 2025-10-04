using System.Linq;
using System.Windows;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Lightning.Extension;

using LightningCAD.Commands;
using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;

namespace LightningCAD.Views
{
    /// <summary>
    /// QuickNumberView.xaml 的交互逻辑
    /// </summary>
    public partial class LNumberView : ViewBase
    {
        private const string group = "QuickNumber";
        private const string prefixKey = "Prefix";
        private const string typeKey = "Type";
        private const string postfixKey = "Postfix";
        private const string currentKey = "Current";
        private const string heightKey = "Height";
        private const string borderKey = "Border";
        private const string leadKey = "Lead";

        public LNumberView()
        {
            InitializeComponent();
            Loaded += LNumberView_Loaded;
        }

        private void LNumberView_Loaded(object sender, RoutedEventArgs e)
        {
            //读取上次设置
            ReadSet();
            //改变类型时初始化
            type.SelectionChanged += (s1, e1) =>
            {
                switch (type.SelectedIndex)
                {
                    case 0:
                        current.Text = "1";
                        break;
                    case 1:
                        current.Text = "a";
                        break;
                    case 2:
                        current.Text = "A";
                        break;
                    default:
                        break;
                }
            };
            //取消CAD命令，关闭窗口，保存设置
            Closed += (s2, e2) => SaveSet();
        }

        /// <summary>
        /// 保存设置
        /// </summary>
        private void SaveSet()
        {
            Information.God.SetValue(group, prefixKey, prefix.Text);
            Information.God.SetValue(group, typeKey, type.SelectedIndex);
            Information.God.SetValue(group, postfixKey, postfix.Text);
            Information.God.SetValue(group, currentKey, current.Text);
            Information.God.SetValue(group, heightKey, textHeight.Text);
            Information.God.SetValue(group, borderKey, border.SelectedIndex);
            Information.God.SetValue(group, leadKey, lead.SelectedIndex);
        }

        /// <summary>
        /// 读取设置
        /// </summary>
        private void ReadSet()
        {
            Information.God.GetValue(group, prefixKey, out object prefixValue);
            Information.God.GetValue(group, typeKey, out object typeValue);
            Information.God.GetValue(group, postfixKey, out object postfixValue);
            Information.God.GetValue(group, currentKey, out object currentValue);
            Information.God.GetValue(group, heightKey, out object heightValue);
            Information.God.GetValue(group, borderKey, out object borderValue);
            Information.God.GetValue(group, leadKey, out object leadValue);

            prefix.Text = prefixValue?.ToString() ?? string.Empty;
            type.SelectedIndex = (int)(typeValue ?? 0);
            postfix.Text = postfixValue?.ToString() ?? string.Empty;
            current.Text = currentValue?.ToString() ?? "1";
            textHeight.Text = heightValue?.ToString() ?? "3.5";
            border.SelectedIndex = (int)(borderValue ?? 0);
            lead.SelectedIndex = (int)(leadValue ?? 0);
        }

        /// <summary>
        /// 编号自增
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private string AddOne(string str)
        {
            switch (type.Text)
            {
                case "a,b,c...":
                case "A,B,C...":
                    return StrAddOne(str);
                default:
                    return (int.Parse(str) + 1).ToString();
            }
        }

        private string StrAddOne(string str)
        {
            char A = type.Text == "a,b,c..." ? 'a' : 'A';
            char Z = type.Text == "a,b,c..." ? 'z' : 'Z';
            string letterTemp = str.Trim();
            int length = letterTemp.Length;

            int res = 0;
            for (int i = 0; i < length; i++)//先转成数字 a=1 z=26 az=52
            {
                res = res * 26 + letterTemp[i] - A + 1;
            }

            res++;

            string endCol;
            string endColSignal;
            int iCnt = res / 26;
            if (res >= 26 && res % 26 == 0)
            {
                int icell = iCnt - 2;
                endCol = (icell < 0 ? string.Empty : ((char)(A + icell)).ToString()) + Z;
            }
            else
            {
                endColSignal = iCnt == 0 ? "" : ((char)(A + (iCnt - 1))).ToString();
                int icell = res - iCnt * 26 - 1;
                if (icell < 0)
                    icell = 0;
                endCol = endColSignal + ((char)(A + icell)).ToString();
            }
            return endCol;
        }

        /// <summary>
        /// 获取矩形边框
        /// </summary>
        /// <param name="dBText"></param>
        /// <returns></returns>
        private static Polyline GetPolyline(DBText dBText)
        {
            Point3d min = dBText.Bounds.Value.MinPoint;
            Point3d max = dBText.Bounds.Value.MaxPoint;
            double width = max.X - min.X;
            double height = max.Y - min.Y;
            Polyline poly = new Polyline()
            {
                Closed = true,
            };
            poly.AddVertexAt(0, new Point2d(min.X - width * 0.08, min.Y - height * 0.2), 0, 0, 0);
            poly.AddVertexAt(1, new Point2d(max.X + width * 0.08, min.Y - height * 0.2), 0, 0, 0);
            poly.AddVertexAt(2, new Point2d(max.X + width * 0.08, max.Y + height * 0.2), 0, 0, 0);
            poly.AddVertexAt(3, new Point2d(min.X - width * 0.08, max.Y + height * 0.2), 0, 0, 0);
            return poly;
        }

        /// <summary>
        /// 获取圆形边框
        /// </summary>
        /// <param name="dBText"></param>
        /// <returns></returns>
        private static Circle GetCircle(DBText dBText)
        {
            Point3d min = dBText.Bounds.Value.MinPoint;
            Point3d max = dBText.Bounds.Value.MaxPoint;
            double R = min.DistanceTo(max);
            return new Circle()
            {
                Radius = R / 2 * 1.1,
                Center = new Point3d((min.X + max.X) / 2, (min.Y + max.Y) / 2, 0)
            };
        }

        /// <summary>
        /// 获取无边框的引线
        /// </summary>
        /// <param name="dBText"></param>
        /// <param name="point">起点</param>
        /// <returns></returns>
        private static Polyline GetNonLead(DBText dBText, Point3d point)
        {
            Polyline poly = new Polyline()
            {
                Closed = false
            };
            poly.AddVertexAt(0, point.ToPoint2d(), 0, 0, 0);
            Point3d min = dBText.Bounds.Value.MinPoint;
            Point3d max = dBText.Bounds.Value.MaxPoint;
            double height = max.Y - min.Y;
            Extents3d extents;
            if (dBText.Bounds == null)
            {
                return poly;
            }
            extents = dBText.Bounds.Value;

            if (extents.GetCenter().X < point.X)
            {
                poly.AddVertexAt(1, new Point2d(max.X, min.Y - height * 0.2), 0, 0, 0);
                poly.AddVertexAt(2, new Point2d(min.X, min.Y - height * 0.2), 0, 0, 0);
            }
            else
            {
                poly.AddVertexAt(1, new Point2d(min.X, min.Y - height * 0.2), 0, 0, 0);
                poly.AddVertexAt(2, new Point2d(max.X, min.Y - height * 0.2), 0, 0, 0);
            }
            return poly;
        }

        /// <summary>
        /// 获取矩形边框的引线
        /// </summary>
        /// <param name="dBText"></param>
        /// <param name="point">起点</param>
        /// <returns></returns>
        private static Polyline GetRecLead(DBText dBText, Point3d point)
        {
            Polyline poly = new Polyline()
            {
                Closed = false
            };
            poly.AddVertexAt(0, point.ToPoint2d(), 0, 0, 0);
            Point3d min = dBText.Bounds.Value.MinPoint;
            Point3d max = dBText.Bounds.Value.MaxPoint;
            double width = max.X - min.X;
            double height = max.Y - min.Y;
            Extents3d extents;
            if (dBText.Bounds == null)
            {
                return poly;
            }
            extents = dBText.Bounds.Value;

            if (extents.GetCenter().X < point.X)
            {
                poly.AddVertexAt(1, new Point2d(max.X + width * 0.08, min.Y - height * 0.4), 0, 0, 0);
                poly.AddVertexAt(2, new Point2d(min.X - width * 0.08, min.Y - height * 0.4), 0, 0, 0);
            }
            else
            {
                poly.AddVertexAt(1, new Point2d(min.X - width * 0.08, min.Y - height * 0.4), 0, 0, 0);
                poly.AddVertexAt(2, new Point2d(max.X + width * 0.08, min.Y - height * 0.4), 0, 0, 0);
            }
            return poly;
        }

        /// <summary>
        /// 获取圆形边框的引线
        /// </summary>
        /// <param name="dBText"></param>
        /// <param name="point">起点</param>
        /// <returns></returns>
        private static Polyline GetCirLead(DBText dBText, Point3d point)
        {
            Polyline poly = new Polyline()
            {
                Closed = false
            };
            poly.AddVertexAt(0, point.ToPoint2d(), 0, 0, 0);
            Point3d min = dBText.Bounds.Value.MinPoint;
            Point3d max = dBText.Bounds.Value.MaxPoint;
            double r = min.DistanceTo(max) / 2 * 1.1;
            Extents3d extents;
            if (dBText.Bounds == null)
            {
                return poly;
            }
            extents = dBText.Bounds.Value;

            Point3d center = extents.GetCenter();
            if (extents.GetCenter().X < point.X)
            {
                poly.AddVertexAt(1, new Point2d(center.X + r, center.Y - 1.2 * r), 0, 0, 0);
                poly.AddVertexAt(2, new Point2d(center.X - r, center.Y - 1.2 * r), 0, 0, 0);
            }
            else
            {
                poly.AddVertexAt(1, new Point2d(center.X - r, center.Y - 1.2 * r), 0, 0, 0);
                poly.AddVertexAt(2, new Point2d(center.X + r, center.Y - 1.2 * r), 0, 0, 0);
            }
            return poly;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //获取鼠标位置后，获取窗口参数
            while (true)
            {
                PromptPointOptions options = new PromptPointOptions("请指定标注位置");
                PromptPointResult promptPointResult = editor.GetPoint(options);
                if (promptPointResult.Status == PromptStatus.Cancel)
                {
                    return;
                }
                editor.WriteMessage("\n");
                if (promptPointResult.Status != PromptStatus.OK)
                {
                    continue;
                }

                //确保当前编号与所选格式一致
                switch (type.Text)
                {
                    case "a,b,c...":
                        bool isLower = current.Text.All(char.IsLower);
                        if (!isLower)
                        {
                            editor.WriteMessage("当前编号格式不正确!\n");
                            continue;
                        }
                        break;
                    case "A,B,C...":
                        bool isUpper = current.Text.All(char.IsUpper);
                        if (!isUpper)
                        {
                            editor.WriteMessage("当前编号格式不正确!\n");
                            continue;
                        }
                        break;
                    default:
                        bool isDigit = current.Text.All(char.IsDigit);
                        if (!isDigit)
                        {
                            editor.WriteMessage("当前编号格式不正确!\n");
                            continue;
                        }
                        break;
                }

                //确保字高正确
                if (!double.TryParse(textHeight.Text, out double height))
                {
                    editor.WriteMessage("字高不正确!\n");
                    continue;
                }

                string now = prefix.Text != "" ? prefix.Text + current.Text : current.Text;
                if (postfix.Text != "")
                {
                    now += postfix.Text;
                }

                LSetting.CreatStyles(document);
                //创建编号
                DBText dBText = new DBText()
                {
                    TextStyleId = document.GetLTextObjectId(),
                    TextString = now,
                    Height = height,
                    Justify = AttachmentPoint.MiddleCenter,
                };
                if (lead.Text == "有")
                {
                    DragLine dragLine = new DragLine(promptPointResult.Value, "指定标注位置");
                    PromptResult promptResult = editor.Drag(dragLine);
                    editor.WriteMessage("\n");
                    if (promptResult.Status != PromptStatus.OK)
                    {
                        continue;
                    }
                    dBText.AlignmentPoint = dragLine.end;
                    document.Drawing(dBText);
                    //画引线
                    switch (border.Text)
                    {
                        case "圆形":
                            document.Drawing(GetCirLead(dBText, promptPointResult.Value));
                            break;
                        case "矩形":
                            document.Drawing(GetRecLead(dBText, promptPointResult.Value));
                            break;
                        default:
                            document.Drawing(GetNonLead(dBText, promptPointResult.Value));
                            break;
                    }
                }
                else
                {
                    dBText.AlignmentPoint = promptPointResult.Value;
                    //画文字
                    document.Drawing(dBText);
                }
                //画边框
                switch (border.Text)
                {
                    case "圆形":
                        document.Drawing(GetCircle(dBText));
                        break;
                    case "矩形":
                        document.Drawing(GetPolyline(dBText));
                        break;
                    default:
                        break;
                }

                current.Text = AddOne(current.Text);
            }

        }
    }
}
