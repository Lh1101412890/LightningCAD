using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Xml;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;
using LightningCAD.Models.DRBeamModels;

using Button = System.Windows.Controls.Button;
using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace LightningCAD.Views
{
    /// <summary>
    /// DRBeamView.xaml 的交互逻辑
    /// </summary>
    public partial class DRBeamView : ViewBase
    {
        public DRBeamView()
        {
            InitializeComponent();
            Entities = new List<Entity>();
            Hiddens = new List<Entity>();
            ComponentBeams = new List<ComponentBeam>();
            LFlashes = new List<LFlash>();
            Closed += DRBeamView_Closed;
        }

        private void DRBeamView_Closed(object sender, EventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            Visible();
            lFlash?.Delete();
            Entities.DisposeAll();
            foreach (var item in LFlashes)
            {
                item.Delete();
            }
            foreach (var item in ComponentBeams)
            {
                item.Line?.Dispose();
            }
        }

        private List<Entity> Entities;
        private readonly List<Entity> Hiddens;
        private LFlash lFlash;
        private readonly List<LFlash> LFlashes;
        private readonly List<ComponentBeam> ComponentBeams;

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (ComponentBeams.Count == 0)
            {
                LightningApp.ShowMsg("未识别到梁", 3);
                return;
            }
            PromptPointOptions options = new PromptPointOptions("指定基点");
            PromptPointResult result = editor.GetPoint(options);
            if (result.Status != PromptStatus.OK)
            {
                return;
            }
            //对齐点
            Point3d value = result.Value;

            // 导出数据至xml
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                AddExtension = true,
                DefaultExt = ".ctr",
                Filter = "CAD到Revit建模数据文件|*.ctr",
                //InitialDirectory = PCInfo.MyDocumentsDirectoryInfo + "\\Lightning\\LightningRevit"
            };
            if (saveFileDialog.ShowDialog() != true)
            {
                return;
            }
            string path = saveFileDialog.FileName;
            SaveCtrFile(path, ComponentBeams, value.ToPoint2d());
            editor.WriteMessage($"结构梁识别成功！文件路径：{path}\n");
        }

        // 点选图层
        private void Layer_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (Entities.Count == 0)
            {
                LightningApp.ShowMsg("未指定范围", 2);
                return;
            }
            if (sender is Button button)
            {
                PromptEntityResult promptEntityResult = editor.GetEntity(new PromptEntityOptions(button.Content.ToString()) { AllowNone = false, AllowObjectOnLockedLayer = true });
                if (promptEntityResult.Status != PromptStatus.OK)
                {
                    return;
                }
                string layer = "";
                Entity entity = document.GetObject(promptEntityResult.ObjectId) as Entity;
                layer = entity.Layer;
                entity.Dispose();

                switch (button.Content)
                {
                    case "点选柱线":
                        if (!columnBox.Text.Split('，').ToList().Exists(l => l == layer))
                        {
                            string str = columnBox.Text + "，" + layer;
                            columnBox.Text = str.Trim('，');
                        }
                        break;
                    case "点选墙线":
                        if (!wallBox.Text.Split('，').ToList().Exists(l => l == layer))
                        {
                            string str = wallBox.Text + "，" + layer;
                            wallBox.Text = str.Trim('，');
                        }
                        break;
                    case "点选梁线":
                        if (!beamBox.Text.Split('，').ToList().Exists(l => l == layer))
                        {
                            string str = beamBox.Text + "，" + layer;
                            beamBox.Text = str.Trim('，');
                        }
                        break;
                    default:
                        break;
                }
                editor.WriteMessage(layer + "\n");

                try
                {
                    using (DocumentLock @lock = document.LockDocument())
                    {
                        using (Transaction transaction = database.NewTransaction())
                        {
                            foreach (var item in Entities.Where(en => en.Layer == layer))
                            {
                                item.ObjectId.GetObject(OpenMode.ForWrite, false, true);
                                item.Visible = false;
                                Hiddens.Add(item);
                            }
                            transaction.Commit();
                        }
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    ex.LogTo(Information.God);
                }
            }
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            Visible();
            columnBox.Text = "";
            wallBox.Text = "";
            beamBox.Text = "";
        }

        private void Range_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions()
            {
                MessageForAdding = "选择范围"
            };
            PromptSelectionResult result = editor.GetSelection(promptSelectionOptions);
            if (result.Status != PromptStatus.OK)
            {
                return;
            }
            lFlash?.Delete();
            foreach (var item in LFlashes)
            {
                item?.Delete();
            }
            LFlashes.Clear();
            foreach (var item in ComponentBeams)
            {
                item.Line?.Dispose();
            }
            ComponentBeams.Clear();
            Visible();
            Hiddens.Clear();
            Entities.Clear();

            ObjectId[] objectIds = result.Value.GetObjectIds();
            List<Entity> entities = document.GetObjects(objectIds, false, true).Cast<Entity>().ToList();
            Entities = entities;
            double x1 = double.MaxValue;
            double y1 = double.MaxValue;
            double x2 = double.MinValue;
            double y2 = double.MinValue;
            foreach (var item in Entities)
            {
                Point3d minPoint = item.Bounds.Value.MinPoint;
                Point3d maxPoint = item.Bounds.Value.MaxPoint;
                if (minPoint.X < x1)
                {
                    x1 = minPoint.X;
                }
                if (minPoint.Y < y1)
                {
                    y1 = minPoint.Y;
                }
                if (maxPoint.X > x2)
                {
                    x2 = maxPoint.X;
                }
                if (maxPoint.Y > y2)
                {
                    y2 = maxPoint.Y;
                }
            }
            Polyline polyline = new Polyline();
            polyline.AddVertexAt(0, new Point2d(x1, y1), 0, 0, 0);
            polyline.AddVertexAt(0, new Point2d(x2, y1), 0, 0, 0);
            polyline.AddVertexAt(0, new Point2d(x2, y2), 0, 0, 0);
            polyline.AddVertexAt(0, new Point2d(x1, y2), 0, 0, 0);
            polyline.Closed = true;
            polyline.ConstantWidth = 150;
            polyline.Color = ColorEnum.Yellow.ToColor();
            lFlash = new LFlash(polyline, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Main);

            editor.WriteMessage("\n");
        }

        private void Visible_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            Visible();
        }

        private void Visible()
        {
            if (Hiddens.Count == 0)
            {
                CADApp.MainWindow.Handle.Focus();
                return;
            }
            try
            {
                using (DocumentLock @lock = document.LockDocument())
                {
                    using (Transaction transaction = database.NewTransaction())
                    {
                        foreach (var item in Hiddens)
                        {
                            item.ObjectId.GetObject(OpenMode.ForWrite, false, true);
                            item.Visible = true;
                        }
                        transaction.Commit();
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception exp)
            {
                exp.LogTo(Information.God);
            }
            Hiddens.Clear();
        }

        private void Hidden_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (Entities.Count == 0)
            {
                LightningApp.ShowMsg("未指定范围", 2);
                return;
            }
            List<string> layers = GetLayers();
            try
            {
                using (DocumentLock @lock = document.LockDocument())
                {
                    using (Transaction transaction = database.NewTransaction())
                    {
                        foreach (var item in Entities)
                        {
                            if (layers.Exists(l => l == item.Layer))
                            {
                                item.ObjectId.GetObject(OpenMode.ForWrite, false, true);
                                item.Visible = false;
                                Hiddens.Add(item);
                            }
                        }
                        transaction.Commit();
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }

        private List<string> GetLayers()
        {
            List<string> layers = new List<string>();
            foreach (var item in columnBox.Text.Split('，'))
            {
                layers.Add(item);
            }
            foreach (var item in wallBox.Text.Split('，'))
            {
                layers.Add(item);
            }
            foreach (var item in beamBox.Text.Split('，'))
            {
                layers.Add(item);
            }
            return layers;
        }

        private void BeamsSize_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (Entities.Count == 0)
            {
                LightningApp.ShowMsg("未指定范围", 2);
                return;
            }
            List<string> sizes = GetSizes();
            string sizestr = string.Join("，", sizes);
            editor.WriteMessage("梁尺寸统计: " + sizestr + "\n");
        }

        private List<string> GetSizes()
        {
            List<string> sizes = new List<string>();
            foreach (var item in Entities)
            {
                if (item is DBText dBText)
                {
                    string str = dBText.TextString;
                    if (str.Contains('x'))
                    {
                        string size = str.Contains('L') && str.Contains(' ') ? str.Split(' ').ToList().Find(t => t.Contains('x')) : str;
                        if (int.TryParse(size.Split('x')[0], out int _) && int.TryParse(size.Split('x')[1], out int _))
                        {
                            sizes.Add(size);
                        }
                    }
                }
            }
            return sizes.Distinct().OrderBy(s => s).ToList();
        }

        /// <summary>
        /// 合并构件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="components"></param>
        /// <param name="time">合并次数</param>
        /// <returns></returns>
        private static void Join<T>(List<T> components, int time) where T : ComponentBase
        {
            for (int k = 0; k < time; k++)
            {
                for (int i = 0; i < components.Count; i++)
                {
                    for (int j = i + 1; j < components.Count; j++)
                    {
                        if (Tools.TryExtend(components[i].Line, components[j].Line, out Line line, LightningTolerance.Structural))
                        {
                            components[i].Line.Dispose();
                            components[i].Line = line;
                            components[j].Line.Dispose();
                            components.RemoveAt(j);
                            j--;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 通过直线创建构件
        /// </summary>
        /// <typeparam name="T">ComponentWall或ComponentBeam</typeparam>
        /// <param name="lines"></param>
        /// <param name="max">最大尺寸</param>
        /// <param name="components">识别结果</param>
        private static void CreatComponents<T>(List<Line> lines, double max, out List<T> components) where T : ComponentBase, new()
        {
            components = new List<T>();
            for (int i = 0; i < lines.Count;)
            {
                double dis = double.MaxValue;
                int n = -1;
                for (int j = i + 1; j < lines.Count; j++)
                {
                    if (Tools.IsParallelUnion(lines[i], lines[j]))
                    {
                        double width = lines[i].GetDistance(lines[j]);
                        if (width > 140 && width <= max && width < dis)
                        {
                            dis = width;
                            n = j;
                        }
                    }
                }
                if (n != -1)
                {
                    T component = new T();
                    ComponentBase componentBase = component.Creat(lines[i], lines[n]);
                    lines.RemoveAt(n);
                    lines.RemoveAt(i);
                    components.Add(componentBase as T);
                }
                else
                {
                    lines.RemoveAt(i);
                }
            }
        }

        /// <summary>
        /// 保存数据至文件
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <param name="componentBeams">梁数据</param>
        /// <param name="value">对齐点</param>
        private static void SaveCtrFile(string path, List<ComponentBeam> componentBeams, Point2d value)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(Information.GetFileInfo("Files\\Beams.xml").FullName);
            XmlElement root = xml.DocumentElement;
            List<string> sizes = new List<string>();
            foreach (var item in componentBeams)
            {
                sizes.Add(item.Width.ToString() + "x" + item.Height.ToString());

                XmlNode beam = xml.CreateElement("BeamModel");
                root["BeamModels"].AppendChild(beam);

                XmlNode start = xml.CreateElement("Start");
                Point2d point1 = item.Line.StartPoint.ToPoint2d().Reduce(value);
                start.InnerText = point1.X.Accuracy(5).ToString() + "," + point1.Y.Accuracy(5).ToString();
                beam.AppendChild(start);

                XmlNode end = xml.CreateElement("End");
                Point2d point2 = item.Line.EndPoint.ToPoint2d().Reduce(value);
                end.InnerText = point2.X.Accuracy(5).ToString() + "," + point2.Y.Accuracy(5).ToString();
                beam.AppendChild(end);

                XmlNode width = xml.CreateElement("Width");
                width.InnerText = item.Width.ToString();
                beam.AppendChild(width);

                XmlNode height = xml.CreateElement("Height");
                height.InnerText = item.Height.ToString();
                beam.AppendChild(height);
            }
            foreach (var item in sizes.Distinct().OrderBy(s => s))
            {
                XmlNode size = xml.CreateElement("Size");
                size.InnerText = item;
                root["Sizes"].AppendChild(size);
            }
            xml.Save(path);
        }

        private void Recognize_Click(object sender, RoutedEventArgs e)
        {
            if (columnBox.Text == "" && wallBox.Text == "")
            {
                LightningApp.ShowMsg("请至少选择一个墙柱图层", 2);
                return;
            }
            if (!int.TryParse(wallmax.Text, out int wallMaxSize))
            {
                LightningApp.ShowMsg("请输入正确的最大墙厚（整数）", 2);
                return;
            }
            document.CreatLayer(LayerName.Beam, ColorEnum.Red);
            List<string> columnLayers = columnBox.Text.Split('，').ToList();
            List<string> wallLayers = wallBox.Text.Split('，').ToList();
            List<string> beamLayers = beamBox.Text.Split('，').ToList();
            //原始图形
            List<Polyline> cPolylines = new List<Polyline>();
            List<Polyline> wPolylines = new List<Polyline>();
            List<Line> wLines = new List<Line>();
            List<Line> bLines = new List<Line>();
            List<DBText> boTexts = new List<DBText>();

            //图形分类
            foreach (Entity entity in Entities)
            {
                if (entity is Polyline polyline)
                {
                    if (columnLayers.Exists(l => l == polyline.Layer))
                    {
                        cPolylines.Add(polyline);
                    }
                    if (wallLayers.Exists(l => l == polyline.Layer))
                    {
                        wPolylines.Add(polyline);
                    }
                }
                if (entity is Line line)
                {
                    if (wallLayers.Exists(l => l == line.Layer))
                    {
                        wLines.Add(line);
                    }
                    if (beamLayers.Exists(l => l == line.Layer))
                    {
                        bLines.Add(line);
                    }
                }
                if (entity is DBText dBText && dBText.TextString.Contains('x') && !dBText.TextString.Contains('L'))
                {
                    boTexts.Add(dBText);
                }
            }

            //转换柱
            List<Polyline> columns = new List<Polyline>();
            foreach (var item in cPolylines)
            {
                Polyline polyline = item.IsClosed(LightningTolerance.Structural) ? item.GetSimplest() : item.GetBoundsRectangle();
                columns.Add(polyline);
            }

            //处理墙，全部变为直线
            List<Line> wallLines = new List<Line>();
            foreach (var item in wPolylines)
            {
                if (item.IsClosed(LightningTolerance.Structural))
                {
                    Polyline polyline = item.GetSimplest();
                    wallLines.AddRange(polyline.GetLines());
                    polyline.Dispose();
                }
                else
                {
                    wallLines.AddRange(item.GetLines());
                }
            }
            foreach (var item in wLines)
            {
                wallLines.Add(new Line(item.StartPoint, item.EndPoint));
            }

            #region 转换墙
            List<Polyline> walls = new List<Polyline>();
            //去掉和柱子重复的线（如果两个点都在柱子内部）
            List<Line> wlines = wallLines.Where(l => !columns.Exists(c => l.IsInsideByBounds(c, LightningTolerance.Structural))).ToList();

            //垂直线
            List<Line> lines1 = wlines.Where(l => l.IsVertical()).OrderBy(l => l.StartPoint.X).ToList();
            //水平线
            List<Line> lines2 = wlines.Where(l => l.IsHorizontal()).OrderBy(l => l.StartPoint.Y).ToList();
            //斜线
            List<Line> lines3 = wlines.Except(lines1).Except(lines2).ToList();

            //识别Y向墙，第二个参数为最大墙宽度
            CreatComponents(lines1, wallMaxSize, out List<ComponentWall> cwalls1);
            //识别X向墙
            CreatComponents(lines2, wallMaxSize, out List<ComponentWall> cwalls2);
            //识别斜墙
            CreatComponents(lines3, wallMaxSize, out List<ComponentWall> cwalls3);

            //墙合并
            List<ComponentWall> componentWalls = new List<ComponentWall>();
            componentWalls.AddRange(cwalls1);
            componentWalls.AddRange(cwalls2);
            componentWalls.AddRange(cwalls3);
            Join(componentWalls, 5);

            foreach (var item in componentWalls)
            {
                walls.Add(item.GetPolyline());
            }
            #endregion

            var sizes = GetSizes();
            string size = sizes.Count != 0 ? sizes.Last() : "200x50";
            // 最大梁宽度
            double beamMaxSize = double.Parse(size.Split('x')[0]);

            //垂直线
            List<Line> lines4 = bLines.Where(l => l.IsVertical()).OrderBy(l => l.StartPoint.X).ToList();
            //水平线
            List<Line> lines5 = bLines.Where(l => l.IsHorizontal()).OrderBy(l => l.StartPoint.Y).ToList();
            //斜向线
            List<Line> lines6 = bLines.Except(lines4).Except(lines5).ToList();

            //识别Y向
            CreatComponents(lines4, beamMaxSize, out List<ComponentBeam> cbeams1);
            //识别X向
            CreatComponents(lines5, beamMaxSize, out List<ComponentBeam> cbeams2);
            //识别斜向
            CreatComponents(lines6, beamMaxSize, out List<ComponentBeam> cbeams3);

            //梁合并
            List<ComponentBeam> componentBeams = new List<ComponentBeam>();
            componentBeams.AddRange(cbeams1);
            componentBeams.AddRange(cbeams2);
            componentBeams.AddRange(cbeams3);

            Join(componentBeams, 5);

            #region 梁跨识别
            // 第一步：把墙柱作为支座，第一次提取梁
            // 第二步：把墙柱及已识别的梁作为支座，第二次提取梁
            // ……
            // （支座：墙、柱、梁的多段线）
            //
            // 提取梁：
            // 第一种：两端都有支座的梁
            // 第二种：第一跨起点有支座，最后一跨终点有支座，然后中间全是梁，并且能组合成一条直线
            // 第三种（悬挑梁）：一端有支座，另一端没有的，并且没有共线能延长的

            List<Polyline> beams = new List<Polyline>();
            List<Polyline> supports = new List<Polyline>();
            for (int time = 0; time < 15; time++)
            {
                if (componentBeams.Count == 0)
                {
                    break;
                }
                supports.Clear();
                supports.AddRange(columns);
                supports.AddRange(walls);
                supports.AddRange(beams);
                // 第一种
                for (int i = 0; i < componentBeams.Count; i++)
                {
                    if (componentBeams[i].HasSupportAtStartPoint(supports, wallMaxSize / 2) && componentBeams[i].HasSupportAtEndPoint(supports, wallMaxSize / 2))
                    {
                        beams.Add(componentBeams[i].GetPolyline());
                        ComponentBeams.Add(componentBeams[i]);
                        componentBeams.RemoveAt(i);
                        i--;
                    }
                }

                // 第二种
                for (int i = 0; i < componentBeams.Count; i++)
                {
                    if (componentBeams[i].HasSupportAtStartPoint(supports, wallMaxSize / 2))
                    {
                        if (componentBeams[i].Line.IsVertical())
                        {
                            if (componentBeams.Exists(b => 0 <= componentBeams[i].Line.GetDistance(b.Line) && componentBeams[i].Line.GetDistance(b.Line) <= LightningTolerance.Structural
                            && b.Line.StartPoint.Y >= componentBeams[i].Line.EndPoint.Y
                            && b.HasSupportAtEndPoint(supports, wallMaxSize / 2)))
                            {
                                //最近的单跨梁末尾
                                ComponentBeam end = componentBeams.Where(b => 0 <= componentBeams[i].Line.GetDistance(b.Line) && componentBeams[i].Line.GetDistance(b.Line) <= LightningTolerance.Structural
                                && b.Line.StartPoint.Y >= componentBeams[i].Line.EndPoint.Y
                                && b.HasSupportAtEndPoint(supports, wallMaxSize / 2))
                                    .OrderBy(b => b.Line.StartPoint.Y).First();

                                //中间的
                                List<ComponentBeam> middles = componentBeams.Where(b => 0 <= componentBeams[i].Line.GetDistance(b.Line) && componentBeams[i].Line.GetDistance(b.Line) <= LightningTolerance.Structural
                                && componentBeams[i].Line.EndPoint.Y <= b.Line.StartPoint.Y && b.Line.StartPoint.Y <= end.Line.StartPoint.Y
                                && componentBeams[i].Line.EndPoint.Y <= b.Line.EndPoint.Y && b.Line.EndPoint.Y <= end.Line.StartPoint.Y)
                                    .OrderBy(b => b.Line.StartPoint.Y).ToList();

                                if (middles.Count == 0)
                                {
                                    if (end.Line.StartPoint.Y - componentBeams[i].Line.EndPoint.Y > beamMaxSize + LightningTolerance.Structural)
                                    {
                                        continue;
                                    }
                                    var l = new Line(componentBeams[i].Line.StartPoint, end.Line.EndPoint);
                                    //componentBeams[i].Line.Dispose();
                                    componentBeams[i].Line = l;

                                    beams.Add(componentBeams[i].GetPolyline());
                                    ComponentBeams.Add(componentBeams[i]);

                                    componentBeams.RemoveAt(i);
                                    componentBeams.Remove(end);
                                    //end.Line.Dispose();
                                }
                                else
                                {
                                    ComponentBeam first = middles.First();
                                    ComponentBeam last = middles.Last();
                                    if (first.Line.StartPoint.Y - componentBeams[i].Line.EndPoint.Y > beamMaxSize + LightningTolerance.Structural
                                        || end.Line.StartPoint.Y - last.Line.EndPoint.Y > beamMaxSize + LightningTolerance.Structural)
                                    {
                                        continue;
                                    }

                                    bool can = true;
                                    for (int k = 0; k < middles.Count - 1; k++)
                                    {
                                        //中间的小段梁之间的距离大于最大梁宽
                                        if (middles[k + 1].Line.StartPoint.Y - middles[k].Line.EndPoint.Y > beamMaxSize + LightningTolerance.Structural)
                                        {
                                            can = false;
                                        }
                                    }
                                    if (!can)
                                    {
                                        continue;
                                    }

                                    var l = new Line(componentBeams[i].Line.StartPoint, end.Line.EndPoint);
                                    //componentBeams[i].Line.Dispose();
                                    componentBeams[i].Line = l;

                                    beams.Add(componentBeams[i].GetPolyline());
                                    ComponentBeams.Add(componentBeams[i]);

                                    componentBeams.RemoveAt(i);
                                    componentBeams.Remove(end);
                                    //end.Line.Dispose();
                                    componentBeams = componentBeams.Except(middles).ToList();
                                    foreach (var item in middles)
                                    {
                                        //item.Line.Dispose();
                                    }
                                }
                                i--;
                            }
                        }
                        else
                        {
                            if (componentBeams.Exists(b => 0 <= componentBeams[i].Line.GetDistance(b.Line) && componentBeams[i].Line.GetDistance(b.Line) <= LightningTolerance.Structural
                            && b.Line.StartPoint.X >= componentBeams[i].Line.EndPoint.X
                            && b.HasSupportAtEndPoint(supports, wallMaxSize / 2)))
                            {
                                //最近的单跨梁末尾
                                ComponentBeam end = componentBeams.Where(b => 0 <= componentBeams[i].Line.GetDistance(b.Line) && componentBeams[i].Line.GetDistance(b.Line) <= LightningTolerance.Structural
                                && b.Line.StartPoint.X >= componentBeams[i].Line.EndPoint.X
                                && b.HasSupportAtEndPoint(supports, wallMaxSize / 2))
                                    .OrderBy(b => b.Line.StartPoint.X).First();

                                //中间的
                                List<ComponentBeam> middles = componentBeams.Where(b => 0 <= componentBeams[i].Line.GetDistance(b.Line) && componentBeams[i].Line.GetDistance(b.Line) <= LightningTolerance.Structural
                                && componentBeams[i].Line.EndPoint.X <= b.Line.StartPoint.X && b.Line.StartPoint.X <= end.Line.StartPoint.X
                                && componentBeams[i].Line.EndPoint.X <= b.Line.EndPoint.X && b.Line.EndPoint.X <= end.Line.StartPoint.X)
                                    .OrderBy(b => b.Line.StartPoint.X).ToList();

                                if (middles.Count == 0)
                                {
                                    if (end.Line.StartPoint.X - componentBeams[i].Line.EndPoint.X > beamMaxSize + LightningTolerance.Structural)
                                    {
                                        continue;
                                    }
                                    var l = new Line(componentBeams[i].Line.StartPoint, end.Line.EndPoint);
                                    //componentBeams[i].Line.Dispose();
                                    componentBeams[i].Line = l;

                                    beams.Add(componentBeams[i].GetPolyline());
                                    ComponentBeams.Add(componentBeams[i]);

                                    componentBeams.RemoveAt(i);
                                    componentBeams.Remove(end);
                                    //end.Line.Dispose();
                                }
                                else
                                {
                                    ComponentBeam first = middles.First();
                                    ComponentBeam last = middles.Last();
                                    if (first.Line.StartPoint.X - componentBeams[i].Line.EndPoint.X > beamMaxSize + LightningTolerance.Structural
                                        || end.Line.StartPoint.X - last.Line.EndPoint.X > beamMaxSize + LightningTolerance.Structural)
                                    {
                                        continue;
                                    }

                                    bool can = true;
                                    for (int k = 0; k < middles.Count - 1; k++)
                                    {
                                        if (middles[k + 1].Line.StartPoint.X - middles[k].Line.EndPoint.X > beamMaxSize + LightningTolerance.Structural)
                                        {
                                            can = false;
                                        }
                                    }
                                    if (!can)
                                    {
                                        continue;
                                    }

                                    var l = new Line(componentBeams[i].Line.StartPoint, end.Line.EndPoint);
                                    //componentBeams[i].Line.Dispose();
                                    componentBeams[i].Line = l;

                                    beams.Add(componentBeams[i].GetPolyline());
                                    ComponentBeams.Add(componentBeams[i]);

                                    componentBeams.RemoveAt(i);
                                    componentBeams.Remove(end);
                                    componentBeams = componentBeams.Except(middles).ToList();

                                    //end.Line.Dispose();
                                    foreach (var item in middles)
                                    {
                                        //item.Line.Dispose();
                                    }
                                }
                                i--;
                            }

                        }
                    }
                }
            }
            // 第三种
            List<Polyline> beams1 = new List<Polyline>();
            for (int i = 0; i < componentBeams.Count; i++)
            {
                if (componentBeams[i].HasSupportAtStartPoint(supports, wallMaxSize / 2) || componentBeams[i].HasSupportAtEndPoint(supports, wallMaxSize / 2))
                {
                    beams1.Add(componentBeams[i].GetPolyline());
                    ComponentBeams.Add(componentBeams[i]);
                    componentBeams.RemoveAt(i);
                    i--;
                }
            }
            supports.AddRange(beams1);
            beams.AddRange(beams1);
            // 第一种
            for (int i = 0; i < componentBeams.Count; i++)
            {
                if (componentBeams[i].HasSupportAtStartPoint(supports, wallMaxSize / 2) && componentBeams[i].HasSupportAtEndPoint(supports, wallMaxSize / 2))
                {
                    ComponentBeams.Add(componentBeams[i]);
                    componentBeams.RemoveAt(i);
                    i--;
                }
            }
            #endregion

            //设置原位标注
            foreach (var item in boTexts)
            {
                Point3d center = item.GetCenter();
                ComponentBeam componentBeam;
                double rota = item.Rotation * 180 / Math.PI;
                if (Math.Abs(rota) < 1 || Math.Abs(rota) > 359)
                {
                    var beamss = ComponentBeams.Where(b => b.Line.IsHorizontal() && b.Line.StartPoint.Y > center.Y);
                    if (!beamss.Any())
                    {
                        continue;
                    }
                    componentBeam = beamss.Aggregate((nearest, next) =>
                    {
                        Point3d center1 = nearest.Line.GetCenter();
                        Point3d center2 = next.Line.GetCenter();
                        return center2.DistanceTo(center) < center1.DistanceTo(center) ? next : nearest;
                    });
                }
                else if (Math.Abs(rota - 90) < 1)
                {
                    var beamss = ComponentBeams.Where(b => b.Line.IsVertical() && b.Line.StartPoint.X < center.X);
                    if (!beamss.Any())
                    {
                        continue;
                    }
                    componentBeam = beamss.Aggregate((nearest, next) =>
                    {
                        Point3d center1 = nearest.Line.GetCenter();
                        Point3d center2 = next.Line.GetCenter();
                        return center2.DistanceTo(center) < center1.DistanceTo(center) ? next : nearest;
                    });
                }
                else
                {
                    var beamss = ComponentBeams.Where(b => !b.Line.IsHorizontal() && !b.Line.IsVertical());
                    if (!beamss.Any())
                    {
                        continue;
                    }
                    componentBeam = beamss.Aggregate((nearest, next) =>
                    {
                        Point3d center1 = nearest.Line.GetCenter();
                        Point3d center2 = next.Line.GetCenter();
                        return center2.DistanceTo(center) < center1.DistanceTo(center) ? next : nearest;
                    });
                }
                componentBeam?.SetOrthotopic(item.TextString);
            }

            foreach (var item in ComponentBeams)
            {
                Line line = item.Line;
                LFlash lFlash1 = new LFlash(line, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Highlight);
                LFlashes.Add(lFlash1);
                DBText text = new DBText()
                {
                    TextStyleId = document.GetLTextObjectId(),
                    TextString = item.Width.ToString(),
                    Justify = AttachmentPoint.TopCenter,
                    AlignmentPoint = line.GetCenter(),
                    Position = line.GetCenter(),
                    Height = 200,
                    Layer = LayerName.Beam,
                };
                LFlash lFlash = new LFlash(text, Autodesk.AutoCAD.GraphicsInterface.TransientDrawingMode.Highlight);
                LFlashes.Add(lFlash);
            }

            // 清理
            columns.DisposeAll();
            walls.DisposeAll();
            beams.DisposeAll();
            wallLines.DisposeAll();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (Entities.Count == 0)
            {
                LightningApp.ShowMsg("未指定范围", 2);
                return;
            }
            string str = code.Text.ToUpper().Replace('（', '(').Replace('）', ')');
            DBText text = Entities.Find(en => en is DBText dBText && dBText.TextString.StartsWith(str) && dBText.TextString.Contains('x')) as DBText;
            info.Text = text != null ? text.TextString : "未找到";
        }
    }
}