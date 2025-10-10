using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using DRComponent = LightningCAD.Models.DRComponent;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace LightningCAD.Views
{
    /// <summary>
    /// DRWallColumnView.xaml 的交互逻辑
    /// </summary>
    public partial class DRWallColumnView : ViewBase
    {
        private List<DRComponent> Components { get; set; }
        private List<DRColumnModel> Columns { get; set; }
        private List<DRWallModel> Walls { get; set; }
        private List<DRColumnModel> ColumnDetails { get; set; }
        private List<DRWallInfo> WallInfos { get; set; }

        public DRWallColumnView()
        {
            InitializeComponent();
            Loaded += DebugColumnView_Loaded;
            Closed += DebugColumnView_Closed;
        }
        private void DebugColumnView_Loaded(object sender, RoutedEventArgs e)
        {
            Components = new List<DRComponent>();
            Columns = new List<DRColumnModel>();
            Walls = new List<DRWallModel>();
            ColumnDetails = new List<DRColumnModel>();
            WallInfos = new List<DRWallInfo>();
        }
        private void DebugColumnView_Closed(object sender, EventArgs e)
        {
            document.UnLockLayer(LayerName.Column);
            document.UnLockLayer(LayerName.ColumnDetail);
            document.UnLockLayer(LayerName.Wall);
            foreach (var item in Walls.Where(w => w.IsValid))
            {
                item.Delete();
            }
            foreach (var item in Columns.Where(c => c.IsValid))
            {
                item.Delete();
            }
            foreach (var item in ColumnDetails.Where(c => c.IsValid))
            {
                item.Delete();
            }
            document.DeleteLayer(LayerName.Column);
            document.DeleteLayer(LayerName.ColumnDetail);
            document.DeleteLayer(LayerName.Wall);
        }


        private void MenuItem1_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            document.UnLockLayer(LayerName.Column);
            document.UnLockLayer(LayerName.Wall);
            foreach (var item in Columns.Where(c => c.IsValid))
            {
                item.Delete();
            }
            foreach (var item in Walls.Where(w => w.IsValid))
            {
                item.Delete();
            }
            document.LockLayer(LayerName.Column);
            document.LockLayer(LayerName.Wall);
            Columns.Clear();
            Walls.Clear();
            ShowPlane();
        }
        private void MenuItem2_Click(object sender, RoutedEventArgs e)
        {
            Columns = Columns.Where(c => c.IsValid).ToList();
            Walls = Walls.Where(w => w.IsValid).ToList();
            ShowPlane();
        }
        private void PlaneDG_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (!(planeDG.SelectedItem is DRComponent component))
            {
                return;
            }
            if (component.IsValid)
            {
                document.Zoom(component.ComponentObjectId, 6);
            }
            else
            {
                if (component is DRWallModel wall)
                {
                    document.Zoom(wall.OriginalLineIds.First(), 2);
                }
                if (component is DRColumnModel column)
                {
                    document.Zoom(column.OriginalPolylineId, 2);
                }
            }
        }
        private void PlaneRead_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CADApp.MainWindow.Handle.Focus();

                //多段线（柱）
                List<Polyline> polylines = new List<Polyline>();
                //柱编号
                List<DBText> columnTexts = new List<DBText>();
                //直线（墙）
                List<Line> lines = new List<Line>();
                //墙编号
                List<DBText> wallTexts = new List<DBText>();

                #region 读取数据
                TypedValue[] typeds =
                {
                    new TypedValue((int)DxfCode.Operator,"<or"),
                    new TypedValue((int)DxfCode.Start,"lwpolyline"),
                    new TypedValue((int)DxfCode.Start,"text"),
                    new TypedValue((int)DxfCode.Start,"line"),
                    new TypedValue((int)DxfCode.Operator,"or>"),
                };
                SelectionFilter filter = new SelectionFilter(typeds);
                PromptSelectionOptions options = new PromptSelectionOptions() { MessageForAdding = "请选择要识别的平面图" };
                PromptSelectionResult result = editor.GetSelection(options, filter);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }

                var dBObjects = document.GetObjects(result.Value.GetObjectIds());
                foreach (DBObject dBObject in dBObjects)
                {
                    switch (dBObject)
                    {
                        case Polyline polyline:
                            polylines.Add(polyline);
                            break;
                        case Line line:
                            if (line.Layer != "DOTE" && line.Layer != "AXIS") //不读取轴网
                            {
                                lines.Add(line);
                            }
                            else
                            {
                                line.Dispose();
                            }
                            break;
                        case DBText text:
                            if (text.TextString.Contains('Z'))
                            {
                                columnTexts.Add(text);
                            }
                            else if (text.TextString.Contains('Q'))
                            {
                                wallTexts.Add(text);
                            }
                            else
                            {
                                text.Dispose();
                            }
                            break;
                        default:
                            dBObject.Dispose();
                            break;
                    }
                }

                #endregion

                #region 柱识别
                //去掉重复的多段线
                //如果第i个在后面某个内部，则去掉第i个
                for (int i = 0; i < polylines.Count; i++)
                {
                    for (int j = 0; j < polylines.Count; j++)
                    {
                        if (i == j) continue;
                        if (polylines[i].IsInsideByBounds(polylines[j], LightningTolerance.Structural))
                        {
                            polylines[i].Dispose();
                            polylines.RemoveAt(i);
                            i--;
                            break;
                        }
                    }
                }

                //去掉柱构件及已经识别的柱
                for (int i = 0; i < polylines.Count; i++)
                {
                    if (polylines[i].Layer == LayerName.Column || Columns.Exists(col => col.OriginalPolylineId == polylines[i].ObjectId))
                    {
                        polylines[i].Dispose();
                        polylines.RemoveAt(i);
                        i--;
                    }
                }

                //创建柱模型
                for (int i = 0; i < polylines.Count; i++)
                {
                    DBText dBText = polylines[i].GetNearest(columnTexts) as DBText;
                    DRColumnModel original = new DRColumnModel(polylines[i], dBText);
                    Columns.Add(original);
                }

                //清理数据
                polylines.DisposeAll();
                columnTexts.DisposeAll();

                //绘制柱构件
                document.UnLockLayer(LayerName.Column);
                foreach (var item in Columns.Where(c => c.IsValid))
                {
                    item.Drawing(LayerName.Column, ColorEnum.Red);
                }
                document.LockLayer(LayerName.Column);

                var enumerable = Columns.Where(col => !col.IsValid).ToList();
                string str = "";
                foreach (var item in enumerable)
                {
                    str += item.Name + ",";
                }
                editor.WriteMessage($"共识别柱 {Columns.Count} 个，有 {enumerable.Count} 个无效：\n{str}\n");
                #endregion

                #region 墙识别
                //去掉和柱子重复的线（如果两个点都在柱子内部）
                List<Line> wlines = lines.Where(l =>
                {
                    List<DRColumnModel> columnModels = Columns.Where(c => c.IsValid).ToList();
                    return !columnModels.Exists(c =>
                    {
                        Polyline polyline = document.GetObject(c.ComponentObjectId) as Polyline;
                        return l.IsInsideByBounds(polyline, LightningTolerance.Structural);
                    });
                }).ToList();

                //垂直线
                List<Line> lines1 = wlines.Where(l => l.IsVertical()).OrderBy(l => l.StartPoint.X).ToList();
                //水平线
                List<Line> lines2 = wlines.Where(l => l.IsHorizontal()).OrderBy(l => l.StartPoint.Y).ToList();
                //斜线
                List<Line> lines3 = wlines.Except(lines1).Except(lines2).ToList();

                //识别Y向墙
                List<DRWallModel> walls1 = new List<DRWallModel>();
                for (int i = 0; i < lines1.Count;)
                {
                    double dis = double.MaxValue;
                    int n = -1;
                    for (int j = i + 1; j < lines1.Count; j++)
                    {
                        if (Tools.IsParallelUnion(lines1[i], lines1[j]))
                        {
                            double width = lines1[i].GetDistance(lines1[j]);
                            if (width > 150 && width < 500 && width < dis)
                            {
                                dis = width;
                                n = j;
                            }
                        }
                    }
                    DRWallModel wallModel;
                    if (n == -1)
                    {
                        wallModel = new DRWallModel(lines1[i]);
                        lines1.RemoveAt(i);
                    }
                    else
                    {
                        wallModel = new DRWallModel(lines1[i], lines1[n]);
                        lines1.RemoveAt(n);
                        lines1.RemoveAt(i);
                    }
                    walls1.Add(wallModel);
                }

                //识别X向墙
                List<DRWallModel> walls2 = new List<DRWallModel>();
                for (int i = 0; i < lines2.Count;)
                {
                    double dis = 99999;
                    int n = -1;
                    for (int j = i + 1; j < lines2.Count; j++)
                    {
                        if (Tools.IsParallelUnion(lines2[i], lines2[j]))
                        {
                            double width = lines2[i].GetDistance(lines2[j]);
                            if (width > 150 && width < 500 && width < dis)
                            {
                                dis = width;
                                n = j;
                            }
                        }
                    }
                    DRWallModel wallModel;
                    if (n == -1)
                    {
                        wallModel = new DRWallModel(lines2[i]);
                        lines2.RemoveAt(i);
                    }
                    else
                    {
                        wallModel = new DRWallModel(lines2[i], lines2[n]);
                        lines2.RemoveAt(n);
                        lines2.RemoveAt(i);
                    }
                    walls2.Add(wallModel);
                }

                //识别斜墙
                List<DRWallModel> walls3 = new List<DRWallModel>();
                for (int i = 0; i < lines3.Count;)
                {
                    double dis = 99999;
                    int n = -1;
                    for (int j = i + 1; j < lines3.Count; j++)
                    {
                        if (Tools.IsParallelUnion(lines3[i], lines3[j]))
                        {
                            double width = lines3[i].GetDistance(lines3[j]);
                            if (width > 150 && width < 500 && width < dis)
                            {
                                dis = width;
                                n = j;
                            }
                        }
                    }
                    DRWallModel wallModel;
                    if (n == -1)
                    {
                        wallModel = new DRWallModel(lines3[i]);
                        lines3.RemoveAt(i);
                    }
                    else
                    {
                        wallModel = new DRWallModel(lines3[i], lines3[n]);
                        lines3.RemoveAt(n);
                        lines3.RemoveAt(i);
                    }
                    walls3.Add(wallModel);
                }

                //墙合并
                Walls.AddRange(walls1);
                Walls.AddRange(walls2);
                Walls.AddRange(walls3);
                document.UnLockLayer(LayerName.Wall);
                for (int i = 0; i < Walls.Count; i++)
                {
                    if (!Walls[i].IsValid) continue;
                    Line line1 = new Line(Walls[i].StartPoint.ToPoint3d(0), Walls[i].EndPoint.ToPoint3d(0));
                    for (int j = i + 1; j < Walls.Count; j++)
                    {
                        if (!Walls[j].IsValid) continue;
                        Line line2 = new Line(Walls[j].StartPoint.ToPoint3d(0), Walls[j].EndPoint.ToPoint3d(0));
                        if (Tools.TryExtend(line1, line2, out Line line, LightningTolerance.Structural))
                        {
                            Walls[i].StartPoint = line.StartPoint.ToPoint2d();
                            Walls[i].EndPoint = line.EndPoint.ToPoint2d();
                            Walls[i].Delete();
                            Walls[i].DrawingComponent();
                            Walls[j].Delete();
                            Walls.RemoveAt(j);
                            j--;
                        }
                        line2?.Dispose();
                        line?.Dispose();
                    }
                    line1?.Dispose();
                }

                //绘制墙构件
                foreach (var item in Walls.Where(w => w.IsValid))
                {
                    item.DrawingComponent();
                }
                //设置最近的墙编号
                var enumerable1 = Walls.Where(w => w.IsValid).Select(w => w.ComponentObjectId);
                List<Polyline> polylines1 = document.GetObjects(enumerable1).Cast<Polyline>().ToList();
                foreach (var item in wallTexts)
                {
                    var wall = item.GetNearest(polylines1) as Polyline;
                    Walls.Find(w => w.IsValid && w.ComponentObjectId == wall.ObjectId).SetName(item);
                }
                polylines1.DisposeAll();
                foreach (var item in Walls.Where(w => w.IsValid && w.Name == null))
                {
                    item.SetName();
                }
                foreach (var item in Walls.Where(w => w.IsValid))
                {
                    item.DrawingLinkline();
                }
                //清理
                lines.DisposeAll();
                wallTexts.DisposeAll();
                document.LockLayer(LayerName.Wall);

                ShowPlane();
                #endregion
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
        private void PlaneModify_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();

            if (Components.Count == 0)
            {
                editor.WriteMessage("请先识别平面图\n");
                return;
            }

            Polyline polyline;
            bool isColumn;
            PromptEntityOptions options1 = new PromptEntityOptions("请选择要修改的墙或柱");
            options1.SetRejectMessage("不是构件\n");
            options1.AddAllowedClass(typeof(Polyline), true);
            options1.AllowObjectOnLockedLayer = true;
            while (true)
            {
                PromptEntityResult result = editor.GetEntity(options1);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                polyline = document.GetObject(result.ObjectId) as Polyline;
                if (polyline.Layer != LayerName.Column && polyline.Layer != LayerName.Wall)
                {
                    polyline.Dispose();
                    editor.WriteMessage("不是构件\n");
                }
                else
                {
                    isColumn = polyline.Layer == LayerName.Column;
                    break;
                }
            }

            //修改柱编号
            if (isColumn)
            {
                DRColumnModel column = Columns.Find(c => c.IsValid && c.ComponentObjectId == polyline.ObjectId);
                editor.WriteMessage("柱构件 " + column.Name + "\n");

                DBText dBText;
                PromptEntityOptions options2 = new PromptEntityOptions("请选择修改后的柱编号");
                options2.SetRejectMessage("不是柱编号\n");
                options2.AddAllowedClass(typeof(DBText), true);
                options2.AllowObjectOnLockedLayer = true;
                while (true)
                {
                    PromptEntityResult result = editor.GetEntity(options2);
                    if (result.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    dBText = document.GetObject(result.ObjectId) as DBText;
                    if (dBText.TextString.Contains('Z'))
                    {
                        break;
                    }
                    dBText.Dispose();
                    editor.WriteMessage("不是柱编号\n");
                }

                if (column != null)
                {
                    using (DocumentLock @lock = document.LockDocument())
                    {
                        document.UnLockLayer(LayerName.Column);
                        column.UpdateName(dBText);
                        document.LockLayer(LayerName.Column);
                    }
                }
                editor.WriteMessage("改为 " + column.Name + "\n");
                polyline.Dispose();
                dBText.Dispose();
            }
            //修改墙编号
            else
            {
                DRWallModel wall = Walls.Find(w => w.IsValid && w.ComponentObjectId == polyline.ObjectId);
                string str = wall.Name == "" ? "默认" : wall.Name;
                editor.WriteMessage("墙构件 " + str + "\n");

                DBText dBText;
                PromptEntityOptions options2 = new PromptEntityOptions("请选择修改后的墙编号");
                options2.SetRejectMessage("不是墙编号\n");
                options2.AddAllowedClass(typeof(DBText), true);
                options2.AllowObjectOnLockedLayer = true;
                options2.AllowNone = true;
            A: PromptEntityResult result = editor.GetEntity(options2);
                if (result.Status == PromptStatus.OK)
                {
                    dBText = document.GetObject(result.ObjectId) as DBText;
                    if (!dBText.TextString.Contains('Q'))
                    {
                        dBText.Dispose();
                        editor.WriteMessage("不是墙编号\n");
                        goto A;
                    }
                    if (wall != null)
                    {
                        using (DocumentLock @lock = document.LockDocument())
                        {
                            document.UnLockLayer(LayerName.Wall);
                            wall.UpdateName(dBText);
                            document.LockLayer(LayerName.Wall);
                        }
                    }
                    editor.WriteMessage("改为 " + wall.Name + "\n");
                    polyline.Dispose();
                    dBText.Dispose();
                }
                else if (result.Status == PromptStatus.None)
                {
                    if (wall != null)
                    {
                        using (DocumentLock @lock = document.LockDocument())
                        {
                            document.UnLockLayer(LayerName.Wall);
                            wall.UpdateName(null);
                            document.LockLayer(LayerName.Wall);
                        }
                    }
                    editor.WriteMessage("改为 \"默认\"\n");
                    polyline.Dispose();
                }
            }
        }
        private void PlaneDelete_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (Components.Count == 0)
            {
                editor.WriteMessage("请先识别平面图\n");
                return;
            }

            PromptSelectionOptions prompt = new PromptSelectionOptions() { MessageForAdding = "请选择要删除的墙或柱" };
            SelectionFilter filter = new SelectionFilter(new TypedValue[]
            {
                new TypedValue((int)DxfCode.Operator,"<and"),
                new TypedValue((int)DxfCode.Start,"lwpolyline"),
                new TypedValue((int)DxfCode.Operator,"<or"),
                new TypedValue((int)DxfCode.LayerName,LayerName.Column),
                new TypedValue((int)DxfCode.LayerName,LayerName.Wall),
                new TypedValue((int )DxfCode.Operator,"or>"),
                new TypedValue((int )DxfCode.Operator,"and>"),
            });
            PromptSelectionResult result = editor.GetSelection(prompt, filter);
            if (result.Status != PromptStatus.OK)
            {
                return;
            }
            document.UnLockLayer(LayerName.Column);
            document.UnLockLayer(LayerName.Wall);
            foreach (var item in result.Value.GetObjectIds())
            {
                DRColumnModel columnModel = Columns.Find(col => col.IsValid && col.ComponentObjectId == item);
                if (columnModel != null)
                {
                    columnModel.Delete();
                    Columns.Remove(columnModel);
                }
                DRWallModel wallViewModel = Walls.Find(w => w.IsValid && w.ComponentObjectId == item);
                if (wallViewModel != null)
                {
                    wallViewModel.Delete();
                    Walls.Remove(wallViewModel);
                }
            }
            document.LockLayer(LayerName.Column);
            document.LockLayer(LayerName.Wall);
            ShowPlane();
        }
        private void Filter_Click(object sender, RoutedEventArgs e)
        {
            switch (filterBtn.Content.ToString())
            {
                case "有效":
                    filterBtn.Content = "全部";
                    ShowPlane();
                    break;
                case "无效":
                    filterBtn.Content = "有效";
                    ShowPlane();
                    break;
                default:
                    filterBtn.Content = "无效";
                    ShowPlane();
                    break;
            }
        }
        private void Category_Click(object sender, RoutedEventArgs e)
        {
            switch (categoryBtn.Content.ToString())
            {
                case "墙":
                    categoryBtn.Content = "墙柱";
                    ShowPlane();
                    break;
                case "柱":
                    categoryBtn.Content = "墙";
                    ShowPlane();
                    break;
                default:
                    categoryBtn.Content = "柱";
                    ShowPlane();
                    break;
            }
        }
        private void ShowPlane()
        {
            Components.Clear();
            switch (filterBtn.Content.ToString())
            {
                case "有效":
                    switch (categoryBtn.Content.ToString())
                    {
                        case "墙":
                            Components.AddRange(Walls.Where(w => w.IsValid));
                            break;
                        case "柱":
                            Components.AddRange(Columns.Where(c => c.IsValid));
                            break;
                        default:
                            Components.AddRange(Columns.Where(c => c.IsValid));
                            Components.AddRange(Walls.Where(w => w.IsValid));
                            break;
                    }
                    break;
                case "无效":
                    switch (categoryBtn.Content.ToString())
                    {
                        case "墙":
                            Components.AddRange(Walls.Where(w => !w.IsValid));
                            break;
                        case "柱":
                            Components.AddRange(Columns.Where(c => !c.IsValid));
                            break;
                        default:
                            Components.AddRange(Columns.Where(c => !c.IsValid));
                            Components.AddRange(Walls.Where(w => !w.IsValid));
                            break;
                    }
                    break;
                default:
                    switch (categoryBtn.Content.ToString())
                    {
                        case "墙":
                            Components.AddRange(Walls);
                            break;
                        case "柱":
                            Components.AddRange(Columns);
                            break;
                        default:
                            Components.AddRange(Columns);
                            Components.AddRange(Walls);
                            break;
                    }
                    break;
            }
            planeDG.DataContext = null;
            planeDG.DataContext = Components;
        }
        private void Compare_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Focus();
            if (Components.Count == 0)
            {
                editor.WriteMessage("请先识别平面图\n");
                return;
            }
            foreach (var item in planeDG.Items)
            {
                if (planeDG.ItemContainerGenerator.ContainerFromItem(item) is DataGridRow row)
                {
                    row.Background = new SolidColorBrush(Colors.Transparent);
                }
            }
            try
            {
                //柱比较
                if (!double.TryParse(multiple.Text, out double m))
                {
                    editor.WriteMessage("请输入正确的柱放大整数\n");
                }
                else
                {
                    if (ColumnDetails.Count == 0)
                    {
                        editor.WriteMessage("未识别柱大样\n");
                    }
                    else
                    {
                        foreach (var plane in Columns.Where(c => c.IsValid))
                        {
                            var detail = ColumnDetails.Find(cd => cd.IsValid && cd.Name == plane.Name);
                            switch (detail)
                            {
                                case null:
                                    plane.IsDifferent = true;
                                    break;
                                default:
                                    plane.IsDifferent = !Compare(plane, detail, out bool isMir, out double angle);
                                    break;
                            }
                        }
                        editor.WriteMessage($"柱比较完成\n");
                    }
                }

                //墙比较
                if (WallInfos.Count == 0)
                {
                    editor.WriteMessage("未识别墙表\n");
                }
                else
                {
                    foreach (var item in Walls.Where(w => w.IsValid))
                    {
                        switch (item.Name)
                        {
                            case "Q":
                                {
                                    if (Math.Abs(WallInfos[0].Thickness - item.Width) >= 1)
                                    {
                                        //墙厚不一致
                                        item.IsDifferent = true;
                                    }
                                    else
                                    {
                                        item.IsDifferent = false;
                                    }
                                    break;
                                }

                            default:
                                {
                                    DRWallInfo wallInfo = WallInfos.Find(w => w.Name == item.Name);
                                    switch (wallInfo)
                                    {
                                        case null:
                                            {
                                                //墙厚不一致
                                                item.IsDifferent = true;
                                                break;
                                            }

                                        default:
                                            {
                                                if (Math.Abs(wallInfo.Thickness - item.Width) >= 1)
                                                {
                                                    //墙厚不一致
                                                    item.IsDifferent = true;
                                                }
                                                else
                                                {
                                                    item.IsDifferent = false;
                                                }
                                                break;
                                            }
                                    }
                                    break;
                                }
                        }
                    }
                    editor.WriteMessage("墙比较完成\n");
                }
            }
            catch (Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
        /// <summary>
        /// 比较平面柱和大样柱
        /// </summary>
        /// <param name="columnModel"></param>
        /// <param name="detailModel"></param>
        /// <param name="isMir">是否先镜像</param>
        /// <param name="angle">旋转角度</param>
        /// <returns></returns>
        private bool Compare(DRColumnModel columnModel, DRColumnModel detailModel, out bool isMir, out double angle)
        {
            Polyline column = document.GetObject(columnModel.ComponentObjectId) as Polyline;
            Polyline detail = document.GetObject(detailModel.ComponentObjectId) as Polyline;

            double m = double.Parse(multiple.Text);
            isMir = false;
            angle = 0;
            bool result = false;
            int oriAngle = column.GetOriginalAngle();
            int k = oriAngle == 0 ? 2 : 4;
            for (int i = 0; i < 2; i++)
            {
                if (result)
                {
                    break;
                }
                for (int j = 0; j < 4; j++)
                {
                    isMir = i >= k / 2;
                    angle = i % 2 == 0 ? oriAngle + j * 90 : 90 - oriAngle + j * 90;
                    Polyline polyline = detail.Transform(isMir, angle);
                    Polyline entity = polyline.GetTransformedCopy(Matrix3d.Scaling(1 / m, detail.Bounds.Value.GetCenter())) as Polyline;
                    polyline.Dispose();
                    if (column.ComparePolyline(entity, 5))
                    {
                        entity.Dispose();
                        result = true;
                        break;
                    }
                    entity.Dispose();
                }
            }
            column.Dispose();
            detail.Dispose();
            if (result == false)
            {
                isMir = false;
                angle = 0;
            }
            return result;
        }
        private static bool IsValid(string str)
        {
            return !str.Contains('\\')
                && !str.Contains(':')
                && !str.Contains('{')
                && !str.Contains('}')
                && !str.Contains('[')
                && !str.Contains(']')
                && !str.Contains('|')
                && !str.Contains(';')
                && !str.Contains('<')
                && !str.Contains('>')
                && !str.Contains('?')
                && !str.Contains('`')
                && !str.Contains('~');
        }
        private void Export_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (!double.TryParse(multiple.Text, out double m))
            {
                editor.WriteMessage("请输入正确的柱放大整数\n");
                return;
            }
            try
            {
                if (ColumnDetails.Any(cd => cd.IsValid) && Columns.Any(c => c.IsValid) || (WallInfos.Count > 0 && Walls.Any(c => c.IsValid)))
                {
                    string prefix;
                    //设置前缀
                    PromptStringOptions options1 = new PromptStringOptions("指定柱编号前缀")
                    {
                        AllowSpaces = true,
                        DefaultValue = "F1-",
                        UseDefaultValue = true
                    };
                    while (true)
                    {
                        PromptResult result = editor.GetString(options1);
                        if (result.Status != PromptStatus.OK)
                        {
                            return;
                        }
                        prefix = result.StringResult;
                        //Revit中不能含有的字符
                        if (IsValid(prefix))
                        {
                            break;
                        }
                        editor.WriteMessage("不能含有\\:{}[]|;<>?`~\n");
                    }

                    //设置后缀
                    string postfix;
                    PromptStringOptions options2 = new PromptStringOptions("指定柱编号后缀")
                    {
                        AllowSpaces = true,
                    };
                    while (true)
                    {
                        PromptResult result = editor.GetString(options2);
                        if (result.Status != PromptStatus.OK)
                        {
                            return;
                        }
                        postfix = result.StringResult;
                        //Revit中不能含有的字符
                        if (IsValid(postfix))
                        {
                            break;
                        }
                        editor.WriteMessage("不能含有\\:{}[]|;<>?`~\n");
                    }

                    //指定基点
                    PromptPointOptions options = new PromptPointOptions("指定基点");
                    PromptPointResult result1 = editor.GetPoint(options);
                    if (result1.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    Point3d value = result1.Value;

                    //导出数据至xml
                    XmlDocument xml = new XmlDocument();
                    xml.Load(Information.GetFileInfo("Files\\WallsColumns.xml").FullName);
                    XmlElement root = xml.DocumentElement;
                    root["Prefix"].InnerText = prefix;
                    root["Postfix"].InnerText = postfix;
                    //柱大样
                    foreach (var item in ColumnDetails.Where(cd => cd.IsValid))
                    {
                        XmlNode xmlNode = xml.CreateElement("Detail");
                        XmlNode name = xml.CreateElement("Name");
                        name.InnerText = item.Name;
                        xmlNode.AppendChild(name);
                        XmlNode points = xml.CreateElement("Points");
                        xmlNode.AppendChild(points);
                        Polyline polyline = document.GetObject(item.ComponentObjectId) as Polyline;
                        Polyline entity = polyline.GetTransformedCopy(Matrix3d.Scaling(1 / m, polyline.Bounds.Value.GetCenter())) as Polyline;
                        polyline.Dispose();
                        List<Point2d> point2ds = entity.GetPointsBasedOnCenter(5);
                        entity.Dispose();
                        foreach (var point2D in point2ds)
                        {
                            XmlNode point = xml.CreateElement("Point");
                            point.InnerText = point2D.X.ToString() + "," + point2D.Y.ToString();
                            points.AppendChild(point);
                        }
                        root["Details"].AppendChild(xmlNode);
                    }
                    //平面柱
                    foreach (var item in Columns.Where(c => c.IsValid))
                    {
                        XmlNode xmlNode = xml.CreateElement("Column");
                        XmlNode name = xml.CreateElement("Name");
                        name.InnerText = item.Name;
                        xmlNode.AppendChild(name);
                        XmlNode point = xml.CreateElement("Point");
                        Polyline polyline = document.GetObject(item.ComponentObjectId) as Polyline;
                        Point3d center = polyline.Bounds.Value.GetCenter();
                        polyline.Dispose();
                        double x = (center.X - value.X).Accuracy(5);
                        double y = (center.Y - value.Y).Accuracy(5);
                        point.InnerText = x.ToString() + "," + y.ToString();
                        xmlNode.AppendChild(point);
                        XmlNode isMir = xml.CreateElement("IsMirror");
                        XmlNode angle = xml.CreateElement("Angle");
                        DRColumnModel columnModel = ColumnDetails.Find(cd => cd.IsValid && cd.Name == item.Name);
                        if (columnModel != null)
                        {
                            Compare(item, columnModel, out bool mir, out double ang);
                            isMir.InnerText = mir.ToString();
                            angle.InnerText = ang.ToString();
                        }
                        else
                        {
                            isMir.InnerText = false.ToString();
                            angle.InnerText = "0";
                        }
                        xmlNode.AppendChild(isMir);
                        xmlNode.AppendChild(angle);
                        root["Columns"].AppendChild(xmlNode);
                    }
                    //墙信息
                    foreach (var item in WallInfos)
                    {
                        XmlNode xmlNode = xml.CreateElement("WallInfo");
                        XmlNode name = xml.CreateElement("Name");
                        name.InnerText = item.Name;
                        xmlNode.AppendChild(name);
                        XmlNode width = xml.CreateElement("Width");
                        width.InnerText = item.Thickness.ToString();
                        xmlNode.AppendChild(width);
                        root["WallInfos"].AppendChild(xmlNode);
                    }
                    //平面墙
                    foreach (var item in Walls.Where(w => w.IsValid))
                    {
                        XmlNode xmlNode = xml.CreateElement("Wall");
                        XmlNode name = xml.CreateElement("Name");
                        name.InnerText = item.Name == "Q" ? "Q1" : item.Name;
                        xmlNode.AppendChild(name);
                        XmlNode start = xml.CreateElement("Start");
                        string x1 = (item.StartPoint.X - value.X).Accuracy(5).ToString();
                        string y1 = (item.StartPoint.Y - value.Y).Accuracy(5).ToString();
                        start.InnerText = x1 + "," + y1;
                        xmlNode.AppendChild(start);
                        XmlNode end = xml.CreateElement("End");
                        string x2 = (item.EndPoint.X - value.X).Accuracy(5).ToString();
                        string y2 = (item.EndPoint.Y - value.Y).Accuracy(5).ToString();
                        end.InnerText = x2 + "," + y2;
                        xmlNode.AppendChild(end);
                        root["Walls"].AppendChild(xmlNode);
                    }
                    //保存文件
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
                    xml.Save(path);
                    editor.WriteMessage($"柱大样导出成功！文件路径：{path}\n");
                }
            }
            catch (Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }


        private void DetailDG_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (!(detailDG.SelectedItem is DRColumnModel columnModel))
            {
                return;
            }
            if (columnModel.IsValid)
            {
                document.Zoom(columnModel.ComponentObjectId, 5);
            }
            else
            {
                document.Zoom(columnModel.OriginalPolylineId, 3);
            }
        }
        private void DetailRead_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CADApp.MainWindow.Handle.Focus();

                TypedValue[] typeds =
                {
                new TypedValue((int)DxfCode.Operator,"<or"),
                new TypedValue((int)DxfCode.Start,"lwpolyline"),
                new TypedValue((int)DxfCode.Start,"text"),
                new TypedValue((int)DxfCode.Operator,"or>"),
            };
                SelectionFilter filter = new SelectionFilter(typeds);
                PromptSelectionOptions options1 = new PromptSelectionOptions() { MessageForAdding = "请选择柱大样" };
                var result1 = editor.GetSelection(options1, filter);
                if (result1.Status != PromptStatus.OK)
                {
                    return;
                }

                PromptEntityOptions options2 = new PromptEntityOptions("请选择柱轮廓所在图层");
                options2.SetRejectMessage("不是多段线\n");
                options2.AddAllowedClass(typeof(Polyline), true);
                PromptEntityResult result2 = editor.GetEntity(options2);
                if (result2.Status != PromptStatus.OK)
                {
                    return;
                }

                Polyline polyline = document.GetObject(result2.ObjectId) as Polyline;
                List<Polyline> polylines = new List<Polyline>();
                List<DBText> dBTexts = new List<DBText>();
                using (Transaction tr = database.NewTransaction())
                {
                    foreach (var item in result1.Value.GetObjectIds())
                    {
                        DBObject dBObject = item.GetObject(OpenMode.ForRead, false, true);
                        switch (dBObject)
                        {
                            case Polyline p when p.Layer != LayerName.ColumnDetail && p.Layer == polyline.Layer:
                                polylines.Add(p);
                                break;
                            case DBText t when (t.TextString.Contains('Z')):
                                dBTexts.Add(t);
                                break;
                            default:
                                dBObject.Dispose();
                                break;
                        }
                    }
                    tr.Commit();
                }
                if (dBTexts.Count == 0)
                {
                    polylines.DisposeAll();
                    editor.WriteMessage("不是柱大样图\n");
                    return;
                }

                //去掉重复的多段线
                //如果第i个在后面某个内部，则去掉第i个
                for (int i = 0; i < polylines.Count; i++)
                {
                    for (int j = 0; j < polylines.Count; j++)
                    {
                        if (i == j) continue;
                        if (polylines[i].IsInsideByBounds(polylines[j], LightningTolerance.Structural))
                        {
                            polylines[i].Dispose();
                            polylines.RemoveAt(i);
                            i--;
                            break;
                        }
                    }
                }

                //去掉柱大样及已经识别的柱
                for (int i = 0; i < polylines.Count; i++)
                {
                    if (polylines[i].Layer == LayerName.ColumnDetail || ColumnDetails.Exists(cd => cd.OriginalPolylineId == polylines[i].ObjectId))
                    {
                        polylines[i].Dispose();
                        polylines.RemoveAt(i);
                        i--;
                    }
                }

                //创建柱大样模型
                foreach (var item in polylines)
                {
                    DBText dBText = item.GetNearest(dBTexts) as DBText;
                    DRColumnModel model = new DRColumnModel(item, dBText);
                    ColumnDetails.Add(model);
                }

                //绘制柱大样
                document.UnLockLayer(LayerName.ColumnDetail);
                foreach (var item in ColumnDetails.Where(c => c.IsValid))
                {
                    item.Drawing(LayerName.ColumnDetail, ColorEnum.Yellow);
                }
                document.LockLayer(LayerName.ColumnDetail);

                //清理数据
                polylines.DisposeAll();
                dBTexts.DisposeAll();

                detailDG.DataContext = null;
                detailDG.DataContext = ColumnDetails.Where(cd => cd.IsValid);
                var invalids = ColumnDetails.Where(cd => !cd.IsValid).Select(cd => cd.Name).ToList();
                string str = string.Join("，", invalids);
                editor.WriteMessage($"大样识别共 {ColumnDetails.Count} 个，有 {invalids.Count} 个识别无效：\n{str}\n");

            }
            catch (Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
        private void DetailModify_Click(object sender, RoutedEventArgs e)
        {
            if (Columns.Count == 0)
            {
                editor.WriteMessage("请先识别柱大样\n");
                return;
            }

            Polyline polyline;
            PromptEntityOptions options1 = new PromptEntityOptions("请选择要修改的柱大样");
            options1.SetRejectMessage("不是柱大样\n");
            options1.AddAllowedClass(typeof(Polyline), true);
            options1.AllowObjectOnLockedLayer = true;
            while (true)
            {
                PromptEntityResult result = editor.GetEntity(options1);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                polyline = document.GetObject(result.ObjectId) as Polyline;
                if (polyline.Layer != LayerName.ColumnDetail && polyline.Layer != LayerName.ColumnDetail)
                {
                    polyline.Dispose();
                    editor.WriteMessage("不是柱大样\n");
                }
                else
                {
                    break;
                }
            }

            //修改柱大样编号
            DRColumnModel column = ColumnDetails.Find(c => c.IsValid && c.ComponentObjectId == polyline.ObjectId);
            editor.WriteMessage("柱大样 " + column.Name + "\n");

            DBText dBText;
            PromptEntityOptions options2 = new PromptEntityOptions("请选择修改后的柱大样编号");
            options2.SetRejectMessage("不是柱大样编号\n");
            options2.AddAllowedClass(typeof(DBText), true);
            options2.AllowObjectOnLockedLayer = true;
            while (true)
            {
                PromptEntityResult result = editor.GetEntity(options2);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                dBText = document.GetObject(result.ObjectId) as DBText;
                if (dBText.TextString.Contains('Z'))
                {
                    break;
                }
                dBText.Dispose();
                editor.WriteMessage("不是柱大样编号\n");
            }

            if (column != null)
            {
                using (DocumentLock @lock = document.LockDocument())
                {
                    document.UnLockLayer(LayerName.ColumnDetail);
                    column.UpdateName(dBText);
                    document.LockLayer(LayerName.ColumnDetail);
                }
            }
            editor.WriteMessage("改为 " + column.Name + "\n");
            polyline.Dispose();
            dBText.Dispose();
        }
        private void DetailDelete_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (ColumnDetails.Count == 0)
            {
                editor.WriteMessage("请先识别柱大样\n");
                return;
            }

            PromptSelectionOptions prompt = new PromptSelectionOptions() { MessageForAdding = "请选择要删除的柱大样" };
            SelectionFilter filter = new SelectionFilter(new TypedValue[]
            {
                new TypedValue((int)DxfCode.Operator,"<and"),
                new TypedValue((int)DxfCode.Start,"lwpolyline"),
                new TypedValue((int)DxfCode.LayerName,LayerName.ColumnDetail),
                new TypedValue((int )DxfCode.Operator,"and>"),
            });
            PromptSelectionResult result = editor.GetSelection(prompt, filter);
            if (result.Status != PromptStatus.OK)
            {
                return;
            }
            document.UnLockLayer(LayerName.ColumnDetail);
            foreach (var item in result.Value.GetObjectIds())
            {
                DRColumnModel columnModel = ColumnDetails.Find(col => col.IsValid && col.ComponentObjectId == item);
                if (columnModel != null)
                {
                    columnModel.Delete();
                    ColumnDetails.Remove(columnModel);
                }
            }
            document.LockLayer(LayerName.ColumnDetail);
            detailDG.DataContext = null;
            detailDG.DataContext = ColumnDetails.Where(cd => cd.IsValid);
        }


        private void WallInfoRead_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PromptSelectionOptions options = new PromptSelectionOptions()
                {
                    MessageForAdding = "请选择墙表"
                };
                TypedValue[] typedValues = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Start,"text"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result = editor.GetSelection(options, filter);
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                List<DBText> dBTexts = document.GetObjects(result.Value.GetObjectIds()).Cast<DBText>().ToList();
                List<DBText> infos = dBTexts.Where(t => !(t.TextString.Contains("墙身") || t.TextString.Contains('备'))).ToList();

                List<DBText> Qs = infos.Where(t => t.TextString.Contains('Q')).ToList();
                List<DBText> Ts = infos.Where(t => t.TextString.Contains("mm")).ToList();
                for (int i = 0; i < Qs.Count;)
                {
                    double y1 = Qs.Max(t => t.Position.Y);
                    DBText Q = Qs.First(t => t.Position.Y == y1);
                    Qs.Remove(Q);
                    double y2 = Ts.Max(t => t.Position.Y);
                    DBText T = Ts.First(t => t.Position.Y == y2);
                    Ts.Remove(T);
                    if (int.TryParse(T.TextString.Replace("m", ""), out int th))
                    {
                        DRWallInfo wallInfo = new DRWallInfo()
                        {
                            Name = Q.TextString,
                            Thickness = th,
                        };
                        WallInfos.Add(wallInfo);
                    }
                    Q.Dispose();
                    T.Dispose();
                }
                wallInfoDG.DataContext = null;
                wallInfoDG.DataContext = WallInfos;
            }
            catch (Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
        private void WallInfoAdd_Click(object sender, RoutedEventArgs e)
        {
            WallInfos.Add(new DRWallInfo() { Name = "Q" + (WallInfos.Count + 1).ToString(), Thickness = 0 });
            wallInfoDG.DataContext = null;
            wallInfoDG.DataContext = WallInfos;
        }
        private void WallInfoDelete_Click(object sender, RoutedEventArgs e)
        {
            if (wallInfoDG.SelectedItems.Count != 0)
            {
                foreach (var item in wallInfoDG.SelectedItems)
                {
                    WallInfos.Remove(item as DRWallInfo);
                }
                wallInfoDG.DataContext = null;
                wallInfoDG.DataContext = WallInfos;
            }
        }
    }
}