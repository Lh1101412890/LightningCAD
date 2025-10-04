using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.PlottingServices;

using Lightning.Extension;
using Lightning.Information;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;
using LightningCAD.Models;

using Button = System.Windows.Controls.Button;
using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Exception = System.Exception;
using Label = System.Windows.Controls.Label;
using MessageBox = System.Windows.MessageBox;
using Polyline = Autodesk.AutoCAD.DatabaseServices.Polyline;

namespace LightningCAD.Views
{
    /// <summary>
    /// BatchPrintView.xaml 的交互逻辑
    /// </summary>
    public partial class LBPrintView : ViewBase
    {
        private static List<PrintInfo> PrintInfos;
        private LBPrintInfoView InfoView;
        private static string XmlFile => PCInfo.MyDocumentsDirectoryInfo + "\\Lightning\\LightningCAD\\Prints.xml";
        public LBPrintView()
        {
            InitializeComponent();
            Loaded += BatchPrintView_Loaded;
            Closed += BatchPrintView_Closed;
        }

        /// <summary>
        /// 窗口加载时初始化打印信息、界面元素和历史记录
        /// </summary>
        private void BatchPrintView_Loaded(object sender, RoutedEventArgs e)
        {
            up.Source = Information.GetFileInfo("Commands\\向上箭头.png").ToBitmapImage();
            down.Source = Information.GetFileInfo("Commands\\向下箭头.png").ToBitmapImage();
            PrintInfos = new List<PrintInfo>();
            string name = document.Name.Split('\\').Last();
            fileName.Text = name.Substring(0, name.Length - 4);
            if (File.Exists(XmlFile))
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(XmlFile);
                var info = ReadXml(xmlDocument).Where(p => p.Path == document.Name);
                if (info.Any())
                {
                    PrintInfos = info.First().PrintInfos;
                    dataGrid.DataContext = null;
                    dataGrid.DataContext = PrintInfos;
                    foreach (var item in PrintInfos)
                    {
                        item.CreatFlash();
                    }
                }
            }
        }

        /// <summary>
        /// 窗口关闭时保存打印信息到XML，并清理资源
        /// </summary>
        private void BatchPrintView_Closed(object sender, EventArgs e)
        {
            if (InfoView != null)
            {
                Dispatcher.Invoke(InfoView.Close);
            }
            XmlDocument xmlDocument = new XmlDocument();
            string file = File.Exists(XmlFile) ? XmlFile : Information.GetFileInfo("Files\\Prints.xml").FullName;
            xmlDocument.Load(file);
            XmlElement root = xmlDocument.DocumentElement;
            List<PrintRecording> prints = ReadXml(xmlDocument);
            IEnumerable<PrintRecording> info = prints.Where(p => p.Path == document.Name);
            //更新或者添加当前文档打印信息
            if (info.Any())
            {
                info.First().PrintInfos = PrintInfos;
            }
            else
            {
                prints.Insert(0, new PrintRecording() { Path = document.Name, PrintInfos = PrintInfos });
            }
            SaveXml(prints);
            foreach (var item in PrintInfos)
            {
                item.DeleteFlash();
            }
            PrintInfos.Clear();
            numbers?.Clear();
        }

        /// <summary>
        /// 从XML文档读取所有打印记录
        /// </summary>
        private static List<PrintRecording> ReadXml(XmlDocument document)
        {
            XmlElement root = document.DocumentElement;
            //读取已保存的打印信息
            List<PrintRecording> prints = new List<PrintRecording>();
            foreach (XmlNode print in root.ChildNodes)
            {
                //打印信息
                List<PrintInfo> models = new List<PrintInfo>();
                foreach (XmlNode printModel in print.LastChild)
                {
                    int i = int.Parse(printModel["Number"].InnerText);
                    string[] strings = printModel["Extents2d"].InnerText.Split(',');
                    double minx = double.Parse(strings[0]);
                    double miny = double.Parse(strings[1]);
                    double maxx = double.Parse(strings[2]);
                    double maxy = double.Parse(strings[3]);
                    Extents2d extents = new Extents2d(minx, miny, maxx, maxy);
                    PrintInfo model = new PrintInfo(i, extents)
                    {
                        Size = (PaperSize)Enum.Parse(typeof(PaperSize), printModel["Size"].InnerText),
                        Direction = (PrintDirection)Enum.Parse(typeof(PrintDirection), printModel["Direction"].InnerText),
                        LineWeight = bool.Parse(printModel["LineWeight"].InnerText),
                        Transparent = bool.Parse(printModel["Transparent"].InnerText),
                        Style = (PrintStyle)Enum.Parse(typeof(PrintStyle), printModel["Style"].InnerText),
                        Note = printModel["Note"].InnerText
                    };
                    models.Add(model);
                }
                prints.Add(new PrintRecording() { Path = print.FirstChild.InnerText, PrintInfos = models });
            }
            return prints;
        }

        /// <summary>
        /// 保存打印记录到XML文件
        /// </summary>
        private void SaveXml(List<PrintRecording> prints)
        {
            //保存信息至本地
            XmlDocument xml = new XmlDocument();
            xml.Load(Information.GetFileInfo("Files\\Prints.xml").FullName);
            XmlElement xmlRoot = xml.DocumentElement;
            //解析数据至内存
            foreach (var model in prints)
            {
                if (model.PrintInfos.Count == 0) continue;

                XmlNode node = xml.CreateElement("Print");

                XmlNode fileFullName = xml.CreateElement(nameof(PrintRecording.Path));
                fileFullName.InnerText = model.Path;
                node.AppendChild(fileFullName);

                XmlNode printModels = xml.CreateElement(nameof(PrintRecording.PrintInfos));
                node.AppendChild(printModels);
                foreach (var item in model.PrintInfos)
                {
                    XmlNode printModel = xml.CreateElement(nameof(PrintInfo));
                    printModels.AppendChild(printModel);

                    XmlNode number = xml.CreateElement(nameof(PrintInfo.Number));
                    number.InnerText = item.Number.ToString();
                    printModel.AppendChild(number);

                    XmlNode size = xml.CreateElement(nameof(PrintInfo.Size));
                    size.InnerText = item.Size.ToString();
                    printModel.AppendChild(size);

                    XmlNode direction = xml.CreateElement(nameof(PrintInfo.Direction));
                    direction.InnerText = item.Direction.ToString();
                    printModel.AppendChild(direction);

                    XmlNode lineWeight = xml.CreateElement(nameof(PrintInfo.LineWeight));
                    lineWeight.InnerText = item.LineWeight.ToString();
                    printModel.AppendChild(lineWeight);

                    XmlNode transparent = xml.CreateElement(nameof(PrintInfo.Transparent));
                    transparent.InnerText = item.Transparent.ToString();
                    printModel.AppendChild(transparent);

                    XmlNode style = xml.CreateElement(nameof(PrintInfo.Style));
                    style.InnerText = item.Style.ToString();
                    printModel.AppendChild(style);

                    XmlNode extents2d = xml.CreateElement(nameof(PrintInfo.Extents2d));
                    extents2d.InnerText = item.Extents2d.MinPoint.X + "," + item.Extents2d.MinPoint.Y + "," + item.Extents2d.MaxPoint.X + "," + item.Extents2d.MaxPoint.Y;
                    printModel.AppendChild(extents2d);

                    XmlNode note = xml.CreateElement(nameof(PrintInfo.Note));
                    note.InnerText = item.Note ?? "";
                    printModel.AppendChild(note);
                }
                xmlRoot.AppendChild(node);
            }
            Directory.CreateDirectory(Path.GetDirectoryName(XmlFile));
            xml.Save(XmlFile);
        }

        /// <summary>
        /// 手动添加打印范围，支持多次框选
        /// </summary>
        private void AddRangeButton_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            while (true)
            {
                PromptPointResult result = editor.GetPoint("请框选打印范围（ESC退出）");
                if (result.Status != PromptStatus.OK)
                {
                    return;
                }
                Point3d point1 = result.Value;

                DragRectangle dragPolyline = new DragRectangle(point1, "");
                PromptResult promptResult = editor.Drag(dragPolyline);
                if (promptResult.Status != PromptStatus.OK)
                {
                    return;
                }
                Point3d point2 = dragPolyline.point2;
                double minX = point1.X < point2.X ? point1.X : point2.X;
                double maxX = point1.X > point2.X ? point1.X : point2.X;
                double minY = point1.Y < point2.Y ? point1.Y : point2.Y;
                double maxY = point1.Y > point2.Y ? point1.Y : point2.Y;
                Extents2d extents2D = new Extents2d(minX, minY, maxX, maxY);//得到范围框
                if (minX == maxX || minY == maxY)
                {
                    editor.WriteMessage("框选范围无效\n");
                    continue;
                }
                PrintInfo model = new PrintInfo(PrintInfos.Count + 1, extents2D);
                model.CreatFlash();
                PrintInfos.Add(model);
                dataGrid.DataContext = null;
                dataGrid.DataContext = PrintInfos;
                dataGrid.ScrollIntoView(model);
                editor.WriteMessage("\n");
            }
        }

        /// <summary>
        /// 自动识别块并批量添加打印范围，支持排序
        /// </summary>
        private void AutoRangeButton_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            //基准块
            PromptSelectionOptions options1 = new PromptSelectionOptions() { MessageForAdding = "请选择基准块", };
            TypedValue[] typeds =
            {
                new TypedValue((int)DxfCode.Start,"INSERT"),//过滤块
            };
            SelectionFilter filter = new SelectionFilter(typeds);
            PromptSelectionResult baseResult = editor.GetSelection(options1, filter);
            if (baseResult.Status != PromptStatus.OK) return;
            List<string> names = new List<string>();
            using (Transaction transaction = database.NewTransaction())
            {
                foreach (ObjectId objectId in baseResult.Value.GetObjectIds())
                {
                    BlockReference block = (BlockReference)objectId.GetObject(OpenMode.ForRead, false, true);
                    if (block.Name.StartsWith("*"))
                    {
                        BlockTableRecord btr = block.DynamicBlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord;
                        names.Add(btr.Name);
                        btr.Dispose();
                    }
                    else
                    {
                        names.Add(block.Name);
                    }
                    block.Dispose();
                }
                transaction.Commit();
            }
            names = names.Distinct().ToList();
            editor.WriteMessage($"共找到 {names.Count} 个基准块\n");

            //识别块
            PromptSelectionOptions options2 = new PromptSelectionOptions() { MessageForAdding = "请选择要识别的范围", };
            PromptSelectionResult blocks = editor.GetSelection(options2, filter);
            if (blocks.Status != PromptStatus.OK) return;
            List<Extents2d> extents = new List<Extents2d>();
            using (DocumentLock @lock = document.LockDocument())
            {
                using (Transaction transaction = database.NewTransaction())
                {
                    foreach (ObjectId objectId in blocks.Value.GetObjectIds())
                    {
                        BlockReference block = (BlockReference)objectId.GetObject(OpenMode.ForRead, false, true);
                        if (block.Bounds != null && block.Bounds.Value.GetHeight() != 0 && block.Bounds.Value.GetWidth() != 0)
                        {
                            if (block.Name.StartsWith("*"))
                            {
                                BlockTableRecord btr = block.DynamicBlockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord;
                                if (names.Exists(n => n == btr.Name))
                                {
                                    extents.Add(block.Bounds.Value.ToExtents2d());
                                }
                                btr.Dispose();
                            }
                            else if (names.Exists(n => n == block.Name))
                            {
                                extents.Add(block.Bounds.Value.ToExtents2d());
                            }
                        }
                        block.Dispose();
                    }
                    transaction.Commit();
                }
            }
            if (extents.Count == 0)
            {
                editor.WriteMessage("未识别到有效块!\n");
                return;
            }
            editor.WriteMessage($"共找到 {blocks.Value.Count} 个\n");

            //排序
            PromptKeywordOptions options = new PromptKeywordOptions("请选择排序方式[从左至右(Z)/从上至下(W)/放弃(U)]")
            {
                AppendKeywordsToMessage = true,
            };
            options.Keywords.Add("Z");
            options.Keywords.Add("W");
            options.Keywords.Add("U");
            options.Keywords.Default = "Z";
            PromptResult result = editor.GetKeywords(options);
            if (result.Status != PromptStatus.OK || result.StringResult == "U") return;
            List<Extents2d> newExtents2ds = new List<Extents2d>();
            while (true)
            {
                if (extents.Count == 0) break;
                List<Extents2d> temp;
                double acc = 10;
                if (result.StringResult == "Z")//以左上角为开始位置，从上往下逐行
                {
                    double maxY = extents.Max(ext => ext.GetLeftTop().Y);
                    temp = extents.Where(ext => Math.Abs(ext.GetLeftTop().Y - maxY) < acc).ToList();
                    temp = temp.OrderBy(ex => ex.GetLeftTop().X).ToList();
                }
                else//以左上角为开始位置，从左往右逐列
                {
                    double minX = extents.Min(ext => ext.GetLeftTop().X);
                    temp = extents.Where(ext => Math.Abs(ext.GetLeftTop().X - minX) < acc).ToList();
                    temp = temp.OrderByDescending(ex => ex.GetLeftTop().Y).ToList();
                }
                newExtents2ds.AddRange(temp);
                foreach (var item in temp)
                {
                    extents.Remove(item);
                }
            }

            //创建打印模型
            for (int i = 0; i < newExtents2ds.Count; i++)
            {
                PrintInfo model = new PrintInfo(PrintInfos.Count + 1, newExtents2ds[i]);
                model.CreatFlash();
                PrintInfos.Add(model);
            }

            dataGrid.DataContext = null;
            dataGrid.DataContext = PrintInfos;
        }

        /// <summary>
        /// 重新设置选中项的打印范围
        /// </summary>
        private void ResetRangeButton_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            List<PrintInfo> choices = dataGrid.SelectedItems.Cast<PrintInfo>().OrderBy(p => p.Number).ToList();//升序后的选择项
            if (choices.Count == 0)
            {
                return;
            }
            foreach (PrintInfo printInfo in choices)
            {
                Extents2d extents2D;
                while (true)
                {
                    //重新框选打印范围
                    PromptPointResult promptPointResult = editor.GetPoint($"请重新框选【{printInfo.Number}】的打印范围");
                    if (promptPointResult.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    Point3d point1 = promptPointResult.Value;

                    DragRectangle dragRectangle = new DragRectangle(point1, "指定另一角点");
                    DragRectangle dragPolyline = dragRectangle;
                    PromptResult promptResult = editor.Drag(dragPolyline);
                    if (promptResult.Status != PromptStatus.OK)
                    {
                        return;
                    }
                    Point3d point2 = dragPolyline.point2;

                    double minX = point1.X < point2.X ? point1.X : point2.X;
                    double maxX = point1.X > point2.X ? point1.X : point2.X;
                    double minY = point1.Y < point2.Y ? point1.Y : point2.Y;
                    double maxY = point1.Y > point2.Y ? point1.Y : point2.Y;
                    if (minX != maxX && minY != maxY)
                    {
                        extents2D = new Extents2d(minX, minY, maxX, maxY);//得到范围框
                        break;
                    }
                    else
                    {
                        editor.WriteMessage("框选范围无效\n");
                    }
                }
                editor.WriteMessage("\n");

                printInfo.Extents2d = extents2D;
                printInfo.UpdateFlash();
            }
        }

        /// <summary>
        /// 删除选中的打印范围
        /// </summary>
        private void DeleteRangeButton_Click(object sender, RoutedEventArgs e)
        {
            if (dataGrid.SelectedItems.Count == 0) return;
            MessageBoxResult box = MessageBox.Show("确定删除记录吗？", "删除确认", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (box != MessageBoxResult.OK) return;
            CADApp.MainWindow.Handle.Focus();
            List<PrintInfo> models = dataGrid.SelectedItems.Cast<PrintInfo>().ToList();
            foreach (PrintInfo model in models)
            {
                model.DeleteFlash();
                PrintInfos.Remove(model);
            }
            for (int i = 0; i < PrintInfos.Count; i++)
            {
                PrintInfos[i].Number = i + 1;
                PrintInfos[i].UpdateFlash();
            }
            dataGrid.DataContext = null;
            dataGrid.DataContext = PrintInfos;
        }

        /// <summary>
        /// 清空所有打印范围
        /// </summary>
        private void ClearRangeButton_Click(object sender, RoutedEventArgs e)
        {
            if (dataGrid.Items.Count == 0) return;
            MessageBoxResult box = MessageBox.Show("确定清空记录吗？", "删除确认", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (box != MessageBoxResult.OK) return;
            foreach (PrintInfo model in PrintInfos)
            {
                model.DeleteFlash();
            }
            PrintInfos.Clear();
            CADApp.MainWindow.Handle.Focus();
            dataGrid.DataContext = null;
            dataGrid.DataContext = PrintInfos;
        }

        /// <summary>
        /// 上移或下移选中的打印范围编号，实现批量调整顺序
        /// </summary>
        private void MoveNumberButton_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            int selected = dataGrid.SelectedItems.Count;
            if (selected == 0) return;//没选择返回
            int min = dataGrid.SelectedItems.Cast<PrintInfo>().Min(p => p.Number);//最小的编号
            int max = dataGrid.SelectedItems.Cast<PrintInfo>().Max(p => p.Number); //最大的编号
            if (max - min + 1 != selected) return;//非连续选择返回
            int n;//确保dataGrid选中的行离显示范围边缘有一定的距离，否则会bug
            if ((sender as Button).Name == "uper")
            {
                if (min == 1) return;//选择的是第一个无法上移

                PrintInfos[min - 2].Number = max;//第一个选择的上一个的编号等于最后一个选择的编号
                for (int i = 0; i < selected; i++)//所有选择的编号自减1
                {
                    PrintInfos[min - 1 + i].Number -= 1;
                }
                n = min - 7;//距离上边缘有5行
                if (n < 0) n = 0;
                min -= 2;//起始数字修正
                max -= 1;//起始数字修正
            }
            else
            {
                if (max == PrintInfos.Count) return;//选择的是第一个无法下移

                PrintInfos[max].Number = min;//最后一个选择的下一个的编号等于第一个选择的编号
                for (int i = 0; i < selected; i++)//所有选择的编号自增1
                {
                    PrintInfos[min - 1 + i].Number += 1;
                }
                n = max + 5;//距离下边缘有5行
                if (n > PrintInfos.Count - 1) n = PrintInfos.Count - 1;
                max += 1;//起始数字修正
            }
            PrintInfos = PrintInfos.OrderBy(p => p.Number).Cast<PrintInfo>().ToList();//按顺序重新排序
            foreach (PrintInfo info in PrintInfos)
            {
                info.UpdateFlash();
            }
            dataGrid.DataContext = null;
            dataGrid.DataContext = PrintInfos;
            dataGrid.UpdateLayout();
            List<DataGridRow> rows = new List<DataGridRow>();
            for (int i = min; i < max; i++)
            {
                DataGridRow row = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(PrintInfos[i].Number - 1);
                rows.Add(row);
            }
            foreach (DataGridRow row in rows)
            {
                row.IsSelected = true;
            }
            dataGrid.ScrollIntoView(PrintInfos[n]);//向上或下移就最外围一行显示留出5行边距，除非小于5
            dataGrid.Focus();
        }

        /// <summary>
        /// 双击打印范围，视图自动缩放并定位到该范围
        /// </summary>
        private void DataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (!((sender as DataGrid).SelectedItem is PrintInfo printModel)) return;//printModel为null时返回
            Point2d max = printModel.Extents2d.MaxPoint;
            Point2d min = printModel.Extents2d.MinPoint;
            using (ViewTableRecord view = editor.GetCurrentView())
            {
                double n = 2.5;
                view.Height = (max.Y - min.Y) * n;
                view.Width = (max.X - min.X) * n;
                view.CenterPoint = new Point2d((max.X + min.X) / 2, (max.Y + min.Y) / 2);
                editor.SetCurrentView(view);
            }
        }

        /// <summary>
        /// 右键菜单修改打印参数（纸张、方向、样式等）
        /// </summary>
        private void ContextMenu_Click(object sender, RoutedEventArgs e)
        {
            MenuItem menuItem = e.OriginalSource as MenuItem;
            Label label = menuItem.Header as Label;
            CADApp.MainWindow.Handle.Focus();
            foreach (var item in dataGrid.SelectedItems)
            {
                PrintInfo printModel = item as PrintInfo;
                switch (label.Content.ToString())
                {
                    case "A4":
                        printModel.Size = PaperSize.A4;
                        break;
                    case "A3":
                        printModel.Size = PaperSize.A3;
                        break;
                    case "A2":
                        printModel.Size = PaperSize.A2;
                        break;
                    case "A1":
                        printModel.Size = PaperSize.A1;
                        break;
                    case "A0":
                        printModel.Size = PaperSize.A0;
                        break;
                    case "纵向":
                        printModel.Direction = PrintDirection.纵向;
                        break;
                    case "横向":
                        printModel.Direction = PrintDirection.横向;
                        break;
                    case "黑白":
                        printModel.Style = PrintStyle.黑白;
                        break;
                    case "彩色":
                        printModel.Style = PrintStyle.彩色;
                        break;
                    case "true":
                        if ((label.Parent as MenuItem).Header.ToString() == "线宽")
                        {
                            printModel.LineWeight = true;
                        }
                        else
                        {
                            printModel.Transparent = true;
                        }
                        break;
                    case "false":
                        if ((label.Parent as MenuItem).Header.ToString() == "线宽")
                        {
                            printModel.LineWeight = false;
                        }
                        else
                        {
                            printModel.Transparent = false;
                        }
                        break;
                    default:
                        break;
                }
                printModel.UpdateFlash();
            }
            dataGrid.DataContext = null;
            dataGrid.DataContext = PrintInfos;
        }

        private List<int> numbers;

        private static bool IsValid(string str)
        {
            return !str.Contains('\\')
                && !str.Contains('/')
                && !str.Contains(':')
                && !str.Contains('*')
                && !str.Contains('?')
                && !str.Contains('"')
                && !str.Contains('<')
                && !str.Contains('>')
                && !str.Contains('|');
        }

        /// <summary>
        /// 执行批量打印，支持自定义范围和合并PDF
        /// </summary>
        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (PrintInfos.Count == 0) return;
                string myDocument = PCInfo.MyDocumentsDirectoryInfo.FullName;
                //如果未设置名字就赋予默认名字
                string name = string.IsNullOrEmpty(fileName.Text) ? "Lightning批量打印" : fileName.Text;
                if (!IsValid(name))
                {
                    MessageBox.Show("文件名不能包含\\/:*?\"<>|这些特殊字符！", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                DirectoryInfo directory = new DirectoryInfo($"{myDocument}\\Lightning\\LightningCAD\\批量打印");
                if (!directory.Exists)
                {
                    directory.Create();
                }
                List<string> files = new List<string>();
                numbers = new List<int>();
                if (custom.IsChecked == true)
                {
                    if (string.IsNullOrEmpty(range.Text))
                    {
                        return;
                    }
                    string text = range.Text.Replace(" ", "").Replace("，", ",");
                    char[] chars = text.ToCharArray();
                    foreach (char c in chars)
                    {
                        if ((c < 48 || c > 57) && c != ',' && c != '-')
                        {
                            MessageBox.Show("自定义范围有误！");
                        }
                    }
                    string[] strings = text.Split(',');
                    foreach (string s in strings)
                    {
                        if (s.Contains('-'))
                        {
                            int one = int.Parse(s.Split('-').First());
                            int two = int.Parse(s.Split('-').Last());
                            int start = Math.Min(one, two);
                            int end = Math.Max(one, two);
                            for (int i = start; i <= end; i++)
                            {
                                numbers.Add(i);
                            }
                        }
                        else
                        {
                            numbers.Add(int.Parse(s));
                        }
                    }
                    for (int i = 0; i < numbers.Count; i++)
                    {
                        int n = numbers[i] - 1;
                        string file = string.IsNullOrWhiteSpace(PrintInfos[n].Note)
                            ? $"{directory.FullName}\\{name}_{n + 1}.pdf"
                            : $"{directory.FullName}\\{PrintInfos[n].Note}.pdf";
                        Print(PrintInfos[n], file, $"LightningCAD批量打印 {n + 1}/{numbers.Count}");
                        files.Add(file);
                    }
                }
                else
                {
                    for (int n = 0; n < PrintInfos.Count; n++)
                    {
                        string file = string.IsNullOrWhiteSpace(PrintInfos[n].Note)
                            ? $"{directory.FullName}\\{name}_{n + 1}.pdf"
                            : $"{directory.FullName}\\{PrintInfos[n].Note}.pdf";
                        Print(PrintInfos[n], file, $"LightningCAD批量打印 {n + 1}/{PrintInfos.Count}");
                        files.Add(file);
                    }
                }
                Close();
                //合并为单个pdf文件
                if (single.SelectedIndex == 0)
                {
                    string file = $"{directory.FullName}\\{name}.pdf";
                    Tools.JoinFiles(files, file);
                    foreach (string item in files)
                    {
                        File.Delete(item);
                    }
                    Process.Start("Explorer", "/select," + file);
                }
                else
                {
                    Process.Start("Explorer", directory.FullName);
                }
            }
            catch (Exception exp)
            {
                exp.Record();
            }
        }

        #region 更多按钮，打印记录管理

        /// <summary>
        /// 更多按钮，弹出打印记录管理窗口
        /// </summary>
        private void MoreButton_Click(object sender, RoutedEventArgs e)
        {
            if (InfoView != null) return;
            InfoView = new LBPrintInfoView();
            InfoView.copy.Click += CopyRange_Click;
            InfoView.check.Click += ClearInvalid_Click;
            InfoView.delete.Click += DeleteRecord_Click;
            InfoView.clear.Click += ClearRecord_Click;
            InfoView.dataGrid.MouseDoubleClick += (s, ev) =>//打开记录对应的文件目录
            {
                if (InfoView.dataGrid.SelectedItems.Count != 1)
                {
                    return;
                }
                FileInfo fileInfo = new FileInfo(InfoView.dataGrid.SelectedItem.ToString());
                if (fileInfo.Exists)
                {
                    Dispatcher.Invoke(() => InfoView.Close());
                    InfoView = null;
                    CADApp.DocumentManager.MdiActiveDocument = CADApp.DocumentManager.Open(fileInfo.FullName, false);
                }
                else
                {
                    //询问是否要删除本条记录
                    if (MessageBox.Show("该文件已删除，是否删除本条记录？", "警告", MessageBoxButton.OKCancel) == MessageBoxResult.Cancel) return;

                    List<string> files = InfoView.dataGrid.SelectedItems.Cast<string>().Select(i => i.ToString()).ToList();
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(XmlFile);
                    List<PrintRecording> prints = ReadXml(xmlDocument);
                    List<PrintRecording> newPrints = prints.Where(p => !files.Exists(f => f == p.Path)).ToList();
                    SaveXml(newPrints);
                    //如果删除了当前文档的记录
                    if (files.Exists(f => f == document.Name))
                    {
                        CADApp.MainWindow.Handle.Focus();

                        foreach (var item in PrintInfos)
                        {
                            item.DeleteFlash();
                        }
                        PrintInfos.Clear();
                        dataGrid.DataContext = null;
                        dataGrid.DataContext = PrintInfos;
                    }
                    InfoView.RaiseEvent(new RoutedEventArgs(LoadedEvent));
                    ShowInfos();
                }
            };
            InfoView.Loaded += (s, ev) =>//加载打印记录
            {
                if (!File.Exists(XmlFile)) return;
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(XmlFile);
                List<PrintRecording> prints = ReadXml(xmlDocument);
                InfoView.dataGrid.DataContext = null;
                InfoView.dataGrid.DataContext = prints.Select(p => p.Path).ToList();
            };
            InfoView.Closed += (s, ev) => { InfoView = null; };
            ShowInfos();
            InfoView.Show();
        }

        /// <summary>
        /// 复制选中记录的打印范围到当前文档
        /// </summary>
        private void CopyRange_Click(object sender, RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            if (InfoView.dataGrid.SelectedItems.Count != 1 || !File.Exists(XmlFile)) return;
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(XmlFile);
            List<PrintRecording> prints = ReadXml(xmlDocument);
            foreach (var item in PrintInfos)
            {
                item.DeleteFlash();
            }
            PrintInfos = prints.Where(p => p.Path == InfoView.dataGrid.SelectedItem.ToString()).First().PrintInfos;//读取对应的打印数据到PrintModels
            Dispatcher.Invoke(() => { InfoView.Close(); });
            InfoView = null;
            dataGrid.DataContext = null;
            dataGrid.DataContext = PrintInfos;
            foreach (var item in PrintInfos)
            {
                item.CreatFlash();
            }

        }

        /// <summary>
        /// 清理无效（文件已删除）的打印记录
        /// </summary>
        private void ClearInvalid_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(XmlFile))
            {
                return;
            }
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(XmlFile);
            List<PrintRecording> prints = ReadXml(xmlDocument);
            List<PrintRecording> newPrints = prints.Where(p => File.Exists(p.Path)).ToList();
            SaveXml(newPrints);
            InfoView.RaiseEvent(new RoutedEventArgs(LoadedEvent));
            ShowInfos();
        }

        /// <summary>
        /// 删除选中的打印记录
        /// </summary>
        private void DeleteRecord_Click(object sender, RoutedEventArgs e)
        {
            if (InfoView.dataGrid.SelectedItems.Count == 0 || !File.Exists(XmlFile)) return;
            //询问是否要删除选择的记录
            if (MessageBox.Show("确定要删除选择记录吗？", "警告", MessageBoxButton.OKCancel) == MessageBoxResult.Cancel) return;
            List<string> files = InfoView.dataGrid.SelectedItems.Cast<string>().Select(i => i.ToString()).ToList();

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(XmlFile);
            List<PrintRecording> prints = ReadXml(xmlDocument);

            List<PrintRecording> newPrints = prints.Where(p => !files.Exists(f => f == p.Path)).ToList();
            SaveXml(newPrints);

            //如果删除了当前文档的记录
            if (files.Exists(f => f == document.Name))
            {
                CADApp.MainWindow.Handle.Focus();

                foreach (var item in PrintInfos)
                {
                    item.DeleteFlash();
                }
                PrintInfos.Clear();
                dataGrid.DataContext = null;
                dataGrid.DataContext = PrintInfos;
            }
            InfoView.RaiseEvent(new RoutedEventArgs(LoadedEvent));
            ShowInfos();
        }

        /// <summary>
        /// 清空所有打印记录
        /// </summary>
        private void ClearRecord_Click(object sender, RoutedEventArgs e)
        {
            if (InfoView.dataGrid.Items.Count == 0)
            {
                return;
            }
            //询问是否要删除所有记录
            if (MessageBox.Show("确定要清空所有记录吗？", "警告", MessageBoxButton.OKCancel) == MessageBoxResult.Cancel) return;
            CADApp.MainWindow.Handle.Focus();
            //删除本地保存数据
            if (File.Exists(XmlFile))
            {
                File.Delete(XmlFile);
            }

            foreach (var item in PrintInfos)
            {
                item.DeleteFlash();
            }
            //清空内存数据
            PrintInfos.Clear();
            dataGrid.DataContext = null;
            dataGrid.DataContext = PrintInfos;

            Dispatcher.Invoke(() => { InfoView.Close(); });
            InfoView = null;
        }

        /// <summary>
        /// 显示打印记录信息（数量、占用空间等）
        /// </summary>
        private void ShowInfos()
        {
            FileInfo fileInfo = new FileInfo(XmlFile);
            if (fileInfo.Exists)
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(XmlFile);
                List<PrintRecording> prints = ReadXml(xmlDocument);
                InfoView.Title = $"批量打印记录（{prints.Count}条）";
                long length = fileInfo.Length;
                string v = length <= 1024 * 1024 ? Math.Round(length / 1024.0, 2) + "K" : Math.Round(length / 1024.0 / 1024, 2) + "M";
                InfoView.fileSize.Content = $"记录占用{v}";
            }
            else
            {
                InfoView.fileSize.Content = "无记录";
                InfoView.Title = "批量打印记录（0条）";
            }
        }
        #endregion

        /// <summary>
        /// 执行打印任务，生成PDF文件
        /// </summary>
        /// <param name="model">打印信息</param>
        /// <param name="file">输出PDF路径</param>
        /// <param name="prog">进度描述</param>
        public void Print(PrintInfo model, string file, string prog)
        {
            try
            {
                if (model == null) return;
                using (DocumentLock @lock = document.LockDocument())
                {
                    using (Transaction transaction = database.NewTransaction())
                    {
                        BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(database.CurrentSpaceId, OpenMode.ForRead);
                        Layout layout = (Layout)transaction.GetObject(btr.LayoutId, OpenMode.ForRead);
                        CADApp.SetSystemVariable("BACKGROUNDPLOT", 0);//前台打印
                                                                      //创建一个打印设置对象，再进行自定义设置
                        PlotSettings settings = new PlotSettings(layout.ModelType);
                        settings.CopyFrom(layout);

                        //修改打印设置对象
                        using (PlotSettingsValidator validator = PlotSettingsValidator.Current)
                        {
                            string mediaName;
                            switch (model.Size)
                            {
                                case PaperSize.A0:
                                    mediaName = "ISO_full_bleed_A0_(841.00_x_1189.00_MM)";
                                    break;
                                case PaperSize.A1:
                                    mediaName = "ISO_full_bleed_A1_(594.00_x_841.00_MM)";
                                    break;
                                case PaperSize.A2:
                                    mediaName = "ISO_full_bleed_A2_(420.00_x_594.00_MM)";
                                    break;
                                case PaperSize.A3:
                                    mediaName = "ISO_full_bleed_A3_(297.00_x_420.00_MM)";
                                    break;
                                case PaperSize.A4:
                                default:
                                    mediaName = "ISO_full_bleed_A4_(210.00_x_297.00_MM)";
                                    break;
                            }
                            settings.ShowPlotStyles = true;
                            settings.PrintLineweights = model.LineWeight;//是否打印线宽
                            settings.PlotTransparency = model.Transparent;//是否透明度打印
                            settings.PlotPlotStyles = true;//是否按样式打印

                            validator.SetPlotWindowArea(settings, model.Extents2d);//设置打印范围
                            validator.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);//设置打印类型
                            validator.SetStdScaleType(settings, StdScaleType.ScaleToFit);//布满图纸
                            validator.SetPlotCentered(settings, true);//居中打印
                            validator.SetPlotConfigurationName(settings, "DWG To PDF.pc3", mediaName);//打印机名称及尺寸，宽x高
                            switch (model.Direction)
                            {
                                case PrintDirection.横向:
                                    validator.SetPlotRotation(settings, PlotRotation.Degrees090);//打印方向
                                    break;
                                default:
                                    validator.SetPlotRotation(settings, PlotRotation.Degrees000);//打印方向
                                    break;
                            }
                            validator.GetPlotStyleSheetList();//先读取一下样式表，否则会异常
                            switch (model.Style)
                            {
                                case PrintStyle.彩色:
                                    validator.SetCurrentStyleSheet(settings, "acad.ctb");
                                    break;
                                default:
                                    validator.SetCurrentStyleSheet(settings, "monochrome.ctb");
                                    break;
                            }
                        }

                        //打印信息对象验证器
                        PlotInfoValidator infoValidator = new PlotInfoValidator()
                        {
                            MediaMatchingPolicy = MatchingPolicy.MatchEnabled
                        };
                        //打印信息对象
                        PlotInfo plotInfo = new PlotInfo()
                        {
                            Layout = btr.LayoutId,
                            OverrideSettings = settings,
                        };
                        infoValidator.Validate(plotInfo);//验证打印信息对象有效性
                        settings.Dispose();
                        // PlotEngine对象执行真正的打印工作
                        // (也可以创建一个用来预览)
                        //确定打印工厂的状态
                        if (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
                        {
                            editor.WriteMessage("另一个打印正在进行中！\n");
                            return;
                        }
                        //创建打印工程
                        using (PlotEngine plotEngine = PlotFactory.CreatePublishEngine())
                        {
                            //创建一个打印对话框，用于提供打印信息和允许用户取消打印
                            using (PlotProgressDialog dialog = new PlotProgressDialog(false, 1, false))//表数量
                            {
                                dialog.set_PlotMsgString(PlotMessageIndex.DialogTitle, prog);
                                dialog.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "当前进度");
                                dialog.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "总进度");
                                dialog.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "取消打印");

                                dialog.LowerPlotProgressRange = 0;
                                dialog.UpperPlotProgressRange = 100;
                                dialog.PlotProgressPos = 0;

                                //打印对话框启动
                                dialog.OnBeginPlot();
                                dialog.IsVisible = true;
                                //启动打印
                                plotEngine.BeginPlot(dialog, null);

                                //启动文档：打印信息，文档名称，参数，打印份数，是否打印到文件，打印输出文件名
                                plotEngine.BeginDocument(plotInfo, document.Name, null, 1, true, file);

                                //打印对话框启动表
                                dialog.OnBeginSheet();
                                dialog.LowerSheetProgressRange = 0;
                                dialog.UpperSheetProgressRange = 100;
                                dialog.SheetProgressPos = 0;

                                PlotPageInfo plotPageInfo = new PlotPageInfo();
                                plotEngine.BeginPage(plotPageInfo, plotInfo, true, null);

                                plotEngine.BeginGenerateGraphics(null);
                                plotEngine.EndGenerateGraphics(null);
                                plotEngine.EndPage(null);
                                dialog.PlotProgressPos = 100;

                                //打印对话框结束表
                                dialog.OnEndSheet();
                                //结束文档
                                plotEngine.EndDocument(null);

                                //打印对话框结束
                                dialog.OnEndPlot();
                                //结束打印
                                plotEngine.EndPlot(null);
                                plotEngine.Dispose();
                                plotEngine.Destroy();
                            }
                            // 删掉打印日志文件
                            if (File.Exists(LightningApp.Dir + "\\plot.log"))
                            {
                                File.Delete(LightningApp.Dir + "\\plot.log");
                            }
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                exp.Record();
            }
        }

        /// <summary>
        /// 设置打印范围框（插入标准块到图纸）
        /// </summary>
        private void BoundingBox_Click(object sender, RoutedEventArgs e)
        {
            string file = Information.GetFileInfo("Files\\LightningCAD图库.dwg").FullName;
            using (Database libDatabase = new Database(false, true))
            {
                ObjectIdCollection objectIdCollection = new ObjectIdCollection();
                try
                {
                    libDatabase.ReadDwgFile(file, FileOpenMode.OpenForReadAndReadShare, true, "");
                    libDatabase.CloseInput(true);
                }
                catch (Exception exp)
                {
                    exp.Record();
                    return;
                }
                using (Transaction tr = libDatabase.NewTransaction())
                {
                    // 模型空间
                    BlockTable blkTbl = tr.GetObject(libDatabase.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord modelSpace = tr.GetObject(blkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                    List<BlockReference> blockReferences = new List<BlockReference>();
                    foreach (ObjectId item in modelSpace)
                    {
                        if (item.GetObject(OpenMode.ForRead) is BlockReference br)
                        {
                            blockReferences.Add(br);
                        }
                    }
                    string name;
                    if ((sender as Button).Name == "verticalBox")
                    {
                        name = "纵向打印框";
                        isVertical = true;
                    }
                    else
                    {
                        name = "横向打印框";
                        isVertical = false;
                    }
                    BlockReference blockReference;
                    try
                    {
                        blockReference = blockReferences.First(id => id.Name == name);
                    }
                    catch (InvalidOperationException exp)
                    {
                        tr.Abort();
                        editor.WriteMessage("未找到打印范围框\n");
                        exp.Record();
                        return;
                    }
                    objectIdCollection.Add(blockReference.ObjectId);
                    tr.Commit();
                }
                using (DocumentLock @lock = document.LockDocument())
                {
                    using (Transaction transaction = database.NewTransaction())
                    {
                        Polyline polyline = GetRec(Point3d.Origin, isVertical);
                        lFlash = new LFlash(polyline, TransientDrawingMode.Highlight);
                        editor.PointMonitor += Editor_PointMonitor;
                        PromptPointResult promptPointResult = editor.GetPoint("指定范围框位置");
                        editor.PointMonitor -= Editor_PointMonitor;
                        lFlash.Delete();
                        if (promptPointResult.Status != PromptStatus.OK)
                        {
                            transaction.Abort();
                            return;
                        }
                        var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                        var blockRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        IdMapping idMapping = new IdMapping();
                        libDatabase.WblockCloneObjects(objectIdCollection, blockRecord.ObjectId, idMapping, DuplicateRecordCloning.Ignore, false);
                        var block = (BlockReference)idMapping[objectIdCollection[0]].Value.GetObject(OpenMode.ForWrite);
                        block.Position = promptPointResult.Value;
                        transaction.Commit();
                    }
                }
            }
        }

        private LFlash lFlash;
        private bool isVertical;

        /// <summary>
        /// 鼠标移动时动态更新范围框位置
        /// </summary>
        private void Editor_PointMonitor(object sender, PointMonitorEventArgs e)
        {
            lFlash.Update(GetRec(e.Context.ComputedPoint, isVertical));
        }

        /// <summary>
        /// 选中自定义范围时启用输入框
        /// </summary>
        private void Custom_Checked(object sender, RoutedEventArgs e)
        {
            range.IsEnabled = true;
        }

        /// <summary>
        /// 选中预设范围时禁用输入框
        /// </summary>
        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (range != null)
            {
                range.IsEnabled = false;
            }
        }

        /// <summary>
        /// 获取指定位置的打印范围框（多边形）
        /// </summary>
        private static Polyline GetRec(Point3d point, bool isVertical)
        {
            Polyline polyline = new Polyline()
            {
                Closed = true,
                Layer = "0"
            };
            polyline.AddVertexAt(0, point.ToPoint2d(), 0, 0, 0);
            if (isVertical)
            {
                polyline.AddVertexAt(1, new Point2d(point.X + 21000, point.Y), 0, 0, 0);
                polyline.AddVertexAt(2, new Point2d(point.X + 21000, point.Y - 29700), 0, 0, 0);
                polyline.AddVertexAt(3, new Point2d(point.X, point.Y - 29700), 0, 0, 0);
            }
            else
            {
                polyline.AddVertexAt(1, new Point2d(point.X + 29700, point.Y), 0, 0, 0);
                polyline.AddVertexAt(2, new Point2d(point.X + 29700, point.Y - 21000), 0, 0, 0);
                polyline.AddVertexAt(3, new Point2d(point.X, point.Y - 21000), 0, 0, 0);
            }
            return polyline;
        }
    }
}