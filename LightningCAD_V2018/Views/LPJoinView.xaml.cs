using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using Lightning.Extension;

using LightningCAD.LightningExtension;

using Button = System.Windows.Controls.Button;
using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace LightningCAD.Views
{
    /// <summary>
    /// LPJoinView.xaml 的交互逻辑
    /// </summary>
    public partial class LPJoinView : ViewBase
    {
        private class PDF
        {
            public int Number { get; set; }
            public string FileName { get; set; }
        }
        private List<PDF> pdfs;
        public LPJoinView()
        {
            InitializeComponent();
            Loaded += LPJoinView_Loaded;
        }

        private void LPJoinView_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            up.Source = Information.GetFileInfo("Commands\\向上箭头.png").ToBitmapImage();
            down.Source = Information.GetFileInfo("Commands\\向下箭头.png").ToBitmapImage();
            pdfs = new List<PDF>();
        }

        private void AddButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF (*.pdf)|*.pdf",
                Multiselect = true,
                Title = "选择要合并的PDF文件"
            };
            bool? result = openFileDialog.ShowDialog();
            if (result != true)
            {
                return;
            }
            string[] fileNames = openFileDialog.FileNames;

            foreach (var item in fileNames)
            {
                PDF pDF = new PDF()
                {
                    Number = pdfs.Count + 1,
                    FileName = item
                };
                pdfs.Add(pDF);
            }
            pdfList.DataContext = null;
            pdfList.DataContext = pdfs;
        }

        private void DeleteButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (pdfList.SelectedItems.Count == 0) return;
            MessageBoxResult box = MessageBox.Show("确定删除吗？", "删除确认", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (box != MessageBoxResult.OK) return;
            CADApp.MainWindow.Handle.Focus();
            List<PDF> models = pdfList.SelectedItems.Cast<PDF>().ToList();
            foreach (PDF model in models)
            {
                pdfs.Remove(model);
            }
            for (int i = 0; i < pdfs.Count; i++)
            {
                pdfs[i].Number = i + 1;
            }
            pdfList.DataContext = null;
            pdfList.DataContext = pdfs;

        }

        private void ClearButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (pdfList.Items.Count == 0) return;
            MessageBoxResult box = MessageBox.Show("确定清空吗？", "删除确认", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (box != MessageBoxResult.OK) return;
            pdfs.Clear();
            pdfList.DataContext = null;
            pdfList.DataContext = pdfs;
        }

        private void MoveNumberButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            int selected = pdfList.SelectedItems.Count;
            if (selected == 0) return;//没选择返回
            int min = pdfList.SelectedItems.Cast<PDF>().Min(p => p.Number);//最小的编号
            int max = pdfList.SelectedItems.Cast<PDF>().Max(p => p.Number); //最大的编号
            if (max - min + 1 != selected) return;//非连续选择返回
            int n;//确保dataGrid选中的行离显示范围边缘有一定的距离，否则会bug
            if ((sender as Button).Name == "uper")
            {
                if (min == 1) return;//选择的是第一个无法上移

                pdfs[min - 2].Number = max;//第一个选择的上一个的编号等于最后一个选择的编号
                for (int i = 0; i < selected; i++)//所有选择的编号自减1
                {
                    pdfs[min - 1 + i].Number -= 1;
                }
                n = min - 7;//距离上边缘有5行
                if (n < 0) n = 0;
                min -= 2;//起始数字修正
                max -= 1;//起始数字修正
            }
            else
            {
                if (max == pdfs.Count) return;//选择的是第一个无法下移

                pdfs[max].Number = min;//最后一个选择的下一个的编号等于第一个选择的编号
                for (int i = 0; i < selected; i++)//所有选择的编号自增1
                {
                    pdfs[min - 1 + i].Number += 1;
                }
                n = max + 5;//距离下边缘有5行
                if (n > pdfs.Count - 1) n = pdfs.Count - 1;
                max += 1;//起始数字修正
            }
            pdfs = pdfs.OrderBy(p => p.Number).Cast<PDF>().ToList();//按顺序重新排序
            pdfList.DataContext = null;
            pdfList.DataContext = pdfs;
            pdfList.UpdateLayout();
            List<DataGridRow> rows = new List<DataGridRow>();
            for (int i = min; i < max; i++)
            {
                DataGridRow row = (DataGridRow)pdfList.ItemContainerGenerator.ContainerFromIndex(pdfs[i].Number - 1);
                rows.Add(row);
            }
            foreach (DataGridRow row in rows)
            {
                row.IsSelected = true;
            }
            pdfList.ScrollIntoView(pdfs[n]);//向上或下移就最外围一行显示留出5行边距，除非小于5
            pdfList.Focus();
        }

        private void OKButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                if (pdfs.Count == 0)
                {
                    return;
                }
                SaveFileDialog saveFileDialog = new SaveFileDialog()
                {
                    AddExtension = true,
                    DefaultExt = "pdf",
                    Filter = "PDF (*.pdf)|*.pdf",
                };
                bool? result = saveFileDialog.ShowDialog();
                if (result != true)
                {
                    return;
                }
                string file = saveFileDialog.FileName;
                List<string> list = pdfs.Select(p => p.FileName).ToList();
                PDFHelper.JoinFiles(list, file);
                Process.Start("Explorer", "/select," + file);
                Close();
            }
            catch (IOException exp)
            {
                if (exp.HResult == -2147024864)
                {
                    MessageBox.Show(exp.Message, "错误", MessageBoxButton.OK);
                }
                else
                {
                    exp.LogTo(Information.God);
                    MessageBoxResult messageBoxResult = MessageBox.Show("合并失败，请联系up主（QQ：1101412890），\n是否定位至错误文件？", "错误", MessageBoxButton.YesNo, MessageBoxImage.Error);
                    if (messageBoxResult == MessageBoxResult.Yes)
                    {
                        Process.Start("Explorer", "/select," + Information.God.ErrorLog);
                    }
                }
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
                MessageBoxResult messageBoxResult = MessageBox.Show("合并失败，请联系up主（QQ：1101412890），\n是否定位至错误文件？", "错误", MessageBoxButton.YesNo, MessageBoxImage.Error);
                if (messageBoxResult == MessageBoxResult.Yes)
                {
                    Process.Start("Explorer", "/select," + Information.God.ErrorLog);
                }
            }
        }
    }
}