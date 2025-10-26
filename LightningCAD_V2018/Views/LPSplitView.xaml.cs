using System.Diagnostics;
using System.IO;
using System.Windows;

using LightningCAD.LightningExtension;

using DragEventArgs = System.Windows.DragEventArgs;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace LightningCAD.Views
{
    /// <summary>
    /// LPSplitView.xaml 的交互逻辑
    /// </summary>
    public partial class LPSplitView : ViewBase
    {
        public LPSplitView()
        {
            InitializeComponent();
        }

        private void File_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = System.Windows.DragDropEffects.Copy;
            e.Handled = true;
        }
        /// <summary>
        /// 允许拖动数据文件至窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void File_PreviewDrop(object sender, DragEventArgs e)
        {
            foreach (string f in (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop))
            {
                file.Text = f;
            }
        }

        private void File_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PDF(*.pdf)|*.pdf",
                Multiselect = false,
                Title = "选择要拆分的PDF文件"
            };
            bool? result = openFileDialog.ShowDialog();
            if (result != true)
            {
                return;
            }
            file.Text = openFileDialog.FileName;
        }

        private void Split_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(file.Text))
            {
                return;
            }
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog()
            {
                ShowNewFolderButton = true,
            };
            System.Windows.Forms.DialogResult dialogResult = folderBrowserDialog.ShowDialog();
            if (dialogResult != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }
            string dir = folderBrowserDialog.SelectedPath;
            PDFHelper.Split(file.Text, dir);
            Process.Start("Explorer", dir);
        }
    }
}