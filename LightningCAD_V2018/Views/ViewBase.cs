using System;
using System.Windows;
using System.Windows.Forms;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Views
{
    public class ViewBase : Window
    {
        protected readonly Document document;
        protected readonly Database database;
        protected readonly Editor editor;
        public ViewBase()
        {
            Icon = Information.GetFileInfo("Lightning.ico").ToBitmapImage();
            ShowInTaskbar = false;
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.Manual;
            SizeToContent = SizeToContent.WidthAndHeight;

            SourceInitialized += ViewBase_SourceInitialized;
            Activated += ViewBase_Activated;
            Closed += ViewBase_Closed;
            MouseLeave += ViewBase_MouseLeave;

            document = CADApp.DocumentManager.MdiActiveDocument;
            database = document.Database;
            editor = document.Editor;

            CADApp.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
            CADApp.DocumentManager.DocumentActivationChanged += DocumentManager_DocumentActivationChanged;
            document.BeginDocumentClose += Document_BeginDocumentClose;

            StartListen();
        }

        // 激活其他文档时自动切换回原文档
        private void KeepCurrentDocument()
        {
            Document doc = CADApp.DocumentManager.MdiActiveDocument;
            if (doc == null || !doc.Equals(document))
            {
                //跳回到打开窗口时的文档
                foreach (var item in CADApp.DocumentWindowCollection)
                {
                    if (item is Autodesk.AutoCAD.Windows.DrawingDocumentWindow window && window.Document.Equals(document))
                    {
                        window.Activate();
                        break;
                    }
                }
                LightningApp.ShowMsg("命令窗口激活时禁止切换文档", 3);
            }
        }
        private void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e) => KeepCurrentDocument();
        private void DocumentManager_DocumentActivationChanged(object sender, DocumentActivationChangedEventArgs e) => KeepCurrentDocument();

        // 文档即将关闭时先关闭窗口，避免异常
        private void Document_BeginDocumentClose(object sender, DocumentBeginCloseEventArgs e) => Close();

        private void ViewBase_Closed(object sender, EventArgs e)
        {
            document.CancelCommand();
            this.SetLocation(Information.God);
            CADApp.DocumentManager.DocumentActivationChanged -= DocumentManager_DocumentActivationChanged;
            CADApp.DocumentManager.DocumentActivated -= DocumentManager_DocumentActivated;
            StopListen();
        }
        private void ViewBase_SourceInitialized(object sender, EventArgs e)
        {
            this.GetLocation(Information.God);
            IntPtr cad = CADApp.MainWindow.Handle;
            this.SetOwner(cad);
        }
        private void ViewBase_Activated(object sender, EventArgs e)
        {
            CADApp.MainWindow.Handle.Focus();
            Activated -= ViewBase_Activated;
        }
        private void ViewBase_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e) => CADApp.MainWindow.Handle.Focus();

        public void StartListen()
        {
            myKeyEventHandeler = new KeyEventHandler(Hook_KeyDown);
            k_hook.KeyDownEvent += myKeyEventHandeler;//钩住键按下
            k_hook.Start();//安装键盘钩子
        }
        public void StopListen()
        {
            k_hook.KeyDownEvent -= myKeyEventHandeler;//取消按键事件
            k_hook.Stop();
        }
        private KeyEventHandler myKeyEventHandeler;//按键钩子事件处理器
        private readonly KeyboardHook k_hook = new KeyboardHook();
        private void Hook_KeyDown(object sender, KeyEventArgs e)
        {
            //Alt + `关闭窗口
            if (Control.ModifierKeys == Keys.Alt && e.KeyCode == Keys.Oemtilde)
            {
                Close();
            }
        }
    }
}