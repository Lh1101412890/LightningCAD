using System;
using System.Threading;
using System.Windows;

namespace LightningCAD_Install
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        private Mutex mutex;
        public App()
        {
            Startup += new StartupEventHandler(App_Startup);
        }

        private void App_Startup(object sender, StartupEventArgs e)
        {
            mutex = new Mutex(true, nameof(LightningCAD_Install), out bool ret);
            if (!ret)
            {
                MessageBox.Show("安装程序已启动", "LightningCAD");
                Environment.Exit(0);
            }
        }
    }
}