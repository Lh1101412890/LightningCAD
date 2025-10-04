using System;
using System.Windows.Input;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 类名称作为CAD命令
    /// </summary>
    public class CommandBase : ICommand
    {
#pragma warning disable CS0067 //从不使用警告
        public event EventHandler CanExecuteChanged;
#pragma warning disable CS0067

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            //传递过来的参数是被点击的按钮，RibbonButton类
            if (parameter is RibbonButton ribbonButton)
            {
                string command = ribbonButton.CommandParameter.ToString() + "\n";//命令参数
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                document.SendStringToExecute(command, true, false, false);
                document.Editor.WriteMessage(command);
            }
        }
    }
}