using System;
using System.Diagnostics;
using System.IO;

using Autodesk.AutoCAD.Runtime;

using Lightning.Information;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 获取本机运行环境信息
    /// </summary>
    public class LPInfo : CommandBase
    {
        [CommandMethod(nameof(LPInfo))]
        public static void Command()
        {
            // 获取操作系统信息
            string osInfo = Environment.OSVersion.ToString();
            // 获取.NET Framework版本
            string frameworkVersion = Environment.Version.ToString();
            // 获取计算机名称
            string machineName = Environment.MachineName;
            // 获取当前用户
            string userName = Environment.UserName;

            // 获取CAD主机应用程序信息
            string cadVersion = CADApp.Version.ToString();
            string cadProduct = CADApp.GetSystemVariable("PRODUCT").ToString();
            string cadExe = CADApp.GetSystemVariable("ACADVER").ToString();

            string file = Path.Combine(PCInfo.MyDocumentsDirectoryInfo.FullName, "运行环境信息.txt");

            // 输出信息
            File.WriteAllText(file,
                $"操作系统: {osInfo}\n" +
                $".NET Framework版本: {frameworkVersion}\n" +
                $"计算机名称: {machineName}\n" +
                $"当前用户: {userName}\n" +
                $"CAD版本: {cadVersion}\n" +
                $"CAD产品: {cadProduct}\n" +
                $"CAD主程序版本: {cadExe}\n"
            );

            Process.Start("Explorer", "/select," + file);
        }
    }
}