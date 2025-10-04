using System;
using System.Collections.Generic;

using Autodesk.AutoCAD.Runtime;

using LightningCAD.Views;

namespace LightningCAD.Commands
{
    public class Views : CommandBase
    {
        private static readonly List<Type> types = new List<Type>();
        public static void ShowWindow(ViewBase view)
        {
            Type type = view.GetType();
            if (types.Contains(type))
            {
                LightningApp.ShowMsg("窗口已经打开了！", 2);
                return;
            }
            view.Show();
            view.Closed += (sender, e) => { types.Remove(type); };
            types.Add(type);
        }
    }

    /// <summary>
    /// 快速编号
    /// </summary>
    public class LNumber : CommandBase
    {
        [CommandMethod(nameof(LNumber))]
        public static void Command()
        {
            Views.ShowWindow(new LNumberView());
        }
    }

    /// <summary>
    /// 批量打印
    /// </summary>
    public class LBPrint : CommandBase
    {
        [CommandMethod(nameof(LBPrint), CommandFlags.NoBlockEditor | CommandFlags.NoPaperSpace)]
        public static void Command()
        {
            Views.ShowWindow(new LBPrintView());
        }
    }

    /// <summary>
    /// 折断线
    /// </summary>
    public class LBLine : CommandBase
    {
        [CommandMethod(nameof(LBLine))]
        public static void Command()
        {
            Views.ShowWindow(new LBLineView());
        }
    }

    /// <summary>
    /// 创建自定义线型
    /// </summary>
    public class LCLinetype : CommandBase
    {
        [CommandMethod(nameof(LCLinetype))]
        public static void Command()
        {
            Views.ShowWindow(new LCLinetypeView());
        }
    }

    /// <summary>
    /// 墙柱图纸识别
    /// </summary>
    public class DRWallColumn : CommandBase
    {
        [CommandMethod(nameof(DRWallColumn), CommandFlags.NoBlockEditor)]
        public static void Command()
        {
            Views.ShowWindow(new DRWallColumnView());
        }
    }

    /// <summary>
    /// 结构梁识别
    /// </summary>
    public class DRBeam : CommandBase
    {
        [CommandMethod(nameof(DRBeam), CommandFlags.NoBlockEditor)]
        public static void Command()
        {
            Views.ShowWindow(new DRBeamView());
        }
    }

    /// <summary>
    /// PDF合并
    /// </summary>
    public class LPJoin : CommandBase
    {
        [CommandMethod(nameof(LPJoin))]
        public static void Command()
        {
            Views.ShowWindow(new LPJoinView());
        }
    }

    /// <summary>
    /// PDF拆分
    /// </summary>
    public class LPSplit : CommandBase
    {
        [CommandMethod(nameof(LPSplit))]
        public static void Command()
        {
            Views.ShowWindow(new LPSplitView());
        }
    }

}