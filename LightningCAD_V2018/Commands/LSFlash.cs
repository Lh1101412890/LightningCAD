using System;
using System.Timers;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

using Lightning.Extension;

using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Timer = System.Timers.Timer;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 闪屏
    /// </summary>
    public class LSFlash : CommandBase
    {
        private static readonly Timer timer = new Timer();

        [CommandMethod(nameof(LSFlash), CommandFlags.NoBlockEditor | CommandFlags.NoPaperSpace | CommandFlags.Modal)]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Editor editor = document.Editor;

                RibbonButton button = null;
                RibbonControl ribbon = ComponentManager.Ribbon;
                if (ribbon != null)
                {
                    button = (RibbonButton)ribbon.FindTab("Lightning").FindItem(nameof(LSFlash), true);
                }

                if (timer.Enabled == true)
                {
                    timer.Enabled = false;
                    timer.Stop();
                    timer.Elapsed -= Timer_Elapsed;

                    if (button != null)
                    {
                        //取消按钮闪光状态
                        button.LargeImage = Information.GetFileInfo("Commands\\不闪屏.png").ToBitmapImage().Resize(32);
                        button.Image = Information.GetFileInfo("Commands\\不闪屏.png").ToBitmapImage().Resize(16);
                    }

                    editor.WriteMessage("闪屏停止!\n");
                    return;
                }

                PromptStringOptions options = new PromptStringOptions("指定时间间隔(秒)(间隔≥0.001,请谨慎设置时间间隔,过小可能会导致软件崩溃!):[放弃(U)]")
                {
                    AppendKeywordsToMessage = true,
                    DefaultValue = "0.5",
                };

            Loop: PromptResult result = editor.GetString(options);

                if (result.Status != PromptStatus.OK || result.StringResult is "u" || result.StringResult is "U")
                {
                    editor.WriteMessage("*取消*\n");
                    return;
                }

                bool check2 = double.TryParse(result.StringResult, out double ts);

                if (check2 && ts >= 0.001) //要求指定时间间隔≥1ms
                {
                    timer.Interval = (int)(ts * 1000);
                }
                else
                {
                    goto Loop;
                }

                timer.Enabled = true;
                timer.Elapsed += Timer_Elapsed;

                if (button != null)
                {
                    //按下按钮闪光状态
                    button.LargeImage = Information.GetFileInfo("Commands\\闪屏.png").ToBitmapImage().Resize(32);
                    button.Image = Information.GetFileInfo("Commands\\闪屏.png").ToBitmapImage().Resize(16);
                }
                editor.WriteMessage("闪屏开始!\n");

                timer.Start();
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }

        private static void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            dynamic preferences = CADApp.Preferences;
            uint random = (uint)new Random().Next(0, 16777216);
            preferences.Display.GraphicsWinModelBackgrndColor = random;
        }
    }
}
