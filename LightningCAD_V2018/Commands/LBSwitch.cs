using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Runtime;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 背景颜色及主题颜色切换（黑/白）
    /// </summary>
    public class LBSwitch : CommandBase
    {
        [CommandMethod(nameof(LBSwitch), CommandFlags.NoActionRecording | CommandFlags.NoUndoMarker | CommandFlags.Modal)]
        public static void Command()
        {
            AcadPreferences preferences = (AcadPreferences)CADApp.Preferences;
            uint now = preferences.Display.GraphicsWinModelBackgrndColor;

            //16进制的RGB颜色值，前两位为00，然后是R、G、B值
            if (now == 0x00000000)
            {
                preferences.Display.GraphicsWinModelBackgrndColor = 0x00ffffff; //白色
                preferences.Display.GraphicsWinLayoutBackgrndColor = 0x00ffffff; //白色
                CADApp.SetSystemVariable("COLORTHEME", 1);
            }
            else
            {
                preferences.Display.GraphicsWinModelBackgrndColor = 0x00000000; //黑色
                preferences.Display.GraphicsWinLayoutBackgrndColor = 0x00000000; //黑色
                CADApp.SetSystemVariable("COLORTHEME", 0);
            }
        }
    }
}