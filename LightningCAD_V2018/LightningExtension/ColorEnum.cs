using Autodesk.AutoCAD.Colors;

using Color = Autodesk.AutoCAD.Colors.Color;

namespace LightningCAD.LightningExtension
{
    public enum ColorEnum
    {
        /// <summary>
        /// 红色
        /// </summary>
        Red,

        /// <summary>
        /// 黄色
        /// </summary>
        Yellow,

        /// <summary>
        /// 绿色
        /// </summary>
        Green,

        /// <summary>
        /// 青色
        /// </summary>
        Cyan,

        /// <summary>
        /// 蓝色
        /// </summary>
        Blue,

        /// <summary>
        /// 洋红色
        /// </summary>
        Magenta,

        /// <summary>
        /// 白色
        /// </summary>
        White,

        /// <summary>
        /// 深灰
        /// </summary>
        DarkGray,

        /// <summary>
        /// 浅灰
        /// </summary>
        LightGray,
    }

    public static class ColorEnumExtension
    {
        public static Color ToColor(this ColorEnum color)
        {
            switch (color)
            {
                case ColorEnum.Red:
                    return Color.FromColorIndex(ColorMethod.ByColor, 1);
                case ColorEnum.Yellow:
                    return Color.FromColorIndex(ColorMethod.ByColor, 2);
                case ColorEnum.Green:
                    return Color.FromColorIndex(ColorMethod.ByColor, 3);
                case ColorEnum.Cyan:
                    return Color.FromColorIndex(ColorMethod.ByColor, 4);
                case ColorEnum.Blue:
                    return Color.FromColorIndex(ColorMethod.ByColor, 5);
                case ColorEnum.Magenta:
                    return Color.FromColorIndex(ColorMethod.ByColor, 6);
                case ColorEnum.White:
                    return Color.FromColorIndex(ColorMethod.ByColor, 7);
                case ColorEnum.DarkGray:
                    return Color.FromColorIndex(ColorMethod.ByColor, 8);
                case ColorEnum.LightGray:
                    return Color.FromColorIndex(ColorMethod.ByColor, 9);
                default:
                    return Color.FromColorIndex(ColorMethod.ByColor, 7);
            }
        }

        public static int ToIndex(this ColorEnum color)
        {
            switch (color)
            {
                case ColorEnum.Red:
                    return 1;
                case ColorEnum.Yellow:
                    return 2;
                case ColorEnum.Green:
                    return 3;
                case ColorEnum.Cyan:
                    return 4;
                case ColorEnum.Blue:
                    return 5;
                case ColorEnum.Magenta:
                    return 6;
                case ColorEnum.White:
                    return 7;
                case ColorEnum.DarkGray:
                    return 8;
                case ColorEnum.LightGray:
                    return 9;
                default:
                    return 7;
            }
        }
    }
}