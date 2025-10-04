using Autodesk.Windows;

using Lightning.Extension;

using LightningCAD.Commands;
using LightningCAD.LightningExtension;

namespace LightningCAD.Extension
{
    public static class RibbonExtension
    {
        /// <summary>
        /// 添加一个名称为title的RibbonPanel
        /// </summary>
        /// <param name="tab"></param>
        /// <param name="title"></param>
        /// <returns>关联的RibbonPanelSource</returns>
        public static RibbonPanelSource AddPanelSource(this RibbonTab tab, string title)
        {
            RibbonPanelSource panelSource = new RibbonPanelSource()
            {
                Title = title,
                Id = title,
            };
            RibbonPanel panel = new RibbonPanel()
            {
                Source = panelSource,
            };
            tab.Panels.Add(panel);
            return panelSource;
        }

        /// <summary>
        /// 在source中添加一个命令按钮，并返回命令按钮，无需输入item的【Id】、【CommandHandler】、【CommandParameter】、默认显示文字及图片
        /// </summary>
        /// <typeparam name="T">绑定的命令类</typeparam>
        /// <param name="source"></param>
        /// <param name="item"></param>
        /// <param name="image">图片名称，Commmands路径下</param>
        /// <returns>item</returns>
        public static RibbonCommandItem AddCommandItem<T>(this RibbonPanelSource source, RibbonCommandItem item, string image = "") where T : CommandBase, new()
        {
            item.Id = typeof(T).Name;
            item.CommandHandler = new T();
            item.CommandParameter = typeof(T).Name;
            item.ShowText = true;
            if (image != "")
            {
                item.ShowImage = true;
                item.LargeImage = Information.GetFileInfo("Commands\\" + image).ToBitmapImage().Resize(32);
                item.Image = Information.GetFileInfo("Commands\\" + image).ToBitmapImage().Resize(16);
            }
            source.Items.Add(item);
            return item;
        }

        public static RibbonRowPanel AddRibbonRowPanel(this RibbonPanelSource source, RibbonRowPanel item)
        {
            source.Items.Add(item);
            return item;
        }

        /// <summary>
        /// 在source中添加一个命令按钮，并返回命令按钮，无需输入item的【Id】、【CommandHandler】、【CommandParameter】、默认显示文字及图片
        /// </summary>
        /// <typeparam name="T">绑定的命令类</typeparam>
        /// <param name="source"></param>
        /// <param name="item"></param>
        /// <param name="image">图片名称，Commmands路径下</param>
        /// <returns>item</returns>
        public static RibbonCommandItem AddCommandItem<T>(this RibbonRowPanel source, RibbonCommandItem item, string image = "") where T : CommandBase, new()
        {
            item.Id = typeof(T).Name;
            item.CommandHandler = new T();
            item.CommandParameter = typeof(T).Name;
            item.ShowText = true;
            if (image != "")
            {
                item.ShowImage = true;
                item.LargeImage = Information.GetFileInfo("Commands\\" + image).ToBitmapImage().Resize(32);
                item.Image = Information.GetFileInfo("Commands\\" + image).ToBitmapImage().Resize(16);
            }
            source.Items.Add(item);
            source.Items.Add(new RibbonRowBreak());
            return item;
        }

        /// <summary>
        /// 将toolTip添加到item，并返回toolTip，无需输入toolTip的【Command】，【Shortcut】
        /// </summary>
        /// <param name="item"></param>
        /// <param name="toolTip"></param>
        public static RibbonToolTip AddToolTip(this RibbonCommandItem item, RibbonToolTip toolTip)
        {
            toolTip.Command = item.CommandParameter.ToString();
            toolTip.Shortcut = toolTip.Command.Substring(0, 3);
            item.ToolTip = toolTip;
            return toolTip;
        }
    }
}