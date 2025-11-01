using System.Diagnostics;
using System.IO;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.Windows;

using Lightning.Extension;
using Lightning.Information;

using LightningCAD;
using LightningCAD.Commands;
using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Orientation = System.Windows.Controls.Orientation;

//标记扩展程序的入口点
[assembly: ExtensionApplication(typeof(LightningApp))]
namespace LightningCAD
{
    public class LightningApp : IExtensionApplication
    {
        private static bool IsGod => PCInfo.IsGod;
        private static RibbonTab Tab { get; set; }
        /// <summary>
        /// CAD工作目录，第一个打开文档所在的目录，会在此处生成打印日志文件
        /// </summary>
        public static string Dir { get; private set; } = "";

        /// <summary>
        /// 显示消息
        /// </summary>
        /// <param name="msg">提示信息</param>
        /// <param name="time">显示时长</param>
        /// <param name="always">是否一直显示</param>
        public static void ShowMsg(string msg, int time = 0, bool always = false) => Information.God.ShowMessage(msg, time, always);

        public void Initialize()
        {
            if (IsGod)
            {
                CADApp.MainWindow.WindowState = Window.State.Maximized;
            }
            else
            {
                ShowMsg("Lightning插件作者：【不要干施工】，点击去b站充电，插件群：785371506！", 25);
            }
            // 事件存在就不会被其他插件初始化删掉（此事件会打开多次）
            ComponentManager.ItemInitialized += LRibbon_ItemInitialized;
            CADApp.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
        }

        public void Terminate()
        {
            //关闭消息窗口
            foreach (var item in Process.GetProcessesByName("LightningMessage"))
            {
                item.Kill();
            }
        }

        private void DocumentManager_DocumentCreated(object sender, Autodesk.AutoCAD.ApplicationServices.DocumentCollectionEventArgs e)
        {
            Autodesk.AutoCAD.ApplicationServices.Document document = CADApp.DocumentManager.MdiActiveDocument;
            document.RegLightning();
            Editor editor = document.Editor;
            using (ViewTableRecord view = editor.GetCurrentView())
            {
                if (!view.Target.IsEqualTo(0, 0, 0))
                {
                    view.Target = Point3d.Origin;
                    editor.SetCurrentView(view); // 解决批量打印时位置定位问题
                    editor.WriteMessage("视图目标点已重置为原点\n");
                }
            }
            if (IsGod)
            {
                LSetting.Command();
            }
            if (Dir == "")
            {
                try
                {
                    FileInfo fileInfo = new FileInfo(e.Document.Name);
                    Dir = fileInfo.DirectoryName;
                }
                catch (System.Exception exp)
                {
                    exp.LogTo(Information.God);
                }
            }
        }
        private void LRibbon_ItemInitialized(object sender, RibbonItemEventArgs e) => CreateRibbon();

        [CommandMethod("LRibbon")]
        public static void CreateRibbon()
        {
            RibbonControl ribbon = ComponentManager.Ribbon;
            if (ribbon is null) return;
            Tab = ribbon.FindTab(Information.Brand);
            if (Tab != null) return;
            Tab = new RibbonTab() { Title = Information.Brand, Id = Information.Brand };
            ribbon.Tabs.Add(Tab);
            CreatGeneral();
            CreatText();
            CreatCoordinate();
            CreatDwgReader();
            CreatPDF();
            CreatOthers();
        }
        private static void CreatGeneral()
        {
            RibbonPanelSource general = Tab.AddPanelSource("通用");

            //关闭所有文件
            general.AddCommandItem<LCFiles>(new RibbonButton()
            {
                Text = "关闭\n文件",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "关闭.png").AddToolTip(new RibbonToolTip()
            {
                Title = "关闭文件",
                Content = "关闭所有文件，选择保存或不保存关闭！",
                Image = Information.GetFileInfo("ToolTips\\关闭文件.png").ToBitmapImage().ResizeOfWidth(350),
            });

            //背景切换
            general.AddCommandItem<LBSwitch>(new RibbonButton()
            {
                Text = "背景\n切换",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "背景切换.png").AddToolTip(new RibbonToolTip()
            {
                Title = "背景切换",
                Content = "将模型空间、布局和窗口界面的背景颜色在黑色和白色之间切换。"
            });

            //选择类似
            general.AddCommandItem<LSSimilar>(new RibbonButton()
            {
                Text = "选择\n类似",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "选择类似.png").AddToolTip(new RibbonToolTip()
            {
                Title = "选择类似",
                Content = "选择和当前选择的对象类似的所有对象【相同图层且相同类型或样式（直线、圆、标注……）】。"
            });

            //计算S L
            general.AddCommandItem<LCGeometry>(new RibbonButton()
            {
                Text = "计算\nS L",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "计算S L.png").AddToolTip(new RibbonToolTip()
            {
                Title = "计算S L",
                Content = "计算面积和长度\n计算长度的类型：直线/多段线/圆/椭圆/圆弧/椭圆弧/样条曲线\n计算面积的类型：多段线/圆/椭圆/样条曲线（闭合图形）\n计算结果在命令行显示，按F2可显示命令行列表，若命令行未打开请按Ctrl+9组合键",
                Image = Information.GetFileInfo("ToolTips\\计算S L_1.png").ToBitmapImage().ResizeOfWidth(350),
                ExpandedContent = "导出数据至Excel",
                ExpandedImage = Information.GetFileInfo("ToolTips\\计算S L_2.png").ToBitmapImage().ResizeOfWidth(350),
            });

            //快速编号
            general.AddCommandItem<LNumber>(new RibbonButton()
            {
                Text = "快速\n编号",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "快速编号.png").AddToolTip(new RibbonToolTip()
            {
                Title = "快速编号",
            });

            //批量打印
            general.AddCommandItem<LBPrint>(new RibbonButton()
            {
                Text = "批量\n打印",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "批量打印.png").AddToolTip(new RibbonToolTip()
            {
                Title = "批量打印",
                Content = "快速添加或识别打印范围，一键批量打印为PDF文件。",
                Image = Information.GetFileInfo("ToolTips\\批量打印1.png").ToBitmapImage().ResizeOfWidth(350),
                ExpandedContent = "自动识别块",
                ExpandedImage = Information.GetFileInfo("ToolTips\\批量打印2.png").ToBitmapImage().ResizeOfWidth(350),
                CustomContent = "如需查看更详细视频讲解，请点击“关于”，单击up头像快速进入up主页。",
            });

            //折断线
            general.AddCommandItem<LBLine>(new RibbonButton()
            {
                Text = "折断\n线",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "折断线.png").AddToolTip(new RibbonToolTip()
            {
                Title = "折断线",
                Content = "设置折断符号比例和延伸比例自定义折断线",
                Image = Information.GetFileInfo("ToolTips\\折断线.png").ToBitmapImage().ResizeOfWidth(350)
            });

            //自定义线型
            general.AddCommandItem<LCLinetype>(new RibbonButton()
            {
                Text = "自定义\n线型",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "自定义线型.png").AddToolTip(new RibbonToolTip()
            {
                Title = "自定义线型",
                Content = "CAD中自定义线型，通过设置一组重复的笔画来定义线型，包含直线、空格、文字",
                Image = Information.GetFileInfo("ToolTips\\自定义线型.png").ToBitmapImage().ResizeOfWidth(350)
            });

            //直线转多段线
            general.AddCommandItem<LLTPolyline>(new RibbonButton()
            {
                Text = "直线转\n多段线",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "直线转多段线.png").AddToolTip(new RibbonToolTip()
            {
                Title = "直线转多段线",
                Content = "将选择的直线转换为多段线"
            });
        }
        private static void CreatText()
        {
            RibbonPanelSource text = Tab.AddPanelSource("文字");

            //提取文字
            text.AddCommandItem<LTExtract>(new RibbonButton()
            {
                Text = "提取\n文字",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "提取文字.png").AddToolTip(new RibbonToolTip()
            {
                Title = "提取文字",
            });

            //合并文字
            text.AddCommandItem<LTJoin>(new RibbonButton()
            {
                Text = "合并\n文字",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "合并文字.png").AddToolTip(new RibbonToolTip()
            {
                Title = "合并文字",
                Content = "选定第一个单行文字，然后将后续单行文字合并至第一个末尾。",
                Image = Information.GetFileInfo("ToolTips\\合并文字.png").ToBitmapImage().Resize(350),
            });

            //单行文字转多行文字
            text.AddCommandItem<LTMText>(new RibbonButton()
            {
                Text = "单行转\n多行字",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "多行文字.png").AddToolTip(new RibbonToolTip()
            {
                Title = "单行文字转多行文字",
                Content = "将选择的单行文字转换为多行文字"
            });

            //背景遮罩
            text.AddCommandItem<LTWipeout>(new RibbonButton()
            {
                Text = "背景\n遮罩",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "背景遮罩.png").AddToolTip(new RibbonToolTip()
            {
                Title = "背景遮罩",
                Content = "对单行或多行文字的背景进行遮罩\n多行文字的遮罩范围为文字的范围框",
            });

            text.Items.Add(new RibbonSeparator());

            RibbonRowPanel group1 = text.AddRibbonRowPanel(new RibbonRowPanel());
            //文字居左对齐
            group1.AddCommandItem<LTLeft>(new RibbonButton()
            {
                Text = "居左对齐",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "居左对齐.png").AddToolTip(new RibbonToolTip()
            {
                Title = "居左对齐",
                Content = "文字居左对齐",
            });
            //文字居上对齐
            group1.AddCommandItem<LTTop>(new RibbonButton()
            {
                Text = "居上对齐",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "居上对齐.png").AddToolTip(new RibbonToolTip()
            {
                Title = "居上对齐",
                Content = "文字居上对齐",
            });
            //文字居下对齐
            group1.AddCommandItem<LTBottomt>(new RibbonButton()
            {
                Text = "居下对齐",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "居下对齐.png").AddToolTip(new RibbonToolTip()
            {
                Title = "居下对齐",
                Content = "文字居下对齐",
            });

            RibbonRowPanel group2 = text.AddRibbonRowPanel(new RibbonRowPanel());
            //文字居右对齐
            group2.AddCommandItem<LTRight>(new RibbonButton()
            {
                Text = "居右对齐",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "居右对齐.png").AddToolTip(new RibbonToolTip()
            {
                Title = "居右对齐",
                Content = "文字居右对齐",
            });
            //文字水平居中对齐
            group2.AddCommandItem<LTHorizontal>(new RibbonButton()
            {
                Text = "水平居中",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "水平居中对齐.png").AddToolTip(new RibbonToolTip()
            {
                Title = "水平居中对齐",
                Content = "文字水平居中对齐",
            });
            //文字垂直居中对齐
            group2.AddCommandItem<LTVertical>(new RibbonButton()
            {
                Text = "垂直居中",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "垂直居中对齐.png").AddToolTip(new RibbonToolTip()
            {
                Title = "垂直居中对齐",
                Content = "文字垂直居中对齐",
            });
        }
        private static void CreatCoordinate()
        {
            RibbonPanelSource coordinate = Tab.AddPanelSource("坐标");

            RibbonRowPanel group1 = coordinate.AddRibbonRowPanel(new RibbonRowPanel());
            //坐标标注mm
            group1.AddCommandItem<LMmCoordinate>(new RibbonButton()
            {
                Text = "坐标mm",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "坐标标注.png").AddToolTip(new RibbonToolTip()
            {
                Title = "坐标mm",
                Content = "当图形1单位对应1mm时用这个标注",
                Image = Information.GetFileInfo("ToolTips\\坐标标注.png").ToBitmapImage().ResizeOfWidth(350)
            });
            //坐标标注m
            group1.AddCommandItem<LMCoordinate>(new RibbonButton()
            {
                Text = "坐标m",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "坐标标注.png").AddToolTip(new RibbonToolTip()
            {
                Title = "坐标m",
                Content = "当图形1单位对应1m时用这个标注",
                Image = Information.GetFileInfo("ToolTips\\坐标标注.png").ToBitmapImage().ResizeOfWidth(350)
            });
            //坐标标注更新
            group1.AddCommandItem<LCRefresh>(new RibbonButton()
            {
                Text = "坐标更新",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "更新.png").AddToolTip(new RibbonToolTip()
            {
                Title = "坐标更新",
                Content = "更新那些移动过标注点的坐标数据",
                Image = Information.GetFileInfo("ToolTips\\坐标更新.png").ToBitmapImage().ResizeOfWidth(350)
            });

            RibbonRowPanel group2 = coordinate.AddRibbonRowPanel(new RibbonRowPanel());
            //高程标注
            group2.AddCommandItem<LElevation>(new RibbonButton()
            {
                Text = "高程标注",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "高程标注.png").AddToolTip(new RibbonToolTip()
            {
                Title = "高程标注",
                Content = "标注点的高程，移动点水平位置或删除点后高程标注失效",
            });
            //高程更新
            group2.AddCommandItem<LERefresh>(new RibbonButton()
            {
                Text = "高程更新",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "更新.png").AddToolTip(new RibbonToolTip()
            {
                Title = "高程更新",
                Content = "更新那些修改过高程的标注",
            });
            //清理失效
            group2.AddCommandItem<LEClear>(new RibbonButton()
            {
                Text = "高程清理",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "清理.png").AddToolTip(new RibbonToolTip()
            {
                Title = "高程清理",
                Content = "删除失效的高程标注",
            });

            RibbonRowPanel group3 = coordinate.AddRibbonRowPanel(new RibbonRowPanel());
            //导入坐标
            group3.AddCommandItem<LIPoints>(new RibbonButton()
            {
                Text = "导入坐标",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "导入点.png").AddToolTip(new RibbonToolTip()
            {
                Title = "导入坐标",
                Content = "导入Excel文件中的坐标点，需要手动指定列标题（X/Y/Z）。\n请按照以下格式录入数据，在表顶插入一行并在对应位置输入列标题（X/Y/Z），X对应N(北方向),Y对应E(东方向)。",
                Image = Information.GetFileInfo("ToolTips\\导入坐标.png").ToBitmapImage(),
            });
            //多点建线
            group3.AddCommandItem<LPTPolyline>(new RibbonButton()
            {
                Text = "多点建线",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "多点建线.png").AddToolTip(new RibbonToolTip()
            {
                Title = "多点建线",
                Content = "通过Excel文件中的坐标点创建多段线，按从上到下的顺序，依次为多段线的顶点，需要手动指定列标题（X/Y）。\n请按照以下格式录入数据，在表顶插入一行并在对应位置输入列标题（X/Y），X对应N(北方向),Y对应E(东方向)。",
                Image = Information.GetFileInfo("ToolTips\\多点建线.png").ToBitmapImage(),
            });
            //导出坐标
            group3.AddCommandItem<LOPoints>(new RibbonButton()
            {
                Text = "导出坐标",
                Size = RibbonItemSize.Standard,
                Orientation = Orientation.Horizontal,
            }, "导出点.png").AddToolTip(new RibbonToolTip()
            {
                Title = "导出坐标",
                Content = "导出坐标点数据到Excel文件，数据进度含义为精确到小数点后几位。X对应N(北方向),Y对应E(东方向)。",
            });
        }
        private static void CreatDwgReader()
        {
            RibbonPanelSource debug = Tab.AddPanelSource("图纸识别");

            //墙柱识别
            debug.AddCommandItem<DRWallColumn>(new RibbonButton()
            {
                Text = "墙柱",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "结构柱.png").AddToolTip(new RibbonToolTip()
            {
                Title = "墙柱识别",
                Content = "识别墙柱结构平面图，对比平面图和大样柱轮廓尺寸等是否一致，还可导出数据用于Revit建模。数据精度为5。\n点击“关于”命令，点击up头像快速进入up主页还可查看相关视频演示。",
            });

            //梁识别
            debug.AddCommandItem<DRBeam>(new RibbonButton()
            {
                Text = "梁",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,

            }, "结构梁.png").AddToolTip(new RibbonToolTip()
            {
                Title = "梁识别",
                Content = "识别平面梁图，对比标注和梁尺寸等是否一致，还可导出数据用于Revit建模。数据精度为5。",
            });
        }
        private static void CreatPDF()
        {
            RibbonPanelSource pdf = Tab.AddPanelSource("PDF");

            //PDF合并
            pdf.AddCommandItem<LPJoin>(new RibbonButton()
            {
                Text = "PDF\n合并",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "PDF合并.png").AddToolTip(new RibbonToolTip()
            {
                Title = "PDF合并",
                Content = "按顺序添加PDF文件，合并为一个。"
            });

            //PDF拆分
            pdf.AddCommandItem<LPSplit>(new RibbonButton()
            {
                Text = "PDF\n拆分",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "PDF拆分.png").AddToolTip(new RibbonToolTip()
            {
                Title = "PDF拆分",
                Content = "拆分PDF文件。"
            });
        }
        private static void CreatOthers()
        {
            RibbonPanelSource others = Tab.AddPanelSource("其他");

            //版本转换
            others.AddCommandItem<LFVersion>(new RibbonButton()
            {
                Text = "版本\n转换",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "版本转换.png").AddToolTip(new RibbonToolTip()
            {
                Title = "版本转换",
                Content = "选择文件夹，批量转换文件版本。"
            });

            //图库
            others.AddCommandItem<LLibrary>(new RibbonButton()
            {
                Text = "图库",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "图库.png").AddToolTip(new RibbonToolTip()
            {
                Title = "图库",
                Content = "包含部分命令扩展块，以及up画的一些其他图块。"
            });

            //闪屏
            others.AddCommandItem<LSFlash>(new RibbonButton()
            {
                Text = "闪屏",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "不闪屏.png").AddToolTip(new RibbonToolTip()
            {
                Title = "闪屏",
                Content = "工作区颜色随机变化\n指定时间间隔后，工作区颜色随机变化。\n请谨慎设置时间间隔，过小可能会导致软件崩溃！",
            });

            //更新说明
            others.AddCommandItem<LUNotes>(new RibbonButton()
            {
                Text = "更新\n说明",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }, "更新说明.png");

            //CAD习惯设置
            others.AddCommandItem<LSetting>(new RibbonButton()
            {
                Text = "CAD\n习惯\n设置",
                Size = RibbonItemSize.Large,
                Orientation = Orientation.Vertical,
            }).AddToolTip(new RibbonToolTip()
            {
                Title = "up主使用CAD的习惯设置",
                Content =
@"------------------------------------------------------------------
一、将CAD设置为up主习惯的设置
    1. 关闭格栅
    2. 十字光标大小100
    3. 打开快捷特性
    4. 打开动态输入
    5. 二维捕捉模式（端点、中点、圆心、几何中心、节点、象限点、交点、延长线、垂足、切点、平行线）
    6. 打开选择循环
    7. 关闭导航栏
    8. 保存时不创建*.bak文件
    9. 不显示“代理信息”对话框
    10.frame显示各种边框但不打印
    11.打开“对象编组”选择模式，关闭“关联图案填充”选择模式
    12.加载自带线型：BORDER、CENTER、DASHED、DOT、ZIGZAG（波浪折线）
------------------------------------------------------------------
二、创建样式
    1. 各种比例的文字样式
    2. 各种比例的标注样式
    3. 各种比例的多重引线样式
    4. 临时用电、临时用水线型",
            });
        }
    }
}