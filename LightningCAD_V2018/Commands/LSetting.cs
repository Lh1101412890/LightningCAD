using System.Collections.Generic;
using System.IO;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;

using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using Registry = Microsoft.Win32.Registry;
using RegistryKey = Microsoft.Win32.RegistryKey;

namespace LightningCAD.Commands
{
    /// <summary>
    /// CAD习惯设置
    /// </summary>
    public class LSetting : CommandBase
    {
        [CommandMethod(nameof(LSetting))]
        public static void Command()
        {
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            Database database = document.Database;
            Editor editor = document.Editor;
            try
            {
                CADApp.SetSystemVariable("gridmode", 0);//关闭格栅F7//0 关闭，1 打开
                CADApp.SetSystemVariable("cursorsize", 100);//十字光标大小
                CADApp.SetSystemVariable("isavebak", 0);//保存时不创建BAK文件//0 关闭，1 打开
                CADApp.SetSystemVariable("proxynotice", 0);//不显示“代理信息”对话框//0 关闭，1 打开
                CADApp.SetSystemVariable("qpmode", 1);//打开快捷特性//0 关闭，1 打开
                CADApp.SetSystemVariable("dynmode", 1);//打开动态输入//0 关闭，1 打开
                CADApp.SetSystemVariable("navbardisplay", 0);//关闭导航栏//0 关闭，1 打开
                CADApp.SetSystemVariable("pickstyle", 1);//选择组中一个对象就选择整个组，选择填充时不选择边界//0 都关闭，1 打开对象编组、2 打开关联图案填充，3 都打开
                CADApp.SetSystemVariable("selectioncycling", 2);//打开选择循环//0 关闭，1 打开但不显示列表框，2 打开并显示列表框
                CADApp.SetSystemVariable("frame", 2);//将imageframe、dwfframe、pdfframe、dgnframe、xclipframe、wipeoutframe一同设置//0 边框不可见且不打印，在选择时暂时显示边框，1 显示并打印边框，2 显示但不打印边框，3 前面这些设置不一样时为3，不能手动设置为3

                // 0  |  1   |  2   |  4   |  8   |   16   |  32  |   64   | 128  | 256  |  512   |   1024   |   2048   |  4096  |  8192
                // 无 | 端点 | 中点 | 圆心 | 节点 | 象限点 | 交点 | 插入点 | 垂足 | 切点 | 最近点 | 几何中心 | 外观交点 | 延长线 | 平行线
                CADApp.SetSystemVariable("osmode", 1 | 2 | 4 | 8 | 16 | 32 | 128 | 256 | 1024 | 4096 | 8192);//设置对象捕捉
                document.SendStringToExecute("filetab\n", true, false, false);

                AcadPreferences preferences = (AcadPreferences)CADApp.Preferences;
                preferences.Output.AutomaticPlotLog = false;//关闭打印和发布记录日志

                RibbonControl ribbon = ComponentManager.Ribbon;
                if (ribbon != null)//关掉不需要的选项卡
                {
                    string[] tabs = new string[] { "附加模块", "协作", "自动化", "精选应用", "A360" };
                    foreach (var item in tabs)
                    {
                        if (ribbon.Tabs.ToList().Exists(t => t.Title == item))
                        {
                            ribbon.Tabs.First(t => t.Title == item).IsVisible = false;
                        }
                    }
                }

                CreatStyles(document);

                //线型
                using (Transaction transaction = database.NewTransaction())
                {
                    LinetypeTable linetypeTable = (LinetypeTable)transaction.GetObject(database.LinetypeTableId, OpenMode.ForWrite);
                    TextStyleTable textStyleTable = (TextStyleTable)transaction.GetObject(database.TextStyleTableId, OpenMode.ForWrite);
                    //自定义临时用电线型
                    if (!linetypeTable.Has("Lightning-V"))
                    {
                        LinetypeTableRecord linetype = new LinetypeTableRecord()
                        {
                            Name = "Lightning-V",
                            AsciiDescription = "---- V ---- V ---- V ----",
                            PatternLength = 20,
                            NumDashes = 3,
                        };
                        linetype.SetDashLengthAt(0, 15);//15单位的直线
                        linetype.SetDashLengthAt(1, -2.5);//2.5个单位的空格
                        linetype.SetTextAt(1, "V");//文字内容
                        linetype.SetShapeStyleAt(1, textStyleTable["LText-1：1"]);//设置文字的样式
                        linetype.SetShapeNumberAt(1, 0);//设置空格处包含的图案图形
                        linetype.SetShapeScaleAt(1, 1);//图形的缩放比例
                        linetype.SetShapeRotationAt(1, 0);//图形的旋转弧度
                        linetype.SetShapeIsUcsOrientedAt(1, false);//图形是否面向ucs方向
                        linetype.SetShapeOffsetAt(1, new Vector2d(-1.225, -1.75));//图形在X轴Y轴方向上的偏移单位
                        linetype.SetDashLengthAt(2, -2.5);//2.5个单位的空格
                        linetypeTable.Add(linetype);
                        transaction.AddNewlyCreatedDBObject(linetype, true);
                    }
                    //自定义临时用水线型
                    if (!linetypeTable.Has("Lightning-S"))
                    {
                        LinetypeTableRecord linetype = new LinetypeTableRecord()
                        {
                            Name = "Lightning-S",
                            AsciiDescription = "---- S ---- S ---- S ----",
                            PatternLength = 20,//线型总长度，不包含图形
                            NumDashes = 3,//组成线型的笔画数目，不包含图形
                        };
                        linetype.SetDashLengthAt(0, 15);//15单位的直线
                        linetype.SetDashLengthAt(1, -2.5);//2.5个单位的空格
                        linetype.SetTextAt(1, "S");//文字内容
                        linetype.SetShapeStyleAt(1, textStyleTable["LText-1：1"]);//设置文字的样式
                        linetype.SetShapeNumberAt(1, 0);//设置空格处包含的图案图形
                        linetype.SetShapeScaleAt(1, 1);//图形的缩放比例
                        linetype.SetShapeRotationAt(1, 0);//图形的旋转弧度
                        linetype.SetShapeIsUcsOrientedAt(1, false);//图形是否面向ucs方向
                        linetype.SetShapeOffsetAt(1, new Vector2d(-1.225, -1.75));//图形在X轴Y轴方向上的偏移单位
                        linetype.SetDashLengthAt(2, -2.5);//2.5个单位的空格
                        linetypeTable.Add(linetype);
                        transaction.AddNewlyCreatedDBObject(linetype, true);
                    }

                    RegistryKey registry;
#if C18
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R22.0\\InstalledProducts", false);
#elif C19
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R23.0\\InstalledProducts", false);
#elif C20
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R23.1\\InstalledProducts", false);
#elif C21
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R24.0\\InstalledProducts", false);
#elif C22
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R24.1\\InstalledProducts", false);
#elif C23
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R24.2\\InstalledProducts", false);
#elif C24
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R24.3\\InstalledProducts", false);
#elif C25
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R25.0\\InstalledProducts", false);
#elif C26
                    registry = Registry.LocalMachine.OpenSubKey("Software\\Autodesk\\AutoCAD\\R25.1\\InstalledProducts", false);
#endif
                    DirectoryInfo directory = new DirectoryInfo($"{registry.GetValue("")}UserDataCache");
                    string dir = directory.GetDirectories().First(d => d.Name.Contains('-')).FullName;
                    string file = $"{dir}\\Support\\acadiso.lin";//CAD安装目录的线型文件
                    registry.Dispose();
                    //需要加载的CAD自带线型
                    string[] linetypes =
                    {
                        "BORDER",//长虚线、点
                        "CENTER",//中心线
                        "DASHED",//虚线
                        "DOT",//点
                        "ZIGZAG"//波浪线，折线
                    };
                    foreach (var item in linetypes)
                    {
                        if (!linetypeTable.Has(item))
                        {
                            database.LoadLineTypeFile(item, file);
                        }
                    }
                    transaction.Commit();
                }

                document.Editor.WriteMessage("CAD习惯设置完成\n");
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }

        public static readonly List<string> Proportion = new List<string>()
        {
            "100：1",
            "10：1",
            "8：1",
            "4：1",
            "2：1",
            "1：1",
            "1：2",
            "1：4",
            "1：5",
            "1：8",
            "1：10",
            "1：16",
            "1：20",
            "1：30",
            "1：40",
            "1：50",
            "1：100",
            "1：150",
            "1：200",
            "1：250",
            "1：300",
            "1：500",
            "1：1000",
        };

        /// <summary>
        /// 创建文字、标注、多重引线样式
        /// </summary>
        /// <param name="document"></param>
        public static void CreatStyles(Document document)
        {
            using (DocumentLock @lock = document.LockDocument())
            {
                Database database = document.Database;
                using (Transaction transaction = database.NewTransaction())
                {
                    //打开块表
                    BlockTable blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                    //打开字体样式表
                    TextStyleTable textStyleTable = (TextStyleTable)transaction.GetObject(database.TextStyleTableId, OpenMode.ForWrite);
                    //打开标注样式表
                    DimStyleTable dimStyleTable = (DimStyleTable)transaction.GetObject(database.DimStyleTableId, OpenMode.ForWrite);
                    //打开多重引线字典
                    DBDictionary dBDictionary = (DBDictionary)transaction.GetObject(database.MLeaderStyleDictionaryId, OpenMode.ForWrite);

                    if (!textStyleTable.Has("LText"))
                    {
                        TextStyleTableRecord text = new TextStyleTableRecord()
                        {
                            Name = "LText",
                            FileName = "gbenor.shx",
                            BigFontFileName = "gbcbig.shx",
                            TextSize = 0,
                        };
                        textStyleTable.Add(text);
                        transaction.AddNewlyCreatedDBObject(text, true);
                    }

                    if (!textStyleTable.Has("LineText"))
                    {
                        TextStyleTableRecord text = new TextStyleTableRecord()
                        {
                            Name = "LineText",
                            FileName = "gbenor.shx",
                            BigFontFileName = "gbcbig.shx",
                            TextSize = 3.5,
                        };
                        textStyleTable.Add(text);
                        transaction.AddNewlyCreatedDBObject(text, true);
                    }

                    foreach (var item in Proportion)
                    {
                        double first = double.Parse(item.Split('：').First());
                        double last = double.Parse(item.Split('：').Last());

                        string text = $"LText-{item}";
                        if (!textStyleTable.Has(text))
                        {
                            TextStyleTableRecord textStyle = new TextStyleTableRecord()
                            {
                                Name = text,
                                FileName = "gbenor.shx",
                                BigFontFileName = "gbcbig.shx",
                                TextSize = last / first * 3.5,
                            };
                            // 将字体样式添加到字体样式表
                            textStyleTable.Add(textStyle);
                            transaction.AddNewlyCreatedDBObject(textStyle, true);
                        }

                        string dim = $"LDim-{item}";
                        if (!dimStyleTable.Has(dim))
                        {
                            CADApp.SetSystemVariable("dimldrblk", "_DOT");//引线
                            CADApp.SetSystemVariable("dimblk1", "_ARCHTICK");//第1个箭头
                            CADApp.SetSystemVariable("dimblk2", "_ARCHTICK");//第2个箭头
                                                                             //"_NONE" 无
                                                                             //"" 填充闭合
                                                                             //"_DOT" 点
                                                                             //"_DOTSMALL" 小点
                                                                             //"_DOTBLANK" 空心点
                                                                             //"_ORIGIN" 指示原点
                                                                             //"_ORIGIN2" 指示原点 2
                                                                             //"_OPEN" 打开
                                                                             //"_OPEN90" 直角
                                                                             //"_OPEN30" 30 度角
                                                                             //"_CLOSED" 闭合
                                                                             //"_SMALL" 空心小点
                                                                             //"_OBLIQUE" 倾斜
                                                                             //"_BOXFILLED" 填充方框
                                                                             //"_BOXBLANK" 方框
                                                                             //"_CLOSEDBLANK" 空心闭合
                                                                             //"_DATUMFILLED" 填充基准三角形
                                                                             //"_DATUMBLANK" 基准三角形
                                                                             //"_INTEGRAL" 积分
                                                                             //"_ARCHTICK" 建筑标记

                            // 创建新的标注样式
                            DimStyleTableRecord dimStyle = new DimStyleTableRecord()
                            {
                                Name = dim,
                                Dimsah = true,//打开时箭头才生效
                                Dimldrblk = blockTable["_DOT"],
                                Dimblk1 = blockTable["_ARCHTICK"],
                                Dimblk2 = blockTable["_ARCHTICK"],
                                Dimdec = 0,//小数位数

                                Dimtxsty = textStyleTable[text],//文字样式
                                Dimtxt = 3.5,//文字高度

                                Dimtad = 1,//文字位置，垂直//0 居中，1 上，2 外部，3 JIS，4 下
                                Dimjust = 0,//文字位置，水平//0 居中，1 第一条尺寸界线，2 第二条尺寸界线，3 第一条尺寸界线上方，4 第二条尺寸界线上方
                                Dimtxtdirection = false,//文字位置，观察方向//false 从左至右，true 从右至左
                                Dimgap = 1,//文字位置，从尺寸线偏移//

                                Dimdle = 0,//超出标记
                                Dimasz = 1,//箭头大小
                                Dimdli = 7.5,//基线间距
                                Dimexe = 2.5,//超出尺寸线
                                Dimexo = 3,//起点偏移量
                                Dimcen = 1.75,//圆心标记
                                Dimscale = last / first,//全局比例
                                Dimtp = 0,//上偏差
                                Dimtm = 0,//下偏差
                                Dimjogang = 0.7854,//折弯角度
                                Dimtfac = 1,//分数高度比例
                                Dimlfac = 1,//测量单位比例因子

                                //Dimtsz = 20,
                                //Dimtvp = 20,
                                //Dimrnd = 20,
                            };

                            dimStyleTable.Add(dimStyle);
                            transaction.AddNewlyCreatedDBObject(dimStyle, true);
                            CADApp.SetSystemVariable("dimldrblk", ".");//引线
                            CADApp.SetSystemVariable("dimblk1", ".");//第1个箭头
                            CADApp.SetSystemVariable("dimblk2", ".");//第2个箭头
                        }

                        string mleader = "LMleader-" + item;
                        if (!dBDictionary.Contains(mleader))
                        {
                            CADApp.SetSystemVariable("dimblk", "_DOT");//引线
                            MLeaderStyle newMleadStyle = new MLeaderStyle()
                            {
                                TextAttachmentType = TextAttachmentType.AttachmentBottomOfTopLine, //Text连接方式
                                ArrowSymbolId = blockTable["_DOT"],//箭头符号样式
                                ArrowSize = 1,
                                TextAngleType = TextAngleType.HorizontalAngle,//文字保持水平
                                EnableDogleg = false,//关闭基线
                                DoglegLength = 0,//基线距离
                                LandingGap = 0,//基线间隙
                                TextStyleId = textStyleTable["LText"],
                                TextHeight = 3.5,
                                Scale = last / first,
                            };
                            newMleadStyle.PostMLeaderStyleToDb(database, mleader);
                            transaction.AddNewlyCreatedDBObject(newMleadStyle, true);
                            CADApp.SetSystemVariable("dimblk", ".");//引线
                        }
                    }
                    transaction.Commit();
                }
            }
        }
    }
}