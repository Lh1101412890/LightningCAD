using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Lightning.Extension;

using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    public class LFVersion : CommandBase
    {
        private static Editor editor;

        [CommandMethod(nameof(LFVersion))]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                editor = document.Editor;
                PromptKeywordOptions versionKey = new PromptKeywordOptions("请选择目标版本[2004(1)/2007(2)/2010(3)/2013(4)/放弃(U)]")
                {
                    AppendKeywordsToMessage = true,
                };
                versionKey.Keywords.Add("1");
                versionKey.Keywords.Add("2");
                versionKey.Keywords.Add("3");
                versionKey.Keywords.Add("4");
                versionKey.Keywords.Add("U");
                versionKey.Keywords.Default = "3";
                PromptResult versionResult = editor.GetKeywords(versionKey);
                if (versionResult.Status != PromptStatus.OK || versionResult.StringResult == "U")
                {
                    return;
                }
                DwgVersion dwgVersion;
                switch (versionResult.StringResult)
                {
                    case "1":
                        dwgVersion = DwgVersion.AC1800;
                        break;
                    case "2":
                        dwgVersion = DwgVersion.AC1021;
                        break;
                    case "3":
                        dwgVersion = DwgVersion.AC1024;
                        break;
                    case "4":
                        dwgVersion = DwgVersion.AC1027;
                        break;
                    default:
                        dwgVersion = DwgVersion.AC1024;
                        break;
                }
                FolderBrowserDialog folder = new FolderBrowserDialog()
                {
                    ShowNewFolderButton = false
                };
                if (folder.ShowDialog() != DialogResult.OK) { return; }

                PromptKeywordOptions deepthKey = new PromptKeywordOptions("是否转换所有子文件夹中的文件[是(Y)/否(N)/放弃(U)]")
                {
                    AppendKeywordsToMessage = true,
                };
                deepthKey.Keywords.Add("Y");
                deepthKey.Keywords.Add("N");
                deepthKey.Keywords.Add("U");
                deepthKey.Keywords.Default = "Y";
                PromptResult deepthResult = editor.GetKeywords(deepthKey);
                if (deepthResult.Status != PromptStatus.OK || deepthResult.StringResult == "U")
                {
                    return;
                }

                string selectedPath = folder.SelectedPath;
                if (deepthResult.StringResult == "N")
                {
                    Do1(selectedPath, dwgVersion);
                }
                else
                {
                    List<string> paths = Do2(selectedPath, dwgVersion);
                    if (paths.Count != 0)
                    {
                        string str = string.Join("\n", paths);
                        editor.WriteMessage($"读取下列文件时失败（文件可能被占用）：\n{str}\n");
                    }
                }
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }

        private static void Do1(string path, DwgVersion dwgVersion)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            List<string> paths = new List<string>();
            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                if (file.Extension == ".dwg")
                {
                    using (Database database = new Database(false, true))
                    {
                        try
                        {
                            database.ReadDwgFile(file.FullName, FileOpenMode.OpenForReadAndReadShare, true, "");
                            database.CloseInput(true);
                        }
                        catch (System.Exception exp)
                        {
                            paths.Add(file.FullName);
                            exp.LogTo(Information.God);
                            continue;
                        }
                        string ver;
                        switch (dwgVersion)
                        {
                            case DwgVersion.AC1800:
                                ver = "2004";
                                break;
                            case DwgVersion.AC1021:
                                ver = "2007";
                                break;
                            case DwgVersion.AC1024:
                                ver = "2010";
                                break;
                            case DwgVersion.AC1027:
                                ver = "2013";
                                break;
                            default:
                                ver = "2010";
                                break;
                        }
                        string name = file.FullName.Substring(0, file.FullName.Length - 4);
                        string newName = $"{name}_{ver}.dwg";
                        database.SaveAs(newName, dwgVersion);
                        editor.WriteMessage($"文件转换成功：{newName}\n");
                    }
                }
            }
            string str = string.Join("\n", paths);
            editor.WriteMessage($"读取下列文件时失败（文件可能被占用）：\n{str}\n");
        }

        private static List<string> Do2(string path, DwgVersion dwgVersion)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(path);
            List<string> paths = new List<string>();

            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                if (file.Extension == ".dwg")
                {
                    using (Database database = new Database(false, true))
                    {
                        try
                        {
                            database.ReadDwgFile(file.FullName, FileOpenMode.OpenForReadAndReadShare, true, "");
                            database.CloseInput(true);
                        }
                        catch (System.Exception exp)
                        {
                            paths.Add(file.FullName);
                            exp.LogTo(Information.God);
                            continue;
                        }
                        string ver;
                        switch (dwgVersion)
                        {
                            case DwgVersion.AC1800:
                                ver = "2004";
                                break;
                            case DwgVersion.AC1021:
                                ver = "2007";
                                break;
                            case DwgVersion.AC1024:
                                ver = "2010";
                                break;
                            case DwgVersion.AC1027:
                                ver = "2013";
                                break;
                            default:
                                ver = "2010";
                                break;
                        }
                        string name = file.FullName.Substring(0, file.FullName.Length - 4);
                        string newName = $"{name}_{ver}.dwg";
                        database.SaveAs(newName, dwgVersion);
                        editor.WriteMessage($"文件转换成功：{newName}\n");
                    }
                }
            }

            foreach (DirectoryInfo directory in directoryInfo.GetDirectories())
            {
                paths.AddRange(Do2(directory.FullName, dwgVersion));
            }
            return paths;
        }
    }
}