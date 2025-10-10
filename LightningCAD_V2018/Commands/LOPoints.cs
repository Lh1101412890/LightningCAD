using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using Lightning.Database;
using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using MessageBox = System.Windows.MessageBox;
using SaveFileDialog = Autodesk.AutoCAD.Windows.SaveFileDialog;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 导出坐标
    /// </summary>
    public class LOPoints : CommandBase
    {
        [CommandMethod(nameof(LOPoints), CommandFlags.UsePickSet)]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Editor editor = document.Editor;

                PromptSelectionOptions prompt1 = new PromptSelectionOptions() { MessageForAdding = "请选择需导出的点" };
                TypedValue[] typedValues =
                {
                    new TypedValue((int)DxfCode.Start,"point"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result1 = editor.GetSelection(prompt1, filter);
                if (result1.Status != PromptStatus.OK)
                {
                    return;
                }

                List<DBPoint> outPoints = document.GetObjects(result1.Value.GetObjectIds()).Cast<DBPoint>().ToList();

                PromptKeywordOptions keywordOptions = new PromptKeywordOptions("是否导出点高程？[是(Y)/否(N)/放弃(U)]")
                {
                    AppendKeywordsToMessage = true,
                };
                keywordOptions.Keywords.Add("Y");
                keywordOptions.Keywords.Add("N");
                keywordOptions.Keywords.Add("U");
                keywordOptions.Keywords.Default = "Y";
                PromptResult result = editor.GetKeywords(keywordOptions);
                if (result.Status != PromptStatus.OK || result.StringResult == "U")
                {
                    return;
                }

                // 用户确认坐标数据精度(小数位数)
                PromptKeywordOptions accKey = new PromptKeywordOptions("请输入数据精度[0(0)/0.0(1)/0.00(2)/0.000(3)/0.0000(4)/0.00000(5)/放弃(U)]")
                {
                    AppendKeywordsToMessage = true,
                };
                accKey.Keywords.Add("0");
                accKey.Keywords.Add("1");
                accKey.Keywords.Add("2");
                accKey.Keywords.Add("3");
                accKey.Keywords.Add("4");
                accKey.Keywords.Add("5");
                accKey.Keywords.Add("U");
                accKey.Keywords.Default = "3";
                PromptResult accResult = editor.GetKeywords(accKey);
                if (accResult.Status != PromptStatus.OK || accResult.StringResult == "U")
                {
                    return;
                }

                int acc = int.Parse(accResult.StringResult);

                SaveFileDialog saveFileDialog = new SaveFileDialog("输入文件名称", "点坐标", "xlsx", "导出坐标点文件", SaveFileDialog.SaveFileDialogFlags.DefaultIsFolder);
                DialogResult dialogResult = saveFileDialog.ShowDialog();

                if (dialogResult != DialogResult.OK)
                {
                    editor.WriteMessage("*取消*");
                    return;
                }
                // 保存文件路径
                string file = saveFileDialog.Filename;
                FileInfo saveFile = new FileInfo(file);
                FileInfo fileInfo = new FileInfo(Information.GetFileInfo("Files\\导入导出点.xlsx").FullName);
                try
                {
                    fileInfo.CopyTo(file, true);
                }
                catch (IOException exp)
                {
                    if (exp.Message.Contains("because it is being used by another process."))
                    {
                        editor.WriteMessage("指定的文件被打开，请关闭后重试！\n");
                        return;
                    }
                    else
                    {
                        editor.WriteMessage("IO异常\n");
                        exp.LogTo(Information.God);
                        return;
                    }
                }

                List<Point3d> points = new List<Point3d>(outPoints.Count);
                //转换为Point3d点
                foreach (DBPoint dBPoint in outPoints)
                {
                    points.Add(dBPoint.Position);
                }

                #region 对Excel文件进行数据库操作，写入数据
                OleDbConnection connection = new OleDbConnection(OleDbHelper.GetConnectionString(saveFile.FullName, true, true));
                OleDbCommand cmd = connection.CreateCommand();
                double x;
                double y;
                double z;
                int num;

                //提示是否安装AccessDatabaseEngine_64
                try
                {
                    connection.Open();
                }
                catch (InvalidOperationException e)
                {
                    if (e.Message == "未在本地计算机上注册“Microsoft.ACE.OLEDB.12.0”提供程序。")
                    {
                        MessageBox.Show(e.Message);
                        return;
                    }
                }

                for (int i = 0; i < points.Count; i++)
                {
                    num = i + 1;

                    x = Math.Round(points[i].Y, acc);//坐标和软件是反的
                    y = Math.Round(points[i].X, acc);//

                    if (result.StringResult is "Y")//根据用户选择是否保存点高程
                    {
                        z = Math.Round(points[i].Z, acc);
                        cmd.CommandText = $"insert into [Sheet1$](序号,X,Y,Z) values({num},{x},{y},{z})";
                    }
                    else
                    {
                        cmd.CommandText = $"insert into [Sheet1$](序号,X,Y) values({num},{x},{y})";
                    }
                    cmd.ExecuteNonQuery();
                }
                cmd.Dispose();
                connection.Close();
                connection.Dispose();

                #endregion

                //提示操作成功
                editor.WriteMessage($"导出成功,共导出 {points.Count} 个点!文件路径：{file}\n");
                Process.Start("Explorer", file);
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
    }
}