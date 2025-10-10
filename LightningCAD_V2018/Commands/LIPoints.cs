using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows;
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
using OpenFileDialog = Autodesk.AutoCAD.Windows.OpenFileDialog;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 导入坐标
    /// </summary>
    public class LIPoints : CommandBase
    {
        [CommandMethod(nameof(LIPoints))]
        public static void Command()
        {
            Document document = CADApp.DocumentManager.MdiActiveDocument;
            Editor editor = document.Editor;
            try
            {
                PromptKeywordOptions options1 = new PromptKeywordOptions("是否导入点高程？[是(Y)/否(N)/放弃(U)]")
                {
                    AppendKeywordsToMessage = true,
                };
                options1.Keywords.Add("Y");
                options1.Keywords.Add("N");
                options1.Keywords.Add("U");
                options1.Keywords.Default = "Y";
                PromptResult result1 = editor.GetKeywords(options1);
                if (result1.Status != PromptStatus.OK || result1.StringResult == "U")
                {
                    return;
                }
                string hasZ = result1.StringResult;

                OpenFileDialog openFile = new OpenFileDialog("选择坐标数据Excel文件", "Excel", "xlsx", "ExcelPath", OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder);
                if (openFile.ShowDialog() != DialogResult.OK)
                {
                    editor.WriteMessage("*取消*\n");
                    return;
                }

                OleDbConnection connection = new OleDbConnection(OleDbHelper.GetConnectionString(openFile.Filename, true, true));
                try
                {
                    connection.Open();
                }
                catch (OleDbException)
                {
                    MessageBox.Show("该文件已经被其他用户以独占方式打开，或者您没有查看和写入其数据的权限。", "文件打开失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    connection.Dispose(); return;
                }
                System.Data.DataTable dataTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                List<string> names = new List<string>();
                foreach (DataRow item in dataTable.Rows)
                {
                    names.Add(item["TABLE_NAME"].ToString());
                }

                string message = "请选择需导入的表[";
                for (int i = 0; i < names.Count; i++)
                {
                    message += $"{names[i]}({i + 1})/";
                }
                message += "放弃(U)]";

                PromptKeywordOptions options2 = new PromptKeywordOptions(message)
                {
                    AppendKeywordsToMessage = true,
                };
                for (int i = 0; i < names.Count; i++)
                {
                    options2.Keywords.Add((i + 1).ToString());
                }
                options2.Keywords.Add("U");
                options2.Keywords.Default = "1";
                PromptResult result2 = editor.GetKeywords(options2);
                if (result2.Status != PromptStatus.OK || result2.StringResult == "U")
                {
                    dataTable.Dispose();
                    connection.Close();
                    connection.Dispose();
                    return;
                }
                int n = int.Parse(result2.StringResult);
                OleDbCommand cmd = new OleDbCommand()
                {
                    CommandText = $"select * from[{names[n - 1]}]",
                    Connection = connection,
                };
                OleDbDataReader reader = cmd.ExecuteReader();

                List<DBPoint> points = new List<DBPoint>();
                bool v;
                //读取数据
                try
                {
                    while (reader.Read())
                    {
                        bool by = double.TryParse(reader["X"].ToString(), out double y);
                        bool bx = double.TryParse(reader["Y"].ToString(), out double x);
                        if (bx && by)
                        {
                            double z = 0;
                            if (hasZ == "Y")
                            {
                                _ = double.TryParse(reader["Z"].ToString(), out z);
                            }
                            points.Add(new DBPoint(new Point3d(x, y, z)));
                        }
                    }
                    v = true;
                }
                catch (System.IndexOutOfRangeException)
                {
                    editor.WriteMessage("未识别到有效的数据列标题（X/Y/Z）\n");
                    v = false;
                }
                finally
                {
                    reader.Close();
                    reader = null;
                    cmd.Dispose();
                    connection.Close();
                    connection.Dispose();
                }

                if (v)
                {
                    document.Drawing(points);
                    int i = points.Count;
                    if (i > 0)
                    {
                        editor.WriteMessage($"导入成功!共绘制 {i} 个点\n");
                    }
                    else
                    {
                        editor.WriteMessage("未找到点数据!\n");
                    }
                }
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }
    }
}