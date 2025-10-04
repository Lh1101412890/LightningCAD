using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Lightning.Database;
using Lightning.Extension;

using LightningCAD.Extension;
using LightningCAD.LightningExtension;

using CADApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace LightningCAD.Commands
{
    /// <summary>
    /// 计算长度和面积
    /// </summary>
    public class LCGeometry : CommandBase
    {
        //数据计算精度
        private static readonly int acc = 4;

        [CommandMethod(nameof(LCGeometry), CommandFlags.UsePickSet)]
        public static void Command()
        {
            try
            {
                Document document = CADApp.DocumentManager.MdiActiveDocument;
                Database database = document.Database;
                Editor editor = document.Editor;
                PromptSelectionOptions prompt1 = new PromptSelectionOptions() { MessageForAdding = "请选择图形" };
                TypedValue[] typedValues =
                {
                    new TypedValue((int)DxfCode.Operator,"<or"),
                    new TypedValue((int)DxfCode.Start,"line"),
                    new TypedValue((int)DxfCode.Start,"lwpolyline"),
                    new TypedValue((int)DxfCode.Start,"circle"),
                    new TypedValue((int)DxfCode.Start,"ellipse"),
                    new TypedValue((int)DxfCode.Start,"arc"),
                    new TypedValue((int)DxfCode.Start,"spline"),
                    new TypedValue((int)DxfCode.Operator,"or>"),
                };
                SelectionFilter filter = new SelectionFilter(typedValues);
                PromptSelectionResult result1 = editor.GetSelection(prompt1, filter);
                if (result1.Status != PromptStatus.OK)
                {
                    return;
                }
                List<DBObject> objects = document.GetObjects(result1.Value.GetObjectIds());

                List<Line> lines = objects.OfType<Line>().ToList(); //读取所有直线
                List<Polyline> polylines = objects.OfType<Polyline>().ToList(); //读取所有多段线
                List<Circle> circles = objects.OfType<Circle>().ToList(); //读取所有圆
                List<Ellipse> ellipses = objects.OfType<Ellipse>().ToList(); //读取所有椭圆
                List<Arc> arcs = objects.OfType<Arc>().ToList(); //读取所有圆弧
                List<Spline> splines = objects.OfType<Spline>().ToList(); //读取所有样条曲线

                double length = 0;
                double aera = 0;

                double lineLen = 0;
                double polylineLength = 0;
                double polylineAera = 0;
                double circleLength = 0;
                double circleAera = 0;
                double ellipseLength = 0;
                double ellipseAera = 0;
                double splineLength = 0;
                double splineAera = 0;

                List<CalculateLengthAreaModel> models = new List<CalculateLengthAreaModel>();

                #region 统计
                foreach (Line line in lines) //统计直线
                {
                    double temp = Math.Round(line.Length, acc);
                    lineLen += temp;
                    length += temp;
                    models.Add(new CalculateLengthAreaModel() { Type = "直线", Length = temp, Area = 0 });
                }
                foreach (Polyline polyline in polylines) //统计多段线
                {
                    double tempLength = Math.Round(polyline.Length, acc);
                    polylineLength += tempLength;
                    length += tempLength;
                    if (polyline.Closed)
                    {
                        double tempAera = Math.Round(polyline.Area, acc);
                        polylineAera += tempAera;
                        aera += tempAera;
                        models.Add(new CalculateLengthAreaModel() { Type = "多段线", Length = tempLength, Area = tempAera });
                    }
                    else
                    {
                        models.Add(new CalculateLengthAreaModel() { Type = "多段线", Length = tempLength, Area = 0 });
                    }
                }
                foreach (Circle circle in circles) //统计圆
                {
                    double tempLength = Math.Round(circle.Circumference, acc);
                    circleLength += tempLength;
                    length += tempLength;
                    double tempAera = Math.Round(circle.Area, acc);
                    circleAera += tempAera;
                    aera += tempAera;
                    models.Add(new CalculateLengthAreaModel() { Type = "圆", Length = tempLength, Area = tempAera });
                }
                foreach (Arc arc in arcs) //统计圆弧
                {
                    double tempLength = Math.Round(arc.Length, acc);
                    circleLength += tempLength;
                    length += tempLength;
                    models.Add(new CalculateLengthAreaModel() { Type = "圆弧", Length = tempLength, Area = 0 });
                }
                foreach (Ellipse ellipse in ellipses) //统计椭圆
                {
                    double longRadius = ellipse.MajorRadius;
                    double shortRadius = ellipse.MinorRadius;
                    double circumference = Math.Round(2 * Math.PI * Math.Sqrt((Math.Pow(longRadius, 2) + Math.Pow(shortRadius, 2)) / 2), acc);
                    ellipseLength += circumference;
                    length += circumference;
                    if (ellipse.Closed)
                    {
                        double tempAera = Math.Round(ellipse.Area, acc);
                        ellipseAera += tempAera;
                        aera += tempAera;
                        models.Add(new CalculateLengthAreaModel() { Type = "椭圆", Length = circumference, Area = tempAera });
                    }
                    else
                    {
                        models.Add(new CalculateLengthAreaModel() { Type = "椭圆弧", Length = circumference, Area = 0 });
                    }
                }
                foreach (Spline spline in splines) //统计样条曲线
                {
                    //在结束点参数位置获取样条曲线长度，同cad中list命令中的长度一致
                    double tempLength = Math.Round(spline.GetDistanceAtParameter(spline.EndParam), acc);
                    splineLength += tempLength;
                    length += tempLength;
                    if (spline.Closed)
                    {
                        double tempAera = Math.Round(spline.Area, acc);
                        splineAera += tempAera;
                        aera += tempAera;
                        models.Add(new CalculateLengthAreaModel() { Type = "样条曲线", Length = tempLength, Area = tempAera });
                    }
                    else
                    {
                        models.Add(new CalculateLengthAreaModel() { Type = "样条曲线", Length = tempLength, Area = 0 });
                    }
                }
                #endregion

                length = Math.Round(length, acc);
                aera = Math.Round(aera, acc);
                editor.WriteMessage($"计算完成!\n长度：{length}\n面积：{aera}\n");

                PromptKeywordOptions prompt2 = new PromptKeywordOptions("是否导出明细？[是(Y)/否(N)]")
                {
                    AppendKeywordsToMessage = true,
                };
                prompt2.Keywords.Add("Y");
                prompt2.Keywords.Add("N");
                prompt2.Keywords.Default = "N";
                PromptResult result2 = editor.GetKeywords(prompt2);
                LightningApp.ShowMsg("\"F2\" 快速查看复制结果", 3);

                if (result2.Status != PromptStatus.OK)
                {
                    return;
                }
                string command = result2.StringResult;
                if (command == "N") return;

                var model1 = models.Where(x => x.Type == "直线").ToList();
                var model2 = models.Where(x => x.Type == "多段线").Where(x => x.Area == 0).ToList();
                var model3 = models.Where(x => x.Type == "多段线").Where(x => x.Area != 0).ToList();
                var model4 = models.Where(x => x.Type == "圆").ToList();
                var model5 = models.Where(x => x.Type == "圆弧").ToList();
                var model6 = models.Where(x => x.Type == "椭圆").ToList();
                var model7 = models.Where(x => x.Type == "椭圆弧").ToList();
                var model8 = models.Where(x => x.Type == "样条曲线").Where(x => x.Area == 0).ToList();
                var model9 = models.Where(x => x.Type == "样条曲线").Where(x => x.Area != 0).ToList();
                models.Clear();
                models.AddRange(model1);
                models.AddRange(model2);
                models.AddRange(model3);
                models.AddRange(model4);
                models.AddRange(model5);
                models.AddRange(model6);
                models.AddRange(model7);
                models.AddRange(model8);
                models.AddRange(model9);

                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                DateTime now = DateTime.Now;
                string time = $"{now.Year}-{now.Month}-{now.Day}_{now.Hour}{now.Minute}{now.Second}";
                string file = $"{desktop}\\计算S L明细{time}.xlsx";
                File.Copy(Information.GetFileInfo("Files\\计算S L明细.xlsx").FullName, file, true);

                #region 导出明细数据至Excel
                OleDbConnection connection = new OleDbConnection(OleDbHelper.GetConnectionString(file, true, true));
                OleDbCommand cmd = connection.CreateCommand();

                long num;

                try
                {
                    connection.Open();
                }
                catch (InvalidOperationException exp)
                {
                    if (exp.Message == "未在本地计算机上注册“Microsoft.ACE.OLEDB.12.0”提供程序。")
                    {
                        exp.Record();
                        return;
                    }
                }

                lineLen = Math.Round(lineLen, acc);
                polylineLength = Math.Round(polylineLength, acc);
                polylineAera = Math.Round(polylineAera, acc);
                circleLength = Math.Round(circleLength, acc);
                circleAera = Math.Round(circleAera, acc);
                ellipseLength = Math.Round(ellipseLength, acc);
                ellipseAera = Math.Round(ellipseAera, acc);
                splineLength = Math.Round(splineLength, acc);
                splineAera = Math.Round(splineAera, acc);
                cmd.CommandText = $"insert into [Sheet1$](序号,类型,长度,面积,总长度,总面积,直线长度,多段线长度,多段线面积,[圆(弧)周长],圆面积,[椭圆(弧)周长],椭圆面积,样条曲线长度,样条曲线面积) values(1,'{models[0].Type}',{models[0].Length},{models[0].Area},{length},{aera},{lineLen},{polylineLength},{polylineAera},{circleLength},{circleAera},{ellipseLength},{ellipseAera},{splineLength},{splineAera})";

                cmd.ExecuteNonQuery();

                for (int i = 1; i < models.Count; i++)
                {
                    num = i + 1;
                    cmd.CommandText = $"insert into [Sheet1$](序号,类型,长度,面积) values({num},'{models[i].Type}',{models[i].Length},{models[i].Area})";
                    cmd.ExecuteNonQuery();
                }

                cmd.Dispose();
                connection.Close();
                connection.Dispose();

                #endregion

                editor.WriteMessage($"导出成功!文件路径：{file}");
                Process.Start("Explorer", file);
            }
            catch (System.Exception exp)
            {
                exp.Record();
            }
        }

        public class CalculateLengthAreaModel
        {
            private double length;
            private double area;

            public string Type { get; set; }
            public double Length { get => Math.Round(length, acc); set => length = value; }
            public double Area { get => Math.Round(area, acc); set => area = value; }
        }
    }
}