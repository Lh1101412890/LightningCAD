using System;
using System.Collections.Generic;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

using iText.Kernel.Pdf;

using Lightning.Extension;

using LightningCAD.Extension;

namespace LightningCAD.LightningExtension
{
    public static class Tools
    {
        /// <summary>
        /// 判断两条直线是否平行有交集
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static bool IsParallelUnion(Line line1, Line line2)
        {
            if (line1.IsVertical())
            {
                if (line2.IsVertical())
                {
                    double y1 = Math.Min(line1.StartPoint.Y, line1.EndPoint.Y);
                    double y2 = Math.Max(line1.StartPoint.Y, line1.EndPoint.Y);
                    double y3 = Math.Min(line2.StartPoint.Y, line2.EndPoint.Y);
                    double y4 = Math.Max(line2.StartPoint.Y, line2.EndPoint.Y);
                    if ((y1 <= y3 && y3 <= y2)
                        || (y1 <= y4 && y4 <= y2)
                        || (y3 <= y1 && y2 <= y4))
                    {
                        return true;
                    }
                }
            }
            else if (line1.IsHorizontal())
            {
                if (line2.IsHorizontal())
                {
                    double x1 = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double x2 = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double x3 = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double x4 = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    if ((x1 <= x3 && x3 <= x2)
                        || (x1 <= x4 && x4 <= x2)
                        || (x3 <= x1 && x2 <= x4))
                    {
                        return true;
                    }
                }
            }
            else
            {
                double k1 = line1.GetK();
                double length = line1.GetDistance(line2);
                if (length >= 0)
                {
                    double b1 = line1.GetB();
                    double b2 = line2.GetB();
                    double x1 = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double x2 = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double x3 = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double x4 = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    double b = Math.Abs(b1 - b2);
                    double c = Math.Sqrt(b * b - length * length);
                    double p = 0.5 * (b + length + c);
                    double s = Math.Sqrt(p * (p - b) * (p - length) * (p - c));
                    double x = s * 2 / b;
                    if ((k1 > 0 && b1 > b2) || (k1 < 0 && b1 < b2))
                    {
                        x3 -= 2 * x;
                        x4 -= 2 * x;
                    }
                    else
                    {
                        x3 += 2 * x;
                        x4 += 2 * x;
                    }
                    if ((x1 <= x3 && x3 <= x2)
                        || (x1 <= x4 && x4 <= x2)
                        || (x3 <= x1 && x2 <= x4))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 判断两条直线是否可以延长
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static bool TryExtend(Line line1, Line line2, out Line line, double acc = LightningTolerance.Global)
        {
            line = null;
            if (line1.IsHorizontal())
            {
                if (line2.IsHorizontal() && Math.Abs(line1.GetDistance(line2)) < acc)
                {
                    double l1min = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double l1max = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double l2min = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double l2max = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    if ((l1min - acc <= l2min && l2min <= l1max + acc) || (l1min - acc <= l2max && l2max <= l1max + acc) || (l2min - acc <= l1min && l1max <= l2max + acc))
                    {
                        double min = Math.Min(l1min, l2min);
                        double max = Math.Max(l1max, l2max);
                        double y = line1.StartPoint.Y;
                        line = new Line(new Point3d(min, y, 0), new Point3d(max, y, 0));
                        return true;
                    }
                }
            }
            else if (line1.IsVertical())
            {
                if (line2.IsVertical() && Math.Abs(line1.GetDistance(line2)) < acc)
                {
                    double l1min = Math.Min(line1.StartPoint.Y, line1.EndPoint.Y);
                    double l1max = Math.Max(line1.StartPoint.Y, line1.EndPoint.Y);
                    double l2min = Math.Min(line2.StartPoint.Y, line2.EndPoint.Y);
                    double l2max = Math.Max(line2.StartPoint.Y, line2.EndPoint.Y);
                    if ((l1min - acc <= l2min && l2min <= l1max + acc) || (l1min - acc <= l2max && l2max <= l1max + acc) || (l2min - acc <= l1min && l1max <= l2max + acc))
                    {
                        double min = Math.Min(l1min, l2min);
                        double max = Math.Max(l1max, l2max);
                        double x = line1.StartPoint.X;
                        line = new Line(new Point3d(x, min, 0), new Point3d(x, max, 0));
                        return true;
                    }
                }
            }
            else
            {
                double k1 = line1.GetK();
                double b1 = line1.GetB();
                double k2 = line2.GetK();
                double b2 = line2.GetB();
                if ((Math.Abs(k1 - k2) < 0.01) && Math.Abs(b1 - b2) < acc)
                {
                    double l1min = Math.Min(line1.StartPoint.X, line1.EndPoint.X);
                    double l1max = Math.Max(line1.StartPoint.X, line1.EndPoint.X);
                    double l2min = Math.Min(line2.StartPoint.X, line2.EndPoint.X);
                    double l2max = Math.Max(line2.StartPoint.X, line2.EndPoint.X);
                    if ((l1min - acc <= l2min && l2min <= l1max + acc) || (l1min - acc <= l2max && l2max <= l1max + acc) || (l2min - acc <= l1min && l1max <= l2max + acc))
                    {
                        double min = Math.Min(l1min, l2min);
                        double max = Math.Max(l1max, l2max);
                        line = new Line(new Point3d(min, k1 * min + b1, 0), new Point3d(max, k1 * max + b1, 0));
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 获取两条直线的中心线
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <param name="line"></param>
        /// <returns></returns>
        public static bool GetCenterLine(Line line1, Line line2, out Line line)
        {
            line = null;
            double max;
            double min;
            //斜率k为无穷
            if (line1.IsVertical())
            {
                if (!line2.IsVertical())
                {
                    return false;
                }
                max = Math.Max(Math.Max(Math.Max(line1.StartPoint.Y, line1.EndPoint.Y), line2.StartPoint.Y), line2.EndPoint.Y);
                min = Math.Min(Math.Min(Math.Min(line1.StartPoint.Y, line1.EndPoint.Y), line2.StartPoint.Y), line2.EndPoint.Y);
                double x = (line1.StartPoint.X + line1.EndPoint.X + line2.StartPoint.X + line2.EndPoint.X) / 4;
                line = new Line(new Point3d(x, min, 0), new Point3d(x, max, 0));
            }
            //斜率k为0
            else if (line1.IsHorizontal())
            {
                if (!line2.IsHorizontal())
                {
                    return false;
                }
                max = Math.Max(Math.Max(Math.Max(line1.StartPoint.X, line1.EndPoint.X), line2.StartPoint.X), line2.EndPoint.X);
                min = Math.Min(Math.Min(Math.Min(line1.StartPoint.X, line1.EndPoint.X), line2.StartPoint.X), line2.EndPoint.X);
                double y = (line1.StartPoint.Y + line1.EndPoint.Y + line2.StartPoint.Y + line2.EndPoint.Y) / 4;
                line = new Line(new Point3d(min, y, 0), new Point3d(max, y, 0));
            }
            else
            {
                double k1 = line1.GetK();
                double k2 = line2.GetK();
                if (Math.Abs(k1 - k2) >= 0.2)
                {
                    return false;
                }
                double b1 = line1.GetB();
                double b2 = line2.GetB();

                double k3 = -1 / k1;
                double b_b = (b1 + b2) / 2;
                double b_1 = line1.StartPoint.Y - k3 * line1.StartPoint.X;
                double b_2 = line1.EndPoint.Y - k3 * line1.EndPoint.X;
                double b_3 = line2.StartPoint.Y - k3 * line2.StartPoint.X;
                double b_4 = line2.EndPoint.Y - k3 * line2.EndPoint.X;
                double x1 = (b_1 - b_b) / (k1 - k3);
                double x2 = (b_2 - b_b) / (k1 - k3);
                double x3 = (b_3 - b_b) / (k1 - k3);
                double x4 = (b_4 - b_b) / (k1 - k3);

                min = Math.Min(Math.Min(Math.Min(x1, x2), x3), x4);
                max = Math.Max(Math.Max(Math.Max(x1, x2), x3), x4);
                line = new Line(new Point3d(min, k1 * min + b_b, 0), new Point3d(max, k1 * max + b_b, 0));
            }
            return true;
        }

        /// <summary>
        /// 合并pdf文件
        /// </summary>
        /// <param name="pdfFiles">需要合并的pdf文件</param>
        /// <param name="outputPdfPath">输出路径</param>
        public static void JoinFiles(List<string> pdfFiles, string outputPdfPath)
        {
            try
            {
                using (PdfWriter writer = new PdfWriter(outputPdfPath, new WriterProperties()))
                {
                    using (PdfDocument outDoc = new PdfDocument(writer))
                    {
                        foreach (var item in pdfFiles)
                        {
                            PdfReader reader = new PdfReader(item);
                            PdfDocument document = new PdfDocument(reader);
                            int n = document.GetNumberOfPages();
                            List<int> pages = new List<int>();
                            for (int i = 0; i < n; i++)
                            {
                                pages.Add(i + 1);
                            }
                            document.CopyPagesTo(pages, outDoc);
                            document.Close();
                            reader.Close();
                        }
                        outDoc.Close();
                    }
                    writer.Close();
                }
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
        }

        /// <summary>
        /// 拆分pdf文件
        /// </summary>
        /// <param name="pdfPath">被拆分文档</param>
        /// <param name="outputPdfPath">输出文件夹</param>
        public static void Split(string pdfPath, string outputPdfPath)
        {
            PdfReader reader = null;
            PdfDocument srcDoc = null;
            try
            {
                reader = new PdfReader(pdfPath);
                srcDoc = new PdfDocument(reader);

                int totalPages = srcDoc.GetNumberOfPages();
                for (int i = 1; i <= totalPages; i++)
                {
                    string singlePagePath = $"{outputPdfPath}\\{System.IO.Path.GetFileNameWithoutExtension(outputPdfPath)}_{i}.pdf";

                    PdfWriter writer = new PdfWriter(singlePagePath, new WriterProperties());
                    using (PdfDocument destDoc = new PdfDocument(writer))
                    {
                        srcDoc.CopyPagesTo(i, i, destDoc);
                    }
                }
            }
            catch (System.Exception exp)
            {
                exp.LogTo(Information.God);
            }
            finally
            {
                srcDoc?.Close();
                reader?.Close();
            }
        }

    }
}