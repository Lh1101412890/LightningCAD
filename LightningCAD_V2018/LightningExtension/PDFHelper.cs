using System.Collections.Generic;

using iText.Kernel.Pdf;

using Lightning.Extension;

namespace LightningCAD.LightningExtension
{
    public static class PDFHelper
    {
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