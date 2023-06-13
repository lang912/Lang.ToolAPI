using DinkToPdf;
using System.Diagnostics;
using ColorMode = DinkToPdf.ColorMode;
using PaperKind = DinkToPdf.PaperKind;

namespace Lang.ToolBiz
{
    internal class PdfHelper
    {
        /// <summary>
        /// 转换类实例，此示例只能实例化一次，
        /// </summary>
        private static readonly SynchronizedConverter Converter = new SynchronizedConverter(new PdfTools());

        public byte[] HtmlToPdf(string htmlText)
        {
            var doc = new HtmlToPdfDocument()
            {
                Objects =
                {
                    new ObjectSettings()
                    {
                        PagesCount = true,
                        HtmlContent = htmlText,
                        WebSettings = { DefaultEncoding = "utf-8" },
                        HeaderSettings = { FontSize = 9, Right = "", Line = true, Spacing = 2.812}
                    }
                }
            };

            doc.GlobalSettings = new GlobalSettings()
            {
                ColorMode = ColorMode.Color,
                Orientation = Orientation.Portrait,
                PaperSize = PaperKind.A4
            };

            var bytes = Converter.Convert(doc);
            //FileStream fs = new FileStream($"D://{DateTime.Now.ToString("yyyyMMddHHmmssffff")}.pdf", FileMode.Create);
            //fs.Write(bytes, 0, bytes.Length);
            //fs.Dispose();
            return bytes;
        }

        public byte[] WordToPdf(Stream stream)
        {
            return WordToHtml(stream);
            //return HtmlToPdf(html);
        }

        private byte[] WordToHtml(Stream wordStream)
        {
            string output = "D:\\sssssss.pdf";
            MemoryStream outputStream = new MemoryStream();

            // 使用 LibreOffice 的命令行工具进行转换
            var processStartInfo = new ProcessStartInfo
            {
                FileName = "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
                Arguments = $"--headless --convert-to pdf \"{wordStream}\" --outdir \"{output}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (var process = new Process())
            {
                process.StartInfo = processStartInfo;
                process.OutputDataReceived += (sender, e) => Console.WriteLine($"[StdOut]: {e.Data}");
                process.ErrorDataReceived += (sender, e) => Console.WriteLine($"[StdErr]: {e.Data}");

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                process.WaitForExit();
            }

            outputStream.Position = 0;
            byte[] pdfBytes = new byte[outputStream.Length];
            outputStream.Read(pdfBytes, 0, pdfBytes.Length);

           
            outputStream.Close();
            outputStream.Dispose();
            // 打开转换后的 PDF 文件
            //System.Diagnostics.Process.Start(outputFilePath);
            return pdfBytes;
        }
    }
}
