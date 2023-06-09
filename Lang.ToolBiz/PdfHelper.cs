using DinkToPdf;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System.Text;
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
            var html = WordToHtml(stream);
            return HtmlToPdf(html);
        }

        private string WordToHtml(Stream stream)
        {
            var myDocx = new XWPFDocument(stream); //打开07（.docx）以上的版本的文档
            var picInfoList = PicturesHandleAsync(myDocx);

            var sb = new StringBuilder();

            foreach (var para in myDocx.BodyElements)
                switch (para.ElementType)
                {
                    case BodyElementType.PARAGRAPH:
                        {
                            var paragraph = (XWPFParagraph)para;
                            sb.Append(ParaGraphHandle(paragraph, picInfoList));

                            break;
                        }

                    case BodyElementType.TABLE:
                        var paraTable = (XWPFTable)para;
                        sb.Append(TableHandle(paraTable, picInfoList));
                        break;
                }


            return sb.Replace(" style=''", "").ToString();
        }

        /// <summary>
        ///     图片处理
        /// </summary>
        /// <param name="myDocx"></param>
        /// <param name="isImgUploadAliYun">图片是否上传阿里云</param>
        /// <returns></returns>
        public List<PicInfo> PicturesHandleAsync(XWPFDocument myDocx, bool isImgUploadAliYun = false)
        {
            var picInfoList = new List<PicInfo>();
            var picturesList = myDocx.AllPictures;
            foreach (var pictures in picturesList)
            {
                var pData = pictures.Data;
                var picPackagePart = pictures.GetPackagePart();
                var picPackageRelationship = pictures.GetPackageRelationship();
                var picInfo = new PicInfo
                {
                    Id = picPackageRelationship.Id,
                    PicType = picPackagePart.ContentType
                };

                if (string.IsNullOrWhiteSpace(picInfo.Url))
                    picInfo.Url = $"data:{picInfo.PicType};base64,{Convert.ToBase64String(pData)}";
                //先把pData传阿里云得到url  如果有其他方式传改这里 或者转base64

                picInfoList.Add(picInfo);
            }

            return picInfoList;
        }

        /// <summary>
        ///     word中的表格处理
        /// </summary>
        /// <param name="paraTable"></param>
        /// <param name="picInfoList"></param>
        /// <returns></returns>
        public StringBuilder TableHandle(XWPFTable paraTable, List<PicInfo> picInfoList)
        {
            var sb = new StringBuilder();

            var rows = paraTable.Rows;
            sb.Append("<table border='1' cellspacing='0'>");
            foreach (var row in rows)
            {
                var cells = row.GetTableCells();

                sb.Append(
                    "<tr style='");
                //var firstRowCell = cells[0];


                sb.Append(
                    "'>");


                foreach (var cell in cells)
                {
                    var cellCtTc = cell.GetCTTc();
                    var tcPr = cellCtTc.tcPr;


                    sb.Append("<td style='");

                    if (!string.IsNullOrWhiteSpace(tcPr.tcW?.w))
                        sb.Append($"width:{tcPr.tcW.w}px;");
                    if (!string.IsNullOrWhiteSpace(tcPr.shd?.fill))
                        sb.Append($"background-color: #{tcPr.shd.fill};");

                    sb.Append("'>");
                    var cellParagraphs = cell.Paragraphs;
                    foreach (var cellParagraph in cellParagraphs)
                        sb.Append(ParaGraphHandle(cellParagraph, picInfoList));

                    //sb.Append(cell.GetText());
                    sb.Append("</td>");
                }


                sb.Append("</tr>");
            }

            sb.Append("</table>");
            return sb;
        }

        /// <summary>
        ///     word文本对应处理
        /// </summary>
        /// <param name="ctr"></param>
        /// <returns></returns>
        public StringBuilder FontHandle(CT_R ctr)
        {
            var sb = new StringBuilder();

            #region 文本格式

            var textList = ctr.GetTList();
            foreach (var text in textList)
            {
                sb.Append(
                    "<span style='");
                if (!string.IsNullOrWhiteSpace(ctr.rPr?.color?.val))
                    sb.Append(
                        $"color:#{ctr.rPr.color.val};");
                if (!string.IsNullOrWhiteSpace(ctr.rPr?.highlight?.val.ToString()))
                    sb.Append(
                        $"background-color: {ctr.rPr.highlight.val};");
                if (ctr.rPr?.i?.val == true)
                    sb.Append(
                        "font-style:italic;");
                if (ctr.rPr?.b?.val == true)
                    sb.Append(
                        "font-weight:bold;");
                if (ctr.rPr?.sz != null)
                    sb.Append(
                        $"font-size:{ctr.rPr.sz.val}px;");
                if (!string.IsNullOrWhiteSpace(ctr.rPr?.rFonts?.ascii))
                    sb.Append(
                        $"font-family:{ctr.rPr.rFonts.ascii};");

                sb.Append(
                    "'>");

                sb.Append(text.Value);
                sb.Append("</span>");
            }

            #endregion

            return sb;
        }



        /// <summary>
        ///     word图片对应处理
        /// </summary>
        /// <param name="ctr"></param>
        /// <param name="picInfoList"></param>
        /// <returns></returns>
        public StringBuilder DrawingHandle(CT_R ctr, List<PicInfo> picInfoList)
        {
            var sb = new StringBuilder();
            var drawingList = ctr.GetDrawingList();
            foreach (var drawing in drawingList)
            {
                var a = drawing.GetInlineList();
                foreach (var a1 in a)
                {
                    var anyList = a1.graphic.graphicData.Any;

                    foreach (var any1 in anyList)
                    {
                        var pictures = picInfoList
                            .FirstOrDefault(x =>
                                any1.IndexOf("a:blip r:embed=\"" + x.Id + "\"", StringComparison.Ordinal) > -1);
                        if (pictures != null && !string.IsNullOrWhiteSpace(pictures.Url))
                            sb.Append($@"<img src='{pictures.Url}' />");
                    }
                }
            }

            return sb;
        }

        /// <summary>
        ///     word行处理为P标签
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        public StringBuilder TagPHandle(XWPFParagraph paragraph)
        {
            var sb = new StringBuilder();
            sb.Append("<p style='");

            try
            {
                //左右对齐

                var fontAlignment = paragraph.FontAlignment;
                string fontAlignmentName;
                switch (fontAlignment)
                {
                    case 0:
                        fontAlignmentName = "auto";
                        break;
                    case 1:
                        fontAlignmentName = "left";
                        break;
                    case 2:
                        fontAlignmentName = "center";
                        break;
                    case 3:
                        fontAlignmentName = "right";
                        break;
                    default:
                        fontAlignmentName = "auto";
                        break;
                }
                //自动和左对齐不需样式
                if (fontAlignment > 1) sb.Append($"text-align:{fontAlignmentName};");


                var em = paragraph.IndentationFirstLine / 240;

                if (em > 0) sb.Append($"text-indent:{em}em;");
            }
            catch (Exception)
            {
                // ignored
            }

            sb.Append("'>");
            return sb;
        }

        /// <summary>
        ///     word文档对应行内容处理
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="picInfoList"></param>
        /// <returns></returns>
        public StringBuilder ParaGraphHandle(XWPFParagraph paragraph, List<PicInfo> picInfoList)
        {
            var sb = new StringBuilder();

            #region P标签

            sb.Append(TagPHandle(paragraph));

            #endregion


            var runs = paragraph.Runs;
            foreach (var run in runs)
            {
                var ctr = run.GetCTR();

                #region 图片格式

                sb.Append(DrawingHandle(ctr, picInfoList));

                #endregion

                #region 文本格式

                sb.Append(FontHandle(ctr));

                #endregion
            }

            sb.Append("</p>");
            return sb;
        }

        public class PicInfo
        {
            /// <summary>
            ///     图片编号
            /// </summary>
            public string Id { get; set; }

            /// <summary>
            ///     图片类型
            /// </summary>
            public string PicType { get; set; }

            /// <summary>
            ///     上传地址/或者Base64
            /// </summary>
            public string Url { get; set; }
        }
    }
}
