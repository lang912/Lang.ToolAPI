using Lang.ToolBiz.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lang.ToolBiz
{
    public class BPDF
    {
        /// <summary>
        /// word 转换成pdf
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public byte[] WordToPdf(Stream stream) 
        {
            var info= new PdfHelper().WordToPdf(stream);
            File.WriteAllBytes("d:/pdf.pdf", info);
            return info;
        }
    }
}
