using Lang.ToolBiz;
using Lang.ToolBiz.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace Lang.ToolAPI.Controllers
{
    /// <summary>
    /// pdf 控制器
    /// </summary>
    public class PDFController : BaseController
    {
        private BPDF _pdf;
        public PDFController(BPDF bPDF)
        {
            this._pdf = bPDF;
        }

        /// <summary>
        /// word转pdf
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult WordToPdf()
        {
            var files = HttpContext.Request.Form.Files;
            if (files == null || files.Count == 0)
            {
                throw new Exception("没有找到可转换的文件");
            }

            var byteinfo = this._pdf.WordToPdf(files[0].OpenReadStream());
            return File(byteinfo, "application/pdf", $"pdf文件{DateTime.Now.ToString("yyyyMMddHHmmss")}.pdf");
        }
    }
}
