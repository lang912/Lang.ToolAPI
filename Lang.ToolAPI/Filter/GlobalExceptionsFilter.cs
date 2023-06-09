using Lang.ToolBiz.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace Lang.ToolAPI.Filter
{
    public class GlobalExceptionsFilter : IExceptionFilter
    {
        public void OnException(ExceptionContext context)
        {
            context.Result = new JsonResult(new Result() { Msg = $"系统异常：{context.Exception?.Message}" });
            NLog.LogManager.GetCurrentClassLogger().Error($"系统异常：{context.Exception}");
        }
    }
}
