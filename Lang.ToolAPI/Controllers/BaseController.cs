using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace Lang.ToolAPI.Controllers
{
    [Route("ToolAPI/[controller]/[action]")]
    [ApiController]
    public class BaseController : ControllerBase
    {
    }
}
