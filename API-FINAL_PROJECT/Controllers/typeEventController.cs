using BLL;
using DTO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class typeEventController : ControllerBase
    {
        ITypeEventBLL _typeEventBLL;
        public typeEventController(ITypeEventBLL typeEventBLL)
        {
            _typeEventBLL = typeEventBLL;
        }

        [HttpGet("getAllTypeEvent")]
        public IActionResult getAllTypeEvent()
        {
            return Ok(_typeEventBLL.getAllTypeEvent());
        }

        [HttpDelete("deleteTheTypeEvent/{id}")]
        public IActionResult deleteTypeEvent(int id)
        {
            return Ok(_typeEventBLL.deleteTypeEvent(id));
        }

        [HttpPut("updateTheTypeEvent/{id}")]
        public IActionResult updateTypeEvent([FromBody] typeEventDto i, int id)
        {
            return Ok(_typeEventBLL.updateTypeEvent(i, id));
        }

        [HttpPost("addTheTypeEvent")]
        public IActionResult addTheTypeEvent([FromBody] typeEventDto i)
        {
            return Ok(_typeEventBLL.addTypeEvent(i));

        }
    }
}
