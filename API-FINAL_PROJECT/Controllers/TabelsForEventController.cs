using BLL;
using DTO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TabelsForEventController : ControllerBase
    {
        ITabelsForEventBLL _tabelsForEventBLL;
        public TabelsForEventController(ITabelsForEventBLL tabelsForEventBLL)
        {
            _tabelsForEventBLL = tabelsForEventBLL;
        }

        [HttpGet("getAllOwnerOfEvent")]
        public IActionResult getAllOwnerOfEvent()
        {
            return Ok(_tabelsForEventBLL.getAllTabelsForEvent());
        }

        [HttpDelete("deleteTheTabelsForEvent/{id}")]
        public IActionResult deleteTabelsForEvent(int id)
        {
            return Ok(_tabelsForEventBLL.deleteTabelsForEvent(id));
        }

        [HttpPut("updateTheTabelsForEvent/{id}")]
        public IActionResult updateTheTabelsForEvent([FromBody] TabelsForEventDto i, int id)
        {
            return Ok(_tabelsForEventBLL.updateTabelsForEvent(i, id));
        }

        [HttpPost("addTheTabelsForEvent")]
        public IActionResult addTheTabelsForEvent([FromBody] TabelsForEventDto i)
        {
            return Ok(_tabelsForEventBLL.addTabelsForEvent(i));

        }
    }
}
