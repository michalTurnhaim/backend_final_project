using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using DTO;
using BLL;



namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OwnerOfEventController : ControllerBase
    {
        IOwnerOfEventBll _ownerOfEventBll;
        public OwnerOfEventController(IOwnerOfEventBll ownerOfEventBll)
        {
            _ownerOfEventBll = ownerOfEventBll;
        }

        [HttpGet("getAllOwnerOfEvent")]
            public IActionResult getAllOwnerOfEvent()
            {
                return Ok(_ownerOfEventBll.getAllOwnerOfEventd());
            }

            [HttpDelete("deleteTheOwnerOfEvent/{id}")]
            public IActionResult deleteOwnerOfEventd(int id)
            {
                return Ok(_ownerOfEventBll.deleteOwnerOfEvent(id));
            }

            [HttpPut("updateTheOwnerOfEvent/{id}")]
            public IActionResult updateTheOwnerOfEvent([FromBody] OwnerOfEventDto i, int id)
            {
                return Ok(_ownerOfEventBll.updateOwnerOfEventd(i, id));
            }

            [HttpPost("addTheOwnerOfEvent")]
            public IActionResult addTheOwnerOfEvent([FromBody] OwnerOfEventDto i)
            {
                return Ok(_ownerOfEventBll.addOwnerOfEvent(i));

            }
        //[HttpGet("getOwnerOfEventByEmail/{email}")]
        //public IActionResult getOwnerOfEventByEmail(string email)
        //{
        //    return Ok(_ownerOfEventBll.getOwnerOfEventByEmail(email));
        //}
        }
    }

