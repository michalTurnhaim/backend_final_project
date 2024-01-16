using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using BLL;
using DTO;
namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class InvitedToEventController : ControllerBase
    {
        IInvitedToEventBLL _invitedToEventBLL;
        public InvitedToEventController(IInvitedToEventBLL invitedToEventBLL)
        {
            _invitedToEventBLL = invitedToEventBLL;
        }

        [HttpGet("getAllInvitedToEvent")]
        public IActionResult getAllInvitedToEvent()
        {
            return Ok(_invitedToEventBLL.getAllInvitedToEventd());
        }

        [HttpDelete("deleteTheInvitedToEvent/{id}")]
        public IActionResult deleteInvitedToEventd(int id)
        {
            return Ok(_invitedToEventBLL.deleteInvitedToEventd(id));
        }

        [HttpPut("updateTheInvitedToEvent/{id}")]
        public IActionResult updateInvitedToEventd([FromBody] InvitedToEventDto i, int id)
        {
            return Ok(_invitedToEventBLL.updateInvitedToEventd(i, id));
        }

        [HttpPost("addTheInvitedToEvent")]
        public IActionResult addTheInvitedToEvent([FromBody] InvitedToEventDto i)
        {
            return Ok(_invitedToEventBLL.addInvitedToEvent(i));

        }
        [HttpGet("InvitedToEventbyEmail/{email}")]
        public IActionResult InvitedToEventbyEmail(string email)
        {
            return Ok(_invitedToEventBLL.InvitedToEventbyEmail(email));
        }
    }
}
