using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using BLL;
using DTO;

namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class InvitedController : ControllerBase
    {
        IInvitedBLL invitedBLL;

        public InvitedController(IInvitedBLL invitedBLL)
        {
            this.invitedBLL = invitedBLL;
        }

        [HttpGet("getAllInvited")]
       public IActionResult getAllInvited()
        {
            return Ok(invitedBLL.getAllInvited()) ;
        }

        [HttpDelete("deleteTheInvited/{id}")]
        public IActionResult deleteTheInvited(string id)
        {
            return Ok(invitedBLL.deleteInvited(id));
        }

        [HttpPut("updateTheInvited/{id}")]
        public IActionResult updateTheInvited([FromBody] InvitedDto i, string id )
        {
            return Ok(invitedBLL.updateInvited(i,id));
        }
        
        [HttpPost("addTheInvited")]
        public IActionResult addTheInvited([FromBody] InvitedDto i)
        {
            return Ok(invitedBLL.addInvited(i));

        }
        [HttpGet("login/{email}/{password}")]
        public IActionResult login(string email,int password)
        {
            return Ok(invitedBLL.login(email, password));

        }
        [HttpGet("checEmailIfExists/{email}")]
        public IActionResult checEmailIfExists(string email)
        {
            return Ok(invitedBLL.checEmailIfExists(email));

        }

    }
}
