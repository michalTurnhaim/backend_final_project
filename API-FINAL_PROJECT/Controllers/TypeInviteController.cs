using BLL;
using DTO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TypeInviteController : ControllerBase
    {
        ITypeInviteBLL _TypeInviteBLL;
        public TypeInviteController(ITypeInviteBLL typeInviteBLL)
        {
            _TypeInviteBLL = typeInviteBLL;
        }

        [HttpGet("getAllTypeInvite")]
        public IActionResult getAllTypeInvite()
        {
            return Ok(_TypeInviteBLL.getAllTypeInviteBll());
        }

        [HttpDelete("deleteTheTypeInvite/{id}")]
        public IActionResult deleteTypeInvite(int id)
        {
            return Ok(_TypeInviteBLL.deleteTypeInvite(id));
        }

        [HttpPut("updateTheTypeInvite/{id}")]
        public IActionResult updateTheTypeInvite([FromBody] TypeInviteDto i, int id)
        {
            return Ok(_TypeInviteBLL.updateTypeInvite(i, id));
        }

        [HttpPost("addTheTypeInvite")]
        public IActionResult addTheTypeInvite([FromBody] TypeInviteDto i)
        {
            return Ok(_TypeInviteBLL.addTypeInvite(i));

        }
    }
}
