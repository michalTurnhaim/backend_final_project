using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using BLL;
using DTO;

namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PutInvitedOnTabelController : ControllerBase
    {
        IPutInvitedOnTabelBLL _putInvitedOnTabelBLL;
        public PutInvitedOnTabelController(IPutInvitedOnTabelBLL putInvitedOnTabelBLL)
        {
            _putInvitedOnTabelBLL = putInvitedOnTabelBLL;
        }

        [HttpGet("getAllPutInvitedOnTabel")]
        public IActionResult getAllPutInvitedOnTabel()
        {
            return Ok(_putInvitedOnTabelBLL.getAllPutInvitedOnTabel());
        }

        [HttpDelete("deletePutInvitedOnTabel/{id}")]
        public IActionResult deletePutInvitedOnTabel(int id)
        {
            return Ok(_putInvitedOnTabelBLL.deletePutInvitedOnTabel(id));
        }

        [HttpPut("updatePutInvitedOnTabel/{id}")]
        public IActionResult updatePutInvitedOnTabel([FromBody] PutInvitedOnTabelDto i, int id)
        {
            return Ok(_putInvitedOnTabelBLL.updatePutInvitedOnTabel(i, id));
        }

        [HttpPost("addThePutInvitedOnTabel")]
        public IActionResult addThePutInvitedOnTabel([FromBody] PutInvitedOnTabelDto i)
        {
            return Ok(_putInvitedOnTabelBLL.addPutInvitedOnTabel(i));

        }
    }
}
