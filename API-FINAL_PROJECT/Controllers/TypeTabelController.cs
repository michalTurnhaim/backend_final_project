using BLL;
using DTO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TypeTabelController : ControllerBase
    {
        ITypeTabelBLL _typeTabelBLL;
        public TypeTabelController(ITypeTabelBLL typeTabelBLL)
        {
            _typeTabelBLL = typeTabelBLL;
        }

        [HttpGet("getAllTypeTabel")]
        public IActionResult getAllTypeTabel()
        {
            return Ok(_typeTabelBLL.getAllTypeTabel());
        }

        [HttpDelete("deleteTheTypeTabel/{id}")]
        public IActionResult deleteTypeTabel(int id)
        {
            return Ok(_typeTabelBLL.deleteTypeTabel(id));
        }

        [HttpPut("updateTheTypeTabel/{id}")]
        public IActionResult updateTheTypeTabel([FromBody] TypeTabelsDto i, int id)
        {
            return Ok(_typeTabelBLL.updateTypeTabel(i, id));
        }

        [HttpPost("addTheTypeTabel")]
        public IActionResult addTheTypeTabel([FromBody] TypeTabelsDto i)
        {
            return Ok(_typeTabelBLL.addTypeTabel(i));

        }
    }
}
