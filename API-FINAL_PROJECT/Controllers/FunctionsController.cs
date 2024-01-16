using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using BLL;
using DTO;

namespace API_FINAL_PROJECT.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FunctionsController : ControllerBase
    {
        Ifunctions Ifunctions;
        public FunctionsController(Ifunctions ifunctions)
        {
            Ifunctions = ifunctions;
        }

        [HttpPost("up")]
        public async Task<IActionResult> UpFile([FromForm] IFormFile file)
        {

            if (file == null || file.Length == 0)
                return BadRequest("No file was uploaded.");
            var fileName = Guid.NewGuid().ToString() + Path.GetExtension(file.FileName);
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", fileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
            if (fileName.EndsWith(".xlsx"))
                Ifunctions.READExcel(fileName);
            else
                Ifunctions.putNameFileInvitationInData(fileName);

            return Ok("ghj");
        }
        [HttpPost("postowner")]
        public IActionResult putqwner([FromBody] OwnerOfEventNew o)
        {
            return Ok(Ifunctions.addOwnerOfEvent(o));
        }
        [HttpGet("returnNameFile/{id}")]
        public IActionResult returnNameFile(int id)
        { return Ok(Ifunctions.returnNameFile(id)); }
        //פונקציה שמחזירה את כל הארועים שהזמין בעל ארוע אחד לפי מייל
        [HttpGet("returnListOfOwnerByEmail/{email}")]
        public IActionResult returnListOfOwnerByEmail(string email)
        {
            return Ok(Ifunctions.returnListOfOwnerByEmail(email));
        }
        //פונקציה שמחזירה את כל המוזמנים לארוע מסוים
        [HttpGet("invitedToEventDtoList/{id}")]
        public IActionResult invitedToEventDtoList(int id)
        {
            return Ok(Ifunctions.invitedToEventDtoList(id));
        }
        
        [HttpPut("BeforSendingEmail/{idevent}")]
        public IActionResult BeforSendingEmail(int idevent)
        {
            return Ok(Ifunctions.BeforSendingEmail(idevent));
        }
        [HttpGet("SendEmail/{email}/{password}/{fileOrder}")]
        public IActionResult SendEmail(string email,int password, string fileOrder)
        {
            return Ok(Ifunctions.SendEmail(email, password, fileOrder));
        }
    }
}
