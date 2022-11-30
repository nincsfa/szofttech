using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using T0KW0U_10HET.Models;

namespace T0KW0U_11HET.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]

    public class BoatController : Controller
    {

        [HttpGet]
        [Route("questions/all")]
        public IActionResult MindegyHogyHivjak()
        {

            HajosContext context = new HajosContext();
            var kérdések = from x in context.Questions
                           select x.Question1;

            return Ok(kérdések);
        }
    }
}
