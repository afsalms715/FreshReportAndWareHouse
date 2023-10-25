using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WareHouseProdcutApi.Models;
using WareHouseProdcutApi.Services;

namespace WareHouseProdcutApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [EnableCors("MyCorePolicy")]
    public class WareHouseController : ControllerBase
    {
        public readonly WhPrdExpDateService services;
        public WareHouseController(WhPrdExpDateService service)
        {
            this.services = service;
        }

        [HttpGet("Products")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public ActionResult<IEnumerable<WhProduct>> GetProducts(long product)
        {
            if (product == 0)
            {
                ModelState.AddModelError("Product", "Invalid value you Entered Please Check the value your input");//costom error message
                return BadRequest(ModelState);
            }
            var result = services.GetWhProducts(product);
            if (result.Count==0)
            {
                return NotFound();
            }
            return Ok(result);
        }

        [HttpPatch("UpdateExpDate")]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public IActionResult UpdateExpDate(long barcode, string updateDate)
        {
            DateTime Date;
            string msg;
            if (barcode == 0 || updateDate=="")
            {
                return BadRequest();
            }
            if(DateTime.TryParse(updateDate, out Date)){
                msg=services.UpdateExpDate(barcode, Date);
            }
            else
            {
                ModelState.AddModelError("Date", "Invalid date ");
                return BadRequest(ModelState);
            }
            if (msg=="OK")
            {
                return NoContent();
            }
            return StatusCode(500,msg);
        }

        [HttpGet("login")]

        public ActionResult<Users> UserLogin(string User_id,string Password)
        {
            if (User_id == ""||Password=="")
            {
                return BadRequest();
            }
            Users user=services.UserLogin(User_id, Password);
            if (user.USER_ID!=User_id)
            {
                return NotFound();
            }
            return user;
        }
    }
}
