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
    public class WareHouseController : ControllerBase
    {
        public readonly WhPrdExpDateService services;
        public WareHouseController(WhPrdExpDateService service)
        {
            this.services = service;
        }

        [Route("Products")]
        [HttpGet]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public ActionResult<IEnumerable<WhProduct>> GetProducts(long product)
        {
            if (product == 0)
            {
                return BadRequest();
            }
            var result = services.GetWhProducts(product);
            if (result.Count==0)
            {
                return NotFound();
            }
            return Ok(result);
        }
    }
}
