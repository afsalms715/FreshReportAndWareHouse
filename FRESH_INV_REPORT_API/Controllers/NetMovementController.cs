using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FRESH_INV_REPORT_API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class NetMovementController : ControllerBase
    {
        [HttpGet("getNetMovement")]
        public void GetNetMovement()
        {

        }
    }
}
