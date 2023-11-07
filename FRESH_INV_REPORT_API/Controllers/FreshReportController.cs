using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using FRESH_INV_REPORT_API.Services;

namespace FRESH_INV_REPORT_API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FreshReportController : ControllerBase
    {
        [HttpGet("openingDtl")]
        public void GetOpeningDtl()
        {

        }
    }
}
