using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Globalization;
using FRESH_INV_REPORT_API.Services;
using FRESH_INV_REPORT_API.Models;

namespace FRESH_INV_REPORT_API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FreshReportController : ControllerBase
    {
        FreshReportServices Services;
        public FreshReportController(FreshReportServices services)
        {
            Services = services;
        }
        
        [HttpGet("openingId")]
        public ActionResult<IEnumerable<OpeningInvDtl>> GetOpeningDtl(string FromDate,string LocCode)
        {
            if (FromDate =="" || LocCode=="")
            {
                return BadRequest();
            }
            var openingInvDtl = Services.GetOpeningInvId(FromDate, LocCode);
            return openingInvDtl;
        }

        [HttpGet("closingDtl")]
        public ActionResult<IEnumerable<ClosingDtl>> GetClosingDtl(string ToDate,string LocCode)
        {
            var closingDtl = Services.GetClosingInvDtl(ToDate, LocCode);
            return closingDtl;
        }

        [HttpGet("InventoInvDtl")]
        public ActionResult<InventoInvDtl> GetInventoInvDtl(string ToDate,string LocCode)
        {
            InventoInvDtl inventoInvDtl= Services.GetInventoInvDtl(ToDate, LocCode);
            return inventoInvDtl;
        }

        [HttpGet("dateFormate")]
        public string GetFormatedDate(string date)
        {
            DateTime FormatedDate = DateTime.Parse(date);
            return FormatedDate.ToString("dd-MMM-yy").ToUpper();
        }
    }
}
