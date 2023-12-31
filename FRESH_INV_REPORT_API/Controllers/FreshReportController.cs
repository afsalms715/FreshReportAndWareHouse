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
        public ActionResult<IEnumerable<InventoInvDtl>> GetInventoInvDtl(string ToDate,string LocCode)
        {
            var inventoInvDtls= Services.GetInventoInvDtl(ToDate, LocCode);
            return inventoInvDtls;
        }

        [HttpGet("FreshReportGeneration")]
        public IActionResult FreshReportGeneration(string Location,string OpenInvId,string CloseInvId,string InventoInvId,string Sections,string FromDate,string ToDate)
        {
            //Services.ExcelGeneration();
            string filePath=Services.GenerateFreshReport(Location, OpenInvId, CloseInvId, InventoInvId, Sections, FromDate, ToDate);
            if (System.IO.File.Exists(filePath))
            {
                // Return the file as a response
                var fileContent = System.IO.File.ReadAllBytes(filePath);
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                var fileName = "FRESH_GP_REPORT_"+ Location + "_"+ ToDate + ".xlsx"; // Provide a meaningful file name

                // Optionally, you can delete the file after sending it as a response
                // System.IO.File.Delete(filePath);

                return File(fileContent, contentType, fileName);
            }
            else
            {
                // Handle the case where the file does not exist
                return NotFound("File not found");
            }
        }

        [HttpGet("dateFormate")]
        public string GetFormatedDate(string date)
        {
            DateTime FormatedDate = DateTime.Parse(date);
            return FormatedDate.ToString("dd-MMM-yy").ToUpper();
        }
    }
}
