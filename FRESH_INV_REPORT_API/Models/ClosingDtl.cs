using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FRESH_INV_REPORT_API.Models
{
    public class ClosingDtl
    {
        public int Id { get; set; }
        public int ClosingInvId { get; set; }
        public string Category { get; set; }
    }
}
