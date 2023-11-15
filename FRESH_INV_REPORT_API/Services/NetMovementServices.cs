using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FRESH_INV_REPORT_API.Services
{
    public class NetMovementServices
    {
        public string connectionString;
        public NetMovementServices(string connectionString)
        {
            this.connectionString = connectionString;
        }
    }
}
