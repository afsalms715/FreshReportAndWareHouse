using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WareHouseProdcutApi.Models
{
    public class WhProduct
    {
        public int Id { get; set; }
        public long Barcode { get; set; }
        public int ArtCode { get; set; }
        public int Su { get; set; }
        public string SuDescription { get; set; }
        public DateTime ExpDate { get; set; }
    }
}
