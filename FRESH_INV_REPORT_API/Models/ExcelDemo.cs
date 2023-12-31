using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FRESH_INV_REPORT_API.Models
{
    public class ExcelDemo
    {
        public int Id { get; set; }
        public int PRODUCT_CODE{ get; set; }
        public string STOCK_UNIT { get; set; }
        public int SU { get; set; }
        public long BARCODE { get; set; }
        public string SECTION { get; set; }
        public string SU_DESC { get; set; }
        public float UNIT_COST { get; set; }
        public float OPENING_QTY { get; set; }
        public float OPENING_VAL { get; set; }
        public float NET_PURCHASE_QTY { get; set; }
        public float NET_PURCHASE_VAL { get; set; }
        public float SALES_QTY { get; set; }
        public float SALES_VAL { get; set; }
        public float WASTAGE_QTY { get; set; }
        public float WASTAGE_VAL { get; set; }
        public float SYS_QTY { get; set; }
        public float SYS_VAL { get; set; }
        public float CLOSING_QTY { get; set; }
        public float CLOSING_VAL { get; set; }
        public float UNKNWON_WASTAGE_QTY { get; set; }
        public float UNKNOWN_WASTAGE_VAL { get; set; }
        public float COST_OF_SALE { get; set; }
        public float PROFIT { get; set; }
        public float GP { get; set; }

    }
}
