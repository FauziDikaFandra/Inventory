using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace New_Web.Models
{
    public class GoodsReceiptDetails
    {
        public string BranchID { get; set; }
        public string GRCode { get; set; }
        public string Status { get; set; }
        public DateTime DocDate { get; set; }
        public string VendorCode { get; set; }
        public string VendorName { get; set; }
        public string Remarks { get; set; }
        public string ItemCode { get; set; }
        public string Description { get; set; }
        public string CurrentPrice { get; set; }
        public int Quantity { get; set; }
    }
}