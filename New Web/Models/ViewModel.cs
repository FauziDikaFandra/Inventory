using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace New_Web.Models
{
    public class ViewModel
    {
        public List<Inventory> Inventory { get; set; }
        public List<Vendor> Vendor { get; set; }
        public VendorBack VendorBack { get; set; }
    }
}