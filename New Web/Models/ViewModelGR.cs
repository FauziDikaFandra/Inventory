using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace New_Web.Models
{
    public class ViewModelGR
    {
        public List<GoodsReceipt> GoodsReceipt { get; set; }
        public List<Vendor> Vendor { get; set; }
        public List<ItemModel> Item { get; set; }
        public List<Details> Detail { get; set; }

        public List<GoodsReceiptDetails> GoodsReceiptDetails { get; set; }
        public VendorBack VendorBack { get; set; }
        public StatusBack Status { get; set; }
    }
}