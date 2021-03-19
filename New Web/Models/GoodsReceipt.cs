using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
namespace New_Web.Models
{
    public class GoodsReceipt
    {
        public string BranchID { get; set; }
        public string GRCode { get; set; }
        public string Status { get; set; }
        public DateTime DocDate { get; set; }
        public string VendorCode { get; set; }
        public string VendorName { get; set; }
        public string Remarks { get; set; }

        [Required]
        [Display(Name = "Date Of Birth")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MMM/yyyy}")]
        public DateTime DOB { get; set; }

    }
}