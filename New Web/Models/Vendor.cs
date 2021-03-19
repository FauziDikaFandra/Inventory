using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
namespace New_Web.Models
{
    public class Vendor
    {
        public string VendorCode { get; set; }
        public string VendorName { get; set; }

    }

    public class VendorBack
    {
        public string VendorCode { get; set; }

        [Required]
        [Display(Name = "Date Of Birth")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MMM/yyyy}")]
        public DateTime DOB { get; set; }

    }

    public class StatusBack
    {
        public string Status { get; set; }
        public string Remarks { get; set; }
        public string GRCode { get; set; }

    }
}