using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
namespace New_Web.Models
{
    public class Inventory
    {
        public string ItemCode { get; set; }
        public int FirstStock { get; set; }
        public int Sales { get; set; }
        public int Refund { get; set; }
        public int GoodsReturn { get; set; }
        public int GoodsReceiptPO { get; set; }
        public int GoodsReceipt { get; set; }
        public int GoodsIssue { get; set; }
        public int TransferOut { get; set; }
        public int TransferIn { get; set; }
        public int LastStock { get; set; }

        [Required]
        [Display(Name = "Date Of Birth")]
        [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MMM/yyyy}")]
        public DateTime DOB { get; set; }

    }

    


}