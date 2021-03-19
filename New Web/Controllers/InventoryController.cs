using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.SqlClient;
using New_Web.Models;
using System.Dynamic;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace New_Web.Controllers
{
    
    public class InventoryController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;
        void connectionString()
        {
            con.ConnectionString = "Data Source=DESKTOP-IL11GFI;Initial Catalog=New_Web;User ID=sa;Password=star;";
        }

        // GET: Inventory
        public ActionResult Index()
        {
            ViewModel mymodel = new ViewModel();
            mymodel.Inventory = GetInventory();
            mymodel.Vendor = GetVendor();
            mymodel.VendorBack = GetVendorBack("**A", DateTime.Now.Month, DateTime.Now.Year);
            return View(mymodel);

        }

        private List<Inventory> GetInventory()
        {
            List<Inventory> inv = new List<Inventory>();
            //var inv = new List<Inventory>();
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "select * from s_inventory where year(Docdate) = year(getdate()) and month(Docdate) = month(getdate())";
            dr = com.ExecuteReader();

            while (dr.Read())
            {
                var inven = new Inventory();
                inven.ItemCode = Convert.ToString(dr["ItemCode"]);
                inven.FirstStock = Convert.ToInt32(dr["FirstStock"]);
                inven.Sales = Convert.ToInt32(dr["Sales"]);
                inven.Refund = Convert.ToInt32(dr["Refund"]);
                inven.GoodsReturn = Convert.ToInt32(dr["GoodsReturn"]);
                inven.GoodsReceiptPO = Convert.ToInt32(dr["GoodsReceiptPO"]);
                inven.GoodsReceipt = Convert.ToInt32(dr["GoodsReceipt"]);
                inven.GoodsIssue = Convert.ToInt32(dr["GoodsIssue"]);
                inven.TransferOut = Convert.ToInt32(dr["TransferOut"]);
                inven.TransferIn = Convert.ToInt32(dr["TransferIn"]);
                inven.LastStock = Convert.ToInt32(dr["LastStock"]);
                inv.Add(inven);
            }
            con.Close();
            return inv;
        }


        private List<Inventory> GetInventory(int bulan, int tahun, string vendor)
        {
            List<Inventory> inv = new List<Inventory>();
            //var inv = new List<Inventory>();
            connectionString();
            con.Open();
            com.Connection = con;
            if (vendor == "**A")
            {
                com.CommandText = "select a.* from s_inventory a inner join s_item_master b on a.ItemCode = b.ItemCode where year(Docdate) = " + tahun + " and month(Docdate) = " + bulan + "";
            }
            else
            {
                com.CommandText = "select a.* from s_inventory a inner join s_item_master b on a.ItemCode = b.ItemCode where year(Docdate) = " + tahun + " and month(Docdate) = " + bulan + " and VendorCode = '" + vendor + "'";
            }
            
            dr = com.ExecuteReader();

            while (dr.Read())
            {
                var inven = new Inventory();
                inven.ItemCode = Convert.ToString(dr["ItemCode"]);
                inven.FirstStock = Convert.ToInt32(dr["FirstStock"]);
                inven.Sales = Convert.ToInt32(dr["Sales"]);
                inven.Refund = Convert.ToInt32(dr["Refund"]);
                inven.GoodsReturn = Convert.ToInt32(dr["GoodsReturn"]);
                inven.GoodsReceiptPO = Convert.ToInt32(dr["GoodsReceiptPO"]);
                inven.GoodsReceipt = Convert.ToInt32(dr["GoodsReceipt"]);
                inven.GoodsIssue = Convert.ToInt32(dr["GoodsIssue"]);
                inven.TransferOut = Convert.ToInt32(dr["TransferOut"]);
                inven.TransferIn = Convert.ToInt32(dr["TransferIn"]);
                inven.LastStock = Convert.ToInt32(dr["LastStock"]);
                inv.Add(inven);
            }
            con.Close();
            return inv;
        }

        private List<Vendor> GetVendor()
        {
            List<Vendor> Vendor = new List<Vendor>();
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "select * from s_vendor order by VendorName";
            dr = com.ExecuteReader();
            while (dr.Read())
            {
                var inven = new Vendor();
                inven.VendorCode = Convert.ToString(dr["VendorCode"]);
                inven.VendorName = Convert.ToString(dr["VendorName"]);
                Vendor.Add(inven);
            }
            con.Close();
            return Vendor;
        }

        private VendorBack GetVendorBack(string code, int bulan, int tahun)
        {
            VendorBack VendorBack = new VendorBack();
            VendorBack.VendorCode = code;
            string iDate = bulan + "/01/" + tahun;
            DateTime oDate = Convert.ToDateTime(iDate);
            VendorBack.DOB = oDate;
            return VendorBack;
        }

        // GET: Inventory/Details/5
        public ActionResult Details(int bulan, int tahun, string vendor)
        {
            ViewModel mymodel = new ViewModel();
            mymodel.Inventory = GetInventory(bulan, tahun, vendor.Substring(0, 3));
            mymodel.Vendor = GetVendor();
            mymodel.VendorBack = GetVendorBack(vendor.Substring(0, 3),bulan,tahun);
            return View("Index",mymodel);
        }

        public ActionResult DownloadPDF(int bulan, int tahun, string vendor)
        {
            try
            {
                //var model = new GeneratePDFModel();
                ViewModel mymodel = new ViewModel();
                mymodel.Inventory = GetInventory(bulan, tahun, vendor.Substring(0, 3));
                mymodel.Vendor = GetVendor();
                mymodel.VendorBack = GetVendorBack(vendor.Substring(0, 3), bulan, tahun);
                //get the information to display in pdf from database
                //for the time
                //Hard coding values are here, these are the content to display in pdf 
                //var content = "<h2>WOW Rotativa<h2>" +
                // "<p>Ohh This is very easy to generate pdf using Rotativa <p>";
                //var logoFile = @"/Content/img/logo.png";

                //model.PDFContent = content;
                //model.PDFLogo = Server.MapPath(logoFile);

                //Use ViewAsPdf Class to generate pdf using GeneratePDF.cshtml view
                //return new Rotativa.MVC.ViewAsPdf("GeneratePDF", model) { FileName = "Inventory.pdf" };
                return new Rotativa.MVC.ViewAsPdf("GeneratePDF", mymodel) { FileName = "Inventory.pdf" };
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public ActionResult DownloadExcel(int bulan, int tahun, string vendor)
        {
           
            Excel.Application application = new Excel.Application();
            Excel.Workbook workbook = application.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet worksheet = workbook.ActiveSheet;


            ViewModel mymodel = new ViewModel();
            mymodel.Inventory = GetInventory(bulan, tahun, vendor.Substring(0, 3));
            mymodel.Vendor = GetVendor();
            mymodel.VendorBack = GetVendorBack(vendor.Substring(0, 3), bulan, tahun);


            worksheet.Cells[1, 1].Value = "ItemCode";
            worksheet.Cells[1, 2].Value = "FirstStock";
            worksheet.Cells[1, 3].Value = "Sales";
            worksheet.Cells[1, 4].Value = "Refund";
            worksheet.Cells[1, 5].Value = "Return";
            worksheet.Cells[1, 6].Value = "Receipt";
            worksheet.Cells[1, 7].Value = "Adjust +";
            worksheet.Cells[1, 8].Value = "Adjust -";
            worksheet.Cells[1, 9].Value = "Transfer Out";
            worksheet.Cells[1, 10].Value = "Transfer In";
            worksheet.Cells[1, 11].Value = "Last Stock";
            int row = 2;
            foreach (var item in mymodel.Inventory)
            {

                worksheet.Cells[row, 1] = item.ItemCode;
                worksheet.Cells[row, 2] = item.FirstStock;
                worksheet.Cells[row, 3] = item.Sales;
                worksheet.Cells[row, 4] = item.Refund;
                worksheet.Cells[row, 5] = item.GoodsReturn;
                worksheet.Cells[row, 6] = item.GoodsReceiptPO;
                worksheet.Cells[row, 7] = item.GoodsIssue;
                worksheet.Cells[row, 8] = item.GoodsReceipt;
                worksheet.Cells[row, 9] = item.TransferOut;
                worksheet.Cells[row, 10] = item.TransferIn;
                worksheet.Cells[row, 11] = item.LastStock;
                row++;
            }


            workbook.SaveAs(@"D:\Excel\sample.xlsx");
            workbook.Close();
            Marshal.ReleaseComObject(workbook);

            application.Quit();
            Marshal.FinalReleaseComObject(application);

            //return RedirectToAction("Index");
            return File(@"D:\Excel\sample.xlsx", "Application/Excel", "Inventory.xlsx");
        }


        // GET: Inventory/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Inventory/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Inventory/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: Inventory/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: Inventory/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: Inventory/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
