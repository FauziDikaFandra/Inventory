using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using New_Web.Models;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace New_Web.Controllers
{
    public class GoodsReceiptController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;
        SqlDataAdapter da = new SqlDataAdapter();
        DataSet ds = new DataSet();
        void connectionString()
        {
            con.ConnectionString = "Data Source=DESKTOP-IL11GFI;Initial Catalog=New_Web;User ID=sa;Password=star;";
        }

        public DataSet GetSQLDB(string cmd)
        {
            connectionString();
            com = con.CreateCommand();
            com.CommandText = cmd;
            da.SelectCommand = com;
            con.Open();
            da.Fill(ds);
            con.Close();
            return ds;
        }
        // GET: GoodsReceipt
        public ActionResult Index()
        {
            ViewModelGR mymodel = new ViewModelGR();
            mymodel.GoodsReceipt = GetGoodsReceipt("**A", @DateTime.Now, "ALL STATUS");
            mymodel.Vendor = GetVendor();
            mymodel.VendorBack = GetVendorBack("**A", @DateTime.Now);
            mymodel.Status = GetStatusBack("ALL STATUS","","");
            return View(mymodel);
        }

        // GET: GoodsReceipt/Details/5
        public ActionResult Details(string vendor, int tgl, int bulan, int tahun, string status)
        {
            string iDate = bulan + "/"+ tgl + "/" + tahun;
            DateTime oDate = Convert.ToDateTime(iDate);
            ViewModelGR mymodel = new ViewModelGR();
            mymodel.GoodsReceipt = GetGoodsReceipt(vendor.Substring(0, 3), oDate, status);
            mymodel.Vendor = GetVendor();
            mymodel.VendorBack = GetVendorBack(vendor.Substring(0, 3), oDate);
            mymodel.Status = GetStatusBack(status,"","");
            return View("Index", mymodel);
        }

        // GET: GoodsReceipt/Create
        public ActionResult Create()
        {
            ViewModelGR mymodel = new ViewModelGR();
            mymodel.GoodsReceipt = GetGoodsReceipt("**A", @DateTime.Now, "ALL STATUS");
            mymodel.Vendor = GetVendor();
            mymodel.VendorBack = GetVendorBack("**A", @DateTime.Now);
            mymodel.Status = GetStatusBack("ALL STATUS","","");
            return View(mymodel);
        }

        // POST: GoodsReceipt/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            ViewModelGR mymodel = new ViewModelGR();
            try
            {
                
                mymodel.GoodsReceipt = GetGoodsReceipt(collection["vendor"].Substring(0, 3), @DateTime.Now, "OPEN");
                mymodel.Vendor = GetVendor();
                mymodel.VendorBack = GetVendorBack(collection["vendor"].Substring(0, 3), @DateTime.Now);
                mymodel.Status = GetStatusBack("OPEN","","");


                string GRCode; 
                DataSet dx = new DataSet();
                dx = GetSQLDB("select Top 1 Convert(Int,Substring(GRCode,3,3)) + 1 GRCode from s_goods_receipt order by Convert(Int,Substring(GRCode,3,3)) Desc");
                if (dx.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToInt32(dx.Tables[0].Rows[0]["GRCode"]) > 99)
                    {
                        GRCode = "GR" + dx.Tables[0].Rows[0]["GRCode"].ToString();
                    }
                    else if (Convert.ToInt32(dx.Tables[0].Rows[0]["GRCode"]) > 9)
                    {
                        GRCode = "GR0" + dx.Tables[0].Rows[0]["GRCode"].ToString();
                    }
                    else
                    {
                        GRCode = "GR00" + dx.Tables[0].Rows[0]["GRCode"].ToString();
                    }
                    
                }
                else
                {
                    GRCode = "GR001";
                }

                string iDate = collection["bulan"] + "/"+ collection["tgl"] + "/ " + collection["tahun"];
                DateTime oDate = Convert.ToDateTime(iDate);


                connectionString();
                con.Open();
                com.Connection = con;                
                com.CommandText = "insert into s_goods_receipt values ('B1','"+ GRCode +"', '" + collection["status"] + "', " +
                    "'" + oDate + "', '"+ collection["vendor"].Substring(0,3) + "' , '" + collection["vendor"].Substring(5, collection["vendor"].Length - 5) + "', " +
                    "'" + collection["remarks"] + "')";
                int i = com.ExecuteNonQuery();
                con.Close();
                if (i >= 1)
                {

                    return RedirectToAction("Index", mymodel);

                }
                else
                {

                    return View(mymodel);
                }          
               
            }
            catch
            {
                return View(mymodel);
            }
        }

        // GET: GoodsReceipt/Edit/5
        [HttpGet]
        public ActionResult Edit(string GRCode)
        {
            DataSet dx = new DataSet();
            ViewModelGR mymodel = new ViewModelGR();
            dx = GetSQLDB("select * from s_goods_receipt where GRCode = '"+ GRCode + "' ");
            if (dx.Tables[0].Rows.Count > 0)
            {               
                mymodel.GoodsReceipt = GetGoodsReceipt(dx.Tables[0].Rows[0]["VendorCode"].ToString(), Convert.ToDateTime(dx.Tables[0].Rows[0]["DocDate"]), "OPEN");
                mymodel.Vendor = GetVendor();
                mymodel.VendorBack = GetVendorBack(dx.Tables[0].Rows[0]["VendorCode"].ToString(), Convert.ToDateTime(dx.Tables[0].Rows[0]["DocDate"]));
                mymodel.Status = GetStatusBack(dx.Tables[0].Rows[0]["Status"].ToString(), dx.Tables[0].Rows[0]["Remarks"].ToString(), GRCode);
            }
            return View(mymodel);
        }

        // POST: GoodsReceipt/Edit/5
        [HttpPost]
        public ActionResult Edit(FormCollection collection)
        {
            ViewModelGR mymodel = new ViewModelGR();
            try
            {
                mymodel.GoodsReceipt = GetGoodsReceipt(collection["vendor"].Substring(0, 3), @DateTime.Now, "OPEN");
                mymodel.Vendor = GetVendor();
                mymodel.VendorBack = GetVendorBack(collection["vendor"].Substring(0, 3), @DateTime.Now);
                mymodel.Status = GetStatusBack("OPEN", "", "");
                string iDate = collection["bulan"] + "/" + collection["tgl"] + "/ " + collection["tahun"];
                DateTime oDate = Convert.ToDateTime(iDate);



                connectionString();
                con.Open();
                com.Connection = con;
                com.CommandText = "update s_goods_receipt set Status = '" + collection["status"] + "', DocDate = '" + oDate.Date + "', " +
                    "VendorCode =  '" + collection["vendor"].Substring(0, 3) + "' ,VendorName = '" + collection["vendor"].Substring(5, collection["vendor"].Length - 5) + "'," +
                    "Remarks = '" + collection["remarks"] + "' where GRCode = '"+ collection["grcode"] + "'";
                int i = com.ExecuteNonQuery();
                con.Close();
                if (i >= 1)
                {

                    return RedirectToAction("Index", mymodel);

                }
                else
                {

                    return View(mymodel);
                }
            }
            catch
            {
                return View(mymodel);
            }
        }

        [HttpGet]
        public ActionResult Line(string GRCode)
        {
            DataSet dx = new DataSet();
            ViewModelGR mymodel = new ViewModelGR();
            dx = GetSQLDB("select * from s_goods_receipt where GRCode = '" + GRCode + "' ");
            if (dx.Tables[0].Rows.Count > 0)
            {
                mymodel.GoodsReceipt = GetGoodsReceipt(dx.Tables[0].Rows[0]["VendorCode"].ToString(), Convert.ToDateTime(dx.Tables[0].Rows[0]["DocDate"]), "OPEN");
                mymodel.Vendor = GetVendor();
                mymodel.Item = GetItem();
                mymodel.Detail = GetDetails(GRCode);
                mymodel.VendorBack = GetVendorBack(dx.Tables[0].Rows[0]["VendorCode"].ToString(), Convert.ToDateTime(dx.Tables[0].Rows[0]["DocDate"]));
                mymodel.Status = GetStatusBack(dx.Tables[0].Rows[0]["Status"].ToString(), dx.Tables[0].Rows[0]["Remarks"].ToString(), GRCode);
            }
            return View("Details",mymodel);
        }

        [HttpPost]
        public ActionResult Line(FormCollection collection)
        {
            ViewModelGR mymodel = new ViewModelGR();
            try
            {

                
                decimal price = 0;
                DataSet dx = new DataSet();
                dx = GetSQLDB("select CurrentPrice from s_item_master where ItemCode = '" + collection["itemcode"].Substring(0, 8)  + "'");
                if (dx.Tables[0].Rows.Count > 0)
                {
                    price = Convert.ToDecimal(dx.Tables[0].Rows[0]["CurrentPrice"]);
                }


                connectionString();
                con.Open();
                com.Connection = con;
                com.CommandText = "insert into s_goods_receipt_details values ('" + collection["grcode"] + "','" + collection["itemcode"].Substring(0, 8) + "' ," +
                    " '" + collection["itemcode"].Substring(10, collection["itemcode"].Length - 10) + "','"+ price + "','" + collection["quantity"] + "')";
                int i = com.ExecuteNonQuery();
                con.Close();
                mymodel.GoodsReceipt = GetGoodsReceipt("**A", @DateTime.Now, "OPEN");
                mymodel.Vendor = GetVendor();
                mymodel.Item = GetItem();
                mymodel.Detail = GetDetails(collection["grcode"]);
                mymodel.VendorBack = GetVendorBack("**A", @DateTime.Now);
                mymodel.Status = GetStatusBack("OPEN", "", collection["grcode"]);
                if (i >= 1)
                {

                    return View("Details", mymodel);

                }
                else
                {

                    return View(mymodel);
                }

            }
            catch
            {
                return View("Details", mymodel);
            }
        }

        // GET: GoodsReceipt/Delete/5
        public ActionResult DeleteLine(int itemcode, string grcode)
        {
            ViewModelGR mymodel = new ViewModelGR();        
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "delete from s_goods_receipt_details where GRCode = '"+ grcode + "' and ItemCode = '" + itemcode + "'";
            com.ExecuteNonQuery();
            con.Close();
            mymodel.GoodsReceipt = GetGoodsReceipt("**A", @DateTime.Now, "OPEN");
            mymodel.Vendor = GetVendor();
            mymodel.Item = GetItem();
            mymodel.Detail = GetDetails(grcode);
            mymodel.VendorBack = GetVendorBack("**A", @DateTime.Now);
            mymodel.Status = GetStatusBack("OPEN", "", grcode);
            return View("Details", mymodel);
        }

        // POST: GoodsReceipt/Delete/5
        [HttpGet]
        public ActionResult Delete(string grcode, int tgl, int bulan, int tahun, string vendor, string status)
        {
            try
            {
                string iDate = bulan + "/" + tgl + "/" + tahun;
                DateTime oDate = Convert.ToDateTime(iDate);
                ViewModelGR mymodel = new ViewModelGR();
                connectionString();
                con.Open();
                com.Connection = con;
                com.CommandText = "delete from s_goods_receipt where GRCode = '" + grcode + "'";
                com.ExecuteNonQuery();
                con.Close();

                con.Open();
                com.Connection = con;
                com.CommandText = "delete from s_goods_receipt_details where GRCode = '" + grcode + "'";
                com.ExecuteNonQuery();
                con.Close();

                mymodel.GoodsReceipt = GetGoodsReceipt(vendor.Substring(0, 3), oDate, status);
                mymodel.Vendor = GetVendor();
                mymodel.VendorBack = GetVendorBack(vendor.Substring(0, 3), oDate);
                mymodel.Status = GetStatusBack(status, "", "");
                return View("Index", mymodel);
            }
            catch
            {
                return View();
            }
        }

        private List<GoodsReceipt> GetGoodsReceipt(string vendor, DateTime docdate, string status)
        {
            List<GoodsReceipt> gr = new List<GoodsReceipt>();
            connectionString();
            con.Open();
            com.Connection = con;
            if (vendor == "**A" && status == "ALL STATUS")
            {
                com.CommandText = "select * from s_goods_receipt where DocDate = '" + docdate.Date + "'";
            }
            else if (vendor != "**A" && status == "ALL STATUS")
            {
                com.CommandText = "select * from s_goods_receipt where VendorCode = '"+ vendor + "' and DocDate = '"+ docdate.Date + "'";
            }
            else if (vendor == "**A" && status != "ALL STATUS")
            {
                com.CommandText = "select * from s_goods_receipt where Status = '"+ status + "'  and DocDate = '" + docdate.Date + "'";
            }
            else
            {
                com.CommandText = "select * from s_goods_receipt where VendorCode = '" + vendor + "' and Status = '" + status + "' and DocDate = '" + docdate.Date + "'";
            }
            dr = com.ExecuteReader();

            while (dr.Read())
            {
                var greceipt = new GoodsReceipt();
                greceipt.BranchID = Convert.ToString(dr["BranchID"]);
                greceipt.GRCode = Convert.ToString(dr["GRCode"]);
                greceipt.Status = Convert.ToString(dr["Status"]);
                greceipt.DocDate = Convert.ToDateTime(dr["DocDate"]);
                greceipt.VendorCode = Convert.ToString(dr["VendorCode"]);
                greceipt.VendorName = Convert.ToString(dr["VendorName"]);
                greceipt.Remarks = Convert.ToString(dr["Remarks"]);
                gr.Add(greceipt);
            }
            con.Close();
            return gr;
        }

        private List<GoodsReceiptDetails> GetGoodsReceiptDetails(string vendor, DateTime docdate, string status)
        {
            List<GoodsReceiptDetails> gr = new List<GoodsReceiptDetails>();
            connectionString();
            con.Open();
            com.Connection = con;
            if (vendor == "**A" && status == "ALL STATUS")
            {
                com.CommandText = "select * from s_goods_receipt a inner join s_goods_receipt_details b on a.GRCode = b.GRCode where DocDate = '" + docdate.Date + "'";
            }
            else if (vendor != "**A" && status == "ALL STATUS")
            {
                com.CommandText = "select * from s_goods_receipt a inner join s_goods_receipt_details b on a.GRCode = b.GRCode where VendorCode = '" + vendor + "' and DocDate = '" + docdate.Date + "'";
            }
            else if (vendor == "**A" && status != "ALL STATUS")
            {
                com.CommandText = "select * from s_goods_receipt a inner join s_goods_receipt_details b on a.GRCode = b.GRCode where Status = '" + status + "'  and DocDate = '" + docdate.Date + "'";
            }
            else
            {
                com.CommandText = "select * from s_goods_receipt a inner join s_goods_receipt_details b on a.GRCode = b.GRCode where VendorCode = '" + vendor + "' and Status = '" + status + "' and DocDate = '" + docdate.Date + "'";
            }
            dr = com.ExecuteReader();

            while (dr.Read())
            {
                var greceipt = new GoodsReceiptDetails();
                greceipt.GRCode = Convert.ToString(dr["GRCode"]);
                greceipt.Status = Convert.ToString(dr["Status"]);
                greceipt.DocDate = Convert.ToDateTime(dr["DocDate"]);
                greceipt.VendorCode = Convert.ToString(dr["VendorCode"]);
                greceipt.VendorName = Convert.ToString(dr["VendorName"]);
                greceipt.ItemCode = Convert.ToString(dr["ItemCode"]);
                greceipt.Description = Convert.ToString(dr["Description"]);
                greceipt.CurrentPrice = Convert.ToString(dr["CurrentPrice"]);
                greceipt.Quantity = Convert.ToInt32(dr["Quantity"]);
                gr.Add(greceipt);
            }
            con.Close();
            return gr;
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

        private List<ItemModel> GetItem()
        {
            List<ItemModel> Item = new List<ItemModel>();
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "select * from s_item_master order by ItemCode";
            dr = com.ExecuteReader();
            while (dr.Read())
            {
                var inven = new ItemModel();
                inven.ItemCode = Convert.ToString(dr["ItemCode"]);
                inven.VendorCode = Convert.ToString(dr["VendorCode"]);
                inven.Description = Convert.ToString(dr["Description"]);
                inven.CurrentPrice = Convert.ToString(dr["CurrentPrice"]);
                inven.Brand = Convert.ToString(dr["Brand"]);
                Item.Add(inven);
            }
            con.Close();
            return Item;
        }

        private List<Details> GetDetails(string GRCode)
        {
            List<Details> Item = new List<Details>();
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "select * from s_goods_receipt_details where GRCode = '"+ GRCode + "' order by ItemCode";
            dr = com.ExecuteReader();
            while (dr.Read())
            {
                var inven = new Details();
                inven.ItemCode = Convert.ToString(dr["ItemCode"]);
                inven.ItemCode = Convert.ToString(dr["ItemCode"]);
                inven.Description = Convert.ToString(dr["Description"]);
                inven.CurrentPrice = Convert.ToString(dr["CurrentPrice"]);
                inven.Quantity = Convert.ToInt32(dr["Quantity"]);
                Item.Add(inven);
            }
            con.Close();
            return Item;
        }

        private VendorBack GetVendorBack(string code, DateTime tgl)
        {
            VendorBack VendorBack = new VendorBack();
            VendorBack.VendorCode = code;
            string iDate = tgl.Month + "/" + tgl.Day + "/" + tgl.Year;
            DateTime oDate = Convert.ToDateTime(iDate);
            VendorBack.DOB = oDate;
            return VendorBack;
        }

        private StatusBack GetStatusBack(string status, string remarks, string grcode)
        {
            StatusBack statusBack = new StatusBack();
            statusBack.Status = status;
            statusBack.Remarks = remarks;
            statusBack.GRCode = grcode;
            return statusBack;
        }

        public ActionResult DownloadPDF(string vendor,DateTime docdate, string status)
        {
            try
            {
                //var model = new GeneratePDFModel();
                ViewModelGR mymodel = new ViewModelGR();
                mymodel.GoodsReceipt = GetGoodsReceipt(vendor.Substring(0, 3), docdate, status);
                mymodel.GoodsReceiptDetails = GetGoodsReceiptDetails(vendor.Substring(0, 3), docdate, status);
                mymodel.Vendor = GetVendor();
                mymodel.VendorBack = GetVendorBack(vendor.Substring(0, 3), docdate);
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
                if (mymodel.GoodsReceiptDetails.Count > 0)
                {
                    return new Rotativa.MVC.ViewAsPdf("GeneratePDFGR", mymodel) { FileName = "GoodsReceipt.pdf" };
                }
                else
                {
                    return View("Index", mymodel);
                }
                
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public ActionResult DownloadExcel(string vendor, DateTime docdate, string status)
        {

           
            ViewModelGR mymodel = new ViewModelGR();
            mymodel.GoodsReceipt = GetGoodsReceipt(vendor.Substring(0, 3), docdate, status);
            mymodel.GoodsReceiptDetails = GetGoodsReceiptDetails(vendor.Substring(0, 3), docdate, status);
            mymodel.Vendor = GetVendor();
            mymodel.VendorBack = GetVendorBack(vendor.Substring(0, 3), docdate);

            if (mymodel.GoodsReceiptDetails.Count > 0)
            {
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                worksheet.Cells[1, 1].Value = "GRCode";
                worksheet.Cells[1, 2].Value = "Status";
                worksheet.Cells[1, 3].Value = "DocDate";
                worksheet.Cells[1, 4].Value = "VendorCode";
                worksheet.Cells[1, 5].Value = "VendorName";
                worksheet.Cells[1, 6].Value = "ItemCode";
                worksheet.Cells[1, 7].Value = "Description";
                worksheet.Cells[1, 8].Value = "CurrentPrice";
                worksheet.Cells[1, 9].Value = "Quantity";
                int row = 2;
                foreach (var item in mymodel.GoodsReceiptDetails)
                {

                    worksheet.Cells[row, 1] = item.GRCode;
                    worksheet.Cells[row, 2] = item.Status;
                    worksheet.Cells[row, 3] = item.DocDate;
                    worksheet.Cells[row, 4] = item.VendorCode;
                    worksheet.Cells[row, 5] = item.VendorName;
                    worksheet.Cells[row, 6] = item.ItemCode;
                    worksheet.Cells[row, 7] = item.Description;
                    worksheet.Cells[row, 8] = item.CurrentPrice;
                    worksheet.Cells[row, 9] = item.Quantity;
                    row++;
                }


                workbook.SaveAs(@"D:\Excel\sample.xlsx");
                workbook.Close();
                Marshal.ReleaseComObject(workbook);

                application.Quit();
                Marshal.FinalReleaseComObject(application);

                //return RedirectToAction("Index");
                return File(@"D:\Excel\sample.xlsx", "Application/Excel", "GoodsReceipt.xlsx");

            }
            else
            {
                return View("Index", mymodel);
            }


           
        }
    }
}
