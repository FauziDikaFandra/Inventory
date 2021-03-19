using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using New_Web.Models;
using System.Data.SqlClient;
namespace New_Web.Controllers
{
    
    public class HomeController : Controller
    {
        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;
        public ActionResult Index()
        {
            //string test = Session["login"] as string;
            if (!string.IsNullOrEmpty(Session["login"] as string))
            {
                return View();
            }
            else
            {
                var eRr = new Models.Error() { txt = "" };
                return View("Login", eRr);
            }
        }



       
        void connectionString()
        {
            con.ConnectionString = "Data Source=DESKTOP-IL11GFI;Initial Catalog=New_Web;User ID=sa;Password=star;";
        }
        [HttpPost]
        public ActionResult Add(Login data)
        {
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "select * from s_user where UserName = '" + data.UserName + "' and Password = '" + data.Password + "'";
            dr = com.ExecuteReader();
            if (dr.Read())
            {
                Session["login"] = data.UserName;
                con.Close();
                return RedirectToAction("Index", "Home");
            }
            else
            {
                con.Close();
                var eRr = new Models.Error() { txt = "User or Password Invalid !" };
                return View("Login", eRr);

            }


            //var Awal = new Login() { UserName = data.UserName, Password = data.Password };


            //var movie = new Login() { Name = "One Piece" };

            //return View(movie);
            //return Content("Hello World!");
            //return HttpNotFound();
            //return new EmptyResult();
            //return RedirectToAction("Index","Home", new { page = 1, sortBy = "name" }) ;

        }
        [HttpGet]
        public ActionResult Login()
        {
            var eRr = new Models.Error() { txt = "" };
            return View("Login", eRr);
            //return View();
            //var movie = new Login() { Name = "One Piece" };

            //return View(movie);
            //return Content("Hello World!");
            //return HttpNotFound();
            //return new EmptyResult();
            //return RedirectToAction("Index","Home", new { page = 1, sortBy = "name" }) ;

        }
        public ActionResult Logout()
        {
            Session.Clear();            
            return RedirectToAction("Index", "Home");
            //var movie = new Login() { Name = "One Piece" };

            //return View(movie);
            //return Content("Hello World!");
            //return HttpNotFound();
            //return new EmptyResult();
            //return RedirectToAction("Index","Home", new { page = 1, sortBy = "name" }) ;

        }

        
    }
}