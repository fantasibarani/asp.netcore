using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Asp.netCoreMVCCRUD.Models;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using Microsoft.AspNetCore.Http;
using System.IO;

namespace Asp.netCoreMVCCRUD.Controllers
{
    public class HomeController : Controller
    {
     

        /*Untuk koneksi ke dalam database secara manual*/
        //tambahan untuk login authentication
        SqlConnection con = new SqlConnection();
        SqlCommand com = new SqlCommand();
        SqlDataReader dr;
        //akhir authentication
    
        excel_employee db_Employee = new excel_employee();
        string ConnectionInformation = "Data Source = FANTAR-DV3478; Initial Catalog = EmployeeDB; Integrated Security = true";
        SqlDataReader myReader;

        
        //akhir koneksi ke database secara manual
        private readonly ILogger<HomeController> _logger;

        

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        //untuk index post
        [HttpPost]
        public IActionResult Index(Akun UI)
        {
            /* untuk koneksi ke dalam tabel*/
            SqlConnection MainConnection = new SqlConnection(ConnectionInformation);
            MainConnection.Open();
            string MyCommand = "INSERT INTO Akun (Nama , Password) VALUES ('" + UI.Nama + "' , '" + UI.Password + "')";
            SqlCommand myCommand = new SqlCommand(MyCommand, MainConnection);
            myCommand.ExecuteNonQuery();
            ViewBag.Message = "Haii..." + UI.Nama + " Akun anda sudah aktif. " + " Pilih Login untuk masuk ";
            MainConnection.Close();
            /* untuk koneksi ke dalam tabel*/

            return View();
        }
        //akhir untuk index post
        public ActionResult Login()
        {
            return View();
        }
        //Untuk proses Login
        [HttpPost]
        public ActionResult Login(Akun model)
        {
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "SELECT * from Akun WHERE Nama='" + model.Nama + "'and password='" + model.Password + "'";
            dr = com.ExecuteReader();
            if (dr.Read())
            {
                con.Close();
                return View("Login_Success");
            }
            else
            {
                con.Close();
                ViewBag.Message = "Username dan Password tidak terdaftar, Silahkan daftar dahulu!!";
                return View();

            }
        }

        private void connectionString()
        {
            con.ConnectionString = @"Data Source = FANTAR-DV3478; Initial Catalog =EmployeeDB; Integrated Security= True";
        }

        //Akhir untuk proses login

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        //Untuk Export Excel
        public ActionResult ExportToExcel()
        {
            //----------------sql----------------
            connectionString();
            con.Open();
            com.Connection = con;
            com.CommandText = "SELECT * FROM Employees";
         
            //----------------sql----------------

            var data = com.CommandText.ToList();
            var stream = new MemoryStream();

            using (var package = new ExcelPackage(stream))
            {
                var sheet = package.Workbook.Worksheets.Add("Fullname");
                sheet.Cells.LoadFromCollection(com.CommandText, true);

                package.Save();

            }

            stream.Position = 0;

            var filename = $"Loai_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);



        }

        //akhir untuk export ke excel
    }
}
