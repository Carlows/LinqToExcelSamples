using LinqToExcel;
using LinqToExcelSample.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LinqToExcelSample.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View(new List<Alumno>());
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                string extension = Path.GetExtension(file.FileName);
                if (extension.Equals(".xls") || extension.Equals(".xlsx") || extension.Equals(".csv"))
                {
                    try 
                    {  
                        // upload the file
                        string fileName = string.Format("Excel-{0:dd-MM-yyyy-HH-mm-ss}{1}", DateTime.Now, extension);                        
                        string path = Path.Combine(Server.MapPath("~/ExcelTemp"), fileName);  
                        file.SaveAs(path);  

                        // get the file
                        var excel = new ExcelQueryFactory(path);
                        var alumnos = excel.Worksheet<Alumno>().ToList();

                        // delete the file
                        if (System.IO.File.Exists(path))
                        {
                            System.IO.File.Delete(path);
                        }

                        // send or do something with data
                        return View(alumnos);
                    }  
                    catch (Exception ex)  
                    {  
                        ViewBag.Message = "ERROR:" + ex.Message.ToString();  
                    }  
                }
            }
            else
            {
                ViewBag.Message = "You have not specified a file.";
            }

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}