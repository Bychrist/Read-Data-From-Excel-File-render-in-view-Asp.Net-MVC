using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using ReadExcelFile.Models;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using ReadExcelFile.Logs;

namespace ReadExcelFile.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            ViewBag.MarkUsers = null;
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            try
            {
                string path = Server.MapPath("~/Uploads/");
                string filePath = string.Empty;
                if (postedFile != null)
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    filePath = path + DateTime.Now.Ticks + "-" + Path.GetFileName(postedFile.FileName);
                    postedFile.SaveAs(filePath);

                    //read data from excel
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(filePath);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<MarkUser> Users = new List<MarkUser>();
                    for (int row = 2; row <= range.Rows.Count; row++)
                    {
                        MarkUser user = new MarkUser();
                        user.FullName = ((Excel.Range)range.Cells[row, 1]).Text;
                        user.Email = ((Excel.Range)range.Cells[row, 2]).Text;
                        user.Address = ((Excel.Range)range.Cells[row, 3]).Text;
                        Users.Add(user);
                    }
                  //  System.IO.File.Delete(filePath);
                    ViewBag.MarkUsers = Users;
                    TempData["message"] = "Upload was successful";
                    MessageLog.LogError("Upload was successful");
                    return View(nameof(Index));

                }

            

                TempData["message"] = "No file was uploaded";
                MessageLog.LogError("No file was uploaded");
                return View(nameof(Index));

            }
            catch (Exception e)
            {

                TempData["message"] = e.Message;
                MessageLog.LogError(e.Message.ToString());
                return View();
            }

        }



    }


  












}