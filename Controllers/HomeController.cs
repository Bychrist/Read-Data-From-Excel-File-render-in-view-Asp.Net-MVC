﻿using System;
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
using System.Net;
using System.Text;

namespace ReadExcelFile.Controllers
{
    public class HomeController : Controller
    {
      
        public string FtpUserID { get; set; } = "Usernamehere";
        public string FtpPassword { get; set; } = "your password";

        [HttpGet]
        public ActionResult Index()
        {
            ViewBag.MarkUsers = null;
            return View();
        }

        [HttpGet]
        public ActionResult About()
        {
            return View();
        }



        public ActionResult About(HttpPostedFileBase postedFile)
        {
            //FTP Server URL.
            string ftp = "ftp://yourserverurl.com/";

            //FTP Folder name. Leave blank if you want to upload to root folder.
            string ftpFolder = "FtpUpload/";

            byte[] fileBytes = null;

            //Read the FileName and convert it to Byte array.
            string fileName = Path.GetFileName(postedFile.FileName);
            using (StreamReader fileStream = new StreamReader(postedFile.InputStream))
            {
                fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());
                fileStream.Close();
            }

            try
            {
                //Create FTP Request.
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftp + ftpFolder + fileName);
                request.Method = WebRequestMethods.Ftp.UploadFile;

                //Enter FTP Server credentials.
                request.Credentials = new NetworkCredential(FtpUserID, FtpPassword);
                request.ContentLength = fileBytes.Length;
                request.UsePassive = true;
                request.UseBinary = true;
                request.ServicePoint.ConnectionLimit = fileBytes.Length;
                request.EnableSsl = false;

                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(fileBytes, 0, fileBytes.Length);
                    requestStream.Close();
                }

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                response.Close();
            }
            catch (WebException ex)
            {
                throw new Exception((ex.Response as FtpWebResponse).StatusDescription);
            }


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
                    string requestSent = "<?xml version=\"1.0\" encoding=\"UTF - 8\" standalone=\"yes\"?><PaymentRequestCommand><ScheduleId>UAT_1</ScheduleId><ClientId>NIBSS_V2001</ClientId><DebitBankCode>044</DebitBankCode><DebitAccountNumber>0123456789</DebitAccountNumber></PaymentRequestCommand>";
                    TempData["message"] = "Upload was successful";
                    MessageLog.LogError("Upload was successful", requestSent, "16");
                    return View(nameof(Index));

                }

            

                TempData["message"] = "No file was uploaded";
                string request = "<?xml version=\"1.0\" encoding=\"UTF - 8\" standalone=\"yes\"?><PaymentRequestCommand><ScheduleId>UAT_1</ScheduleId><ClientId>NIBSS_V2001</ClientId><DebitBankCode>044</DebitBankCode><DebitAccountNumber>0123456789</DebitAccountNumber></PaymentRequestCommand>";
                MessageLog.LogError("No file was uploaded", request,"03");
                return View(nameof(Index));

            }
            catch (Exception e)
            {

                TempData["message"] = e.Message;
                string request= "<?xml version=\"1.0\" encoding=\"UTF - 8\" standalone=\"yes\"?><PaymentRequestCommand><ScheduleId>UAT_1</ScheduleId><ClientId>NIBSS_V2001</ClientId><DebitBankCode>044</DebitBankCode><DebitAccountNumber>0123456789</DebitAccountNumber></PaymentRequestCommand>";
                MessageLog.LogError(e.Message.ToString(),request,"09");
                return View();
            }

        }



    }


  












}