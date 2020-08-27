using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using System.Web.Services.Description;

namespace ReadExcelFile.Logs
{
    public static class MessageLog
    {
        public static void LogError(string message)
        {
            try
            {
                string path = "~/Logs/" + DateTime.Today.ToString("dd-MM-yy") + ".text";
                if (!File.Exists(System.Web.HttpContext.Current.Server.MapPath(path)))
                {
                    File.Create(System.Web.HttpContext.Current.Server.MapPath(path)).Close();
                }
                using (StreamWriter w = File.AppendText(System.Web.HttpContext.Current.Server.MapPath(path)))
                {
                    w.WriteLine("\r\nlog Entry : ");
                    w.WriteLine("{0}", DateTime.Now.ToString(CultureInfo.InvariantCulture));
                    string err = "Error in: " + System.Web.HttpContext.Current.Request.Url.ToString() + ". \n\nError Message:" + message;
                    w.WriteLine(err);
                    w.WriteLine("========================================");
                    w.Flush();
                    w.Close();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}