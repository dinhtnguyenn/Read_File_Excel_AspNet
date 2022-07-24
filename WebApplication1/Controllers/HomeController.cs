using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using WebApplication1.Models;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        List<Info> list = new List<Info>();

        [HttpGet]
        public ActionResult Index()
        {
            return View(list);
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase fileBase)
        {

            if (fileBase != null || fileBase.ContentLength > 0)
            {
                string _FileName = Path.GetFileName(fileBase.FileName);

                string path = Path.Combine(Server.MapPath("/filene/"), _FileName);

                if (System.IO.File.Exists(path))
                {
                    System.IO.File.Delete(path);
                    fileBase.SaveAs(path);
                }
                else
                {
                    fileBase.SaveAs(path);
                }

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    int j = 1;
                    var stt = xlRange.Cells[i, j++].Value.ToString();
                    var name = xlRange.Cells[i, j++].Value;
                    var classes = xlRange.Cells[i, j++].Value;

                    Info info = new Info(stt, name, classes);
                    list.Add(info);

                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            else
            {
                //khong tim thay file
                return View();
            }


            return View(list);
        }


    }
}