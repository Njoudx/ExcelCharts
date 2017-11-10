using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace Charts.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            string path = Server.MapPath("~/Content/Sheets/Book1.xlsx");
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string[] labels = new string[rowCount];
            string[] values = new string[rowCount];
            for (int i = 0; i < rowCount; i++)
            {
                
                    //for (int j = 1; j <= colCount; j++)
                    //{
                    labels[i] = xlRange.Cells[i+1, 1].Value2.ToString();
                    values[i] = xlRange.Cells[i+1, 2].Value2.ToString();
                    //if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    //{
                    //    ViewBag.Write += xlRange.Cells[i, j].Value2.ToString();
                    //}
                    //}
                
            }
            string json = "{";
            for (int i = 0; i < rowCount; i++)
            {
                json += labels[i]+":"+values[i];
                if (rowCount - i != 1)
                    json += ",";
            }
            json += "}";
            
            ViewBag.Json = json;
            ViewBag.Labels = labels;
            ViewBag.Values = values;
            
            //close and release
            xlWorkbook.Close();

            //quit and release
            xlApp.Quit();
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