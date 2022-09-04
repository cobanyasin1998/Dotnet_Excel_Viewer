using ClosedXML.Excel;
using ExcelViewer.App.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace ExcelViewer.App.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(string a)
        {
            DataTable dt = new DataTable();

            string path = @"C:\Users\Yasin\Documents\GitHub\Dotnet_Excel_Viewer\ExcelViewer\ExcelViewer.App\Models\example.xlsx";

            using (XLWorkbook workbook = new XLWorkbook(path))
            {
                IXLWorksheet worksheet = workbook.Worksheet(1);
                bool FirstRow = true;

                string readRange = "1:1";
                foreach (IXLRow row in worksheet.RowsUsed())
                {

                    if (FirstRow)
                    {

                        readRange = string.Format("{0}:{1}", 1, row.LastCellUsed().Address.ColumnNumber);
                        foreach (IXLCell cell in row.Cells(readRange))
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        FirstRow = false;
                    }
                    else
                    {

                        dt.Rows.Add();
                        int cellIndex = 0;

                        foreach (IXLCell cell in row.Cells(readRange))
                        {
                            var hexCode = OptionalClass.CellExists<Product>(cell, cellIndex);
                            dt.Rows[dt.Rows.Count - 1][cellIndex] = hexCode + "," + cell.Value.ToString();

                            cellIndex++;
                        }
                    }

                    if (FirstRow)
                    {
                        ViewBag.Message = "Empty Excel File!";
                    }
                }


                return View(dt);
            }
        }

        public class Product
        {
            [MinLength(5)]
            [MaxLength(20)]
            public string Name { get; set; }

            [MaxLength(100)]
            public string Description { get; set; }

            public decimal Price { get; set; }


            public long SerialNumber { get; set; }

            [Required]
            public string Supplier { get; set; }

        }
    }
}
