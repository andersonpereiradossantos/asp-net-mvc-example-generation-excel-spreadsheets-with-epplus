using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;

namespace ExampleExcel.Controllers
{
    public class ExportExcelController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public FileResult ExportExcel()
        {
            //Define the packager
            using (ExcelPackage package = new ExcelPackage())
            {
                //Create Worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("NameWorksheet");

                //Define the header customize
                worksheet.Cells["A1:G1"].Merge = true;
                worksheet.Cells["A1"].Value = "SALES OF YEAR";
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells["A2"].Value = "Date";
                worksheet.Cells["B2"].Value = "Product Code";
                worksheet.Cells["C2"].Value = "Location Sale";
                worksheet.Cells["D2"].Value = "Employee";
                worksheet.Cells["E2"].Value = "Amount";
                worksheet.Cells["F2"].Value = "Price";
                worksheet.Cells["G2"].Value = "Total";
                worksheet.Cells["A1:G2"].Style.Font.Bold = true;

                //Define the content 
                worksheet.Cells["A3"].Value = "2021-05-09";
                worksheet.Cells["B3"].Value = "6579";
                worksheet.Cells["C3"].Value = "California";
                worksheet.Cells["D3"].Value = "Santos, Anderson";
                worksheet.Cells["E3"].Value = "10";
                worksheet.Cells["F3"].Value = "500.00";
                worksheet.Cells["G3"].Value = "5,000.00";


                //Format cell to date and numbers
                worksheet.Cells["A3"].Style.Numberformat.Format = "MM/dd/yyyy";
                worksheet.Cells["B3"].Style.Numberformat.Format = "0";
                worksheet.Cells["E3"].Style.Numberformat.Format = "0";
                worksheet.Cells["F3:G3"].Style.Numberformat.Format = "0.00";

                //Define auto-adjustable columns for content
                worksheet.Cells[$"A1:G3"].AutoFitColumns();

                //Export the file
                return File(package.GetAsByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "worksheet.xlsx");
            }
        }
    }
}