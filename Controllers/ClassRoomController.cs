using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace ClassWebApi.Controllers
{
    public class ClassRoomController : Controller
    {
        [Route("GenerateExcel")]
        [HttpGet]
        public IActionResult GenerateExcel(string selectedClass, string attendance)
        {
            string path = "";
            StringBuilder sb = new StringBuilder();
            switch (selectedClass)
            {
                case "6":
                    path = AppDomain.CurrentDomain.BaseDirectory + "//Excel//VI A.xlsx";
                    break;
                case "7":
                    path = @"C:\Amit\Geeta\VII C.xlsx";
                    break;
                case "8":
                    path = @"C:\Amit\Geeta\8c.xlsx";
                    break;
                case "9":
                    path = @"C:\Amit\Geeta\9A.xlsx";
                    break;
                case "10":
                    path = @"C:\Amit\Geeta\10C.xlsx";
                    break;
            }

            if (!string.IsNullOrEmpty(path))
            {
                FileInfo fileInfo = new FileInfo(path);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage package = new ExcelPackage(fileInfo);

                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                // get number of rows and columns in the sheet
                int rows = worksheet.Dimension.Rows; // 20
                int columns = worksheet.Dimension.Columns; // 7

                //loop through the worksheet rows and columns
                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= columns; j++)
                    {

                        sb.Append(worksheet.Cells[i, j].Value == null ? "" : worksheet.Cells[i, j].Value.ToString());
                        /* Do something ...*/
                    }
                }
            }

            return Ok(sb.ToString());
        }
    }
}