﻿using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ClassWebApi.Controllers
{
    public class ClassRoomController : Controller
    {
        [Route("GenerateExcel")]
        [HttpPost]
        [Produces("application/json")]
        public IActionResult GenerateExcel([FromBody] ClassInput input)
        {
            string path = "";

            List<string> totalStudent = new List<string>();
            List<string> studentsPresent = new List<string>();
            List<string> absentNames = new List<string>();
            string csvAbsentee = "";
            string[] excelRows = input.AttendenceData.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            foreach (string row in excelRows)
            {
                string[] rowArray = row.Split('\t', StringSplitOptions.RemoveEmptyEntries);

                if (rowArray.Length >= 2)
                {
                    if (DateTime.TryParse(rowArray[0], CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateTime))
                    {
                        if (dateTime.Date == DateTime.Now.Date)
                        {
                            string[] studentDetail = rowArray[1].Split(' ', StringSplitOptions.RemoveEmptyEntries);
                            studentsPresent.Add(studentDetail[0]);
                        }
                    }
                }
            }

            switch (input.ClassName)
            {
                //case "6":
                //  path = AppDomain.CurrentDomain.BaseDirectory + "//Excel//VI A.xlsx";
                // break;
                case "7":
                    path = AppDomain.CurrentDomain.BaseDirectory + "//Excel//VII C.xlsx";
                    break;
                case "8":
                    path = AppDomain.CurrentDomain.BaseDirectory + "//Excel//8c.xlsx";
                    break;
                case "9":
                    path = AppDomain.CurrentDomain.BaseDirectory + "//Excel//9A.xlsx";
                    break;
                case "10":
                    path = AppDomain.CurrentDomain.BaseDirectory + "//Excel//10C.xlsx";
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
                                                     //int columns = worksheet.Dimension.Columns; // 7

                //loop through the worksheet rows and columns

                for (int i = 1; i <= rows; i++)
                {
                    totalStudent.Add(
                         worksheet.Cells[i, 1].Value == null ? "" : worksheet.Cells[i, 1].Value.ToString());
                }



                List<string> absentee = totalStudent.Except(studentsPresent).ToList();

                foreach (string num in absentee)
                {
                    if (int.TryParse(num, out int rollNo))
                        absentNames.Add(worksheet.Cells[rollNo, 2].Value == null ? "" : worksheet.Cells[rollNo, 2].Value.ToString());

                }
                csvAbsentee = string.Join(',', absentNames);
            }


            //var smtpClient1 = new SmtpClient("smtp.gmail.com")
            //{
            //    Port = 587,
            //    Credentials = new NetworkCredential("geetakumari@afsjal.org", "Scorpion@#$123"),
            //    EnableSsl = true,
            //};

            //var mailMessage = new MailMessage
            //{
            //    From = new MailAddress("amit26feb@yahoo.com"),
            //    Subject = input.ClassName + " absentee list | " + DateTime.Now.ToString("dd/MM/yyyy"),
            //    Body = "<h1>Hello, below rollNo were absent today</h1> <p>" + csvAbsentee + "</p>",
            //    IsBodyHtml = true,
            //};
            //mailMessage.To.Add("geetakumari@afsjal.org");
            //try
            //{
            //    SmtpClient smtpClient = new SmtpClient("smtp.mail.yahoo.com", 465);
            //    smtpClient.UseDefaultCredentials = false;
            //    smtpClient.Credentials = new NetworkCredential()
            //    {
            //        UserName = "amit26feb@yahoo.com",
            //        Password = "Scorpion@1234"
            //    };
            //    smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            //    smtpClient.EnableSsl = true;
            //    smtpClient.Send("geetakumari@afsjal.org", "geetakumari@afsjal.org", "Account verification", "hhhhhh");
            //}
            //catch (Exception ex)
            //{
            //    csvAbsentee = ex.Message;
            //}

            if (csvAbsentee.EndsWith(","))
                csvAbsentee = csvAbsentee.TrimEnd(',');
            return Ok(csvAbsentee);
        }
    }
}