using Export_to_Excel.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Export_to_Excel.Controllers
{
    public class HomeController : Controller
    {
        EmployeedetailsEntities db = new EmployeedetailsEntities();

        public ActionResult Index()
        {
            List<EmployeeViewModel> emplist = new List<EmployeeViewModel>();
            using (var context = new EmployeedetailsEntities())
            {
                var employeeInfoList = context.SP_SelectEmployee().ToList();
                emplist = employeeInfoList.Select(employee => new EmployeeViewModel
                {
                    empid = employee.empid,
                    Name = employee.Name,
                    DOJ = employee.DOJ,
                    Designation = employee.Designation,
                    Salary = employee.Salary
                }).ToList();
                return View(emplist);
            }
        }

        public ActionResult ExportToExcel()
        {
            try
            {
                var employeeData = db.Employees.ToList();

                byte[] excelData;
                using (MemoryStream ms = new MemoryStream())
                {
                    ExcelPackage.LicenseContext = LicenseContext.Commercial;
                    using (ExcelPackage package = new ExcelPackage(ms))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("EmployeeData");

                        
                        worksheet.Cells[1, 1].Value = "EmpID";
                        worksheet.Cells[1, 2].Value = "Name";
                        worksheet.Cells[1, 3].Value = "DOJ";
                        worksheet.Cells[1, 4].Value = "Designation";
                        worksheet.Cells[1, 5].Value = "Salary";

                        int row = 2;
                        foreach (var employee in employeeData)
                        {
                            worksheet.Cells[row, 1].Value = employee.empid;
                            worksheet.Cells[row, 2].Value = employee.Name;
                            worksheet.Cells[row, 3].Value = employee.DOJ.ToString("yyyy-MM-dd"); 
                            worksheet.Cells[row, 4].Value = employee.Designation;
                            worksheet.Cells[row, 5].Value = employee.Salary;
                            row++;
                        }

                        
                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                        excelData = package.GetAsByteArray();
                        
                    }
                }

                return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EmployeeData.xlsx");
            }
            catch (Exception ex)
            {

                TempData["ErrorMessage"] = ex.Message;
               
                return RedirectToAction("Index");
            }
        }


    }
}

 
       
