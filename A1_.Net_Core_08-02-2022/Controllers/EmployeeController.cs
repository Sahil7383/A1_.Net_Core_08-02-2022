using A1_.Net_Core_08_02_2022.Models;
using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using System.Data;

namespace A1_.Net_Core_08_02_2022.Controllers
{
    public class EmployeeController : Controller
    {
        
        public IActionResult EmployeeList()
        {
            
            return View(StaticDB.employees);
        }

        public IActionResult Create()
        {
            return View();
        }

        [HttpPost]
        public IActionResult CreateEmployee(Employee emp)
        {
            StaticDB.employees.Add(emp);

            return RedirectToAction("EmployeeList");
        }

        [HttpGet]
        public IActionResult Edit(int id)
        {
            var emp = StaticDB.employees.Where(x => x.EmployeeId == id).FirstOrDefault();
            return View(emp);
        }

        [HttpPost]
        public IActionResult Edit(Employee updatedEmp)
        {
            var existingEmployee = StaticDB.employees.Where(x => x.EmployeeId == updatedEmp.EmployeeId).FirstOrDefault();
            if(existingEmployee == null)
            {
                return View();
            }
            else
            {
                existingEmployee.EmployeeName = updatedEmp.EmployeeName;
                existingEmployee.EmployeeAge = updatedEmp.EmployeeAge;
                return RedirectToAction("EmployeeList");
            }
        }

        public IActionResult Detail(int Id)
        {
            var existingEmployee  = StaticDB.employees.Where(x => x.EmployeeId == Id).FirstOrDefault();
            return View(existingEmployee);
        }

        public IActionResult Delete(int Id)
        {
            var existingEmployee = StaticDB.employees.Where(x => x.EmployeeId == Id).FirstOrDefault();
            StaticDB.employees.Remove(existingEmployee);
            return RedirectToAction("EmployeeList");
        }


       public IActionResult ExporttoExcel()
        {
            using(var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Employee");
                var currentRow = 1;

                // Region Header
                worksheet.Cell(currentRow,1).Value = "EmployeeId";
                worksheet.Cell(currentRow,2).Value = "EmployeeName";
                worksheet.Cell(currentRow,3).Value = "EmployeeAge";
                worksheet.Columns().AdjustToContents();
                worksheet.Rows().AdjustToContents();

                //endRegion

                // RegionBody 
                foreach(var employee in StaticDB.employees)
                {
                    currentRow++;
                    worksheet.Cell(currentRow,1).Value = employee.EmployeeId;
                    worksheet.Cell(currentRow,2).Value = employee.EmployeeName;
                    worksheet.Cell(currentRow, 3).Value = employee.EmployeeAge;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "Employee.xlsx"
                        );
                }
            }
        }

    }
}
