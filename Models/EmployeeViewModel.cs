using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Export_to_Excel.Models
{
    public class EmployeeViewModel
    {
        public int empid { get; set; }
        public string Name { get; set; }
        public System.DateTime DOJ { get; set; }
        public string Designation { get; set; }
        public Nullable<decimal> Salary { get; set; }
    }
}