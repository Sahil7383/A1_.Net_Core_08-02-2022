namespace A1_.Net_Core_08_02_2022.Models
{
    public class StaticDB
    {
        public static List<Employee> employees = new List<Employee>()
        {
            new Employee(){EmployeeId = 1 , EmployeeName = "Sahil", EmployeeAge = 20},
            new Employee(){EmployeeId = 2 , EmployeeName = "Sagar", EmployeeAge = 21},
        };
    }
}

