using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace Asp.netCoreMVCCRUD.Models
{
    public class Employee
    {
        internal object db_Employee;

        [Key]
        public int EmployeeId { get; set; }
        [Column(TypeName ="nvarchar(250)")]
        [Required (ErrorMessage ="This Field is required")]
        [DisplayName("Full Name")]
        public string Fullname { get; set; }
        [Column(TypeName = "nvarchar(10)")]
        [DisplayName("Emp. Code")]
        public string EmpCode { get; set; }
        [Column(TypeName = "nvarchar(100)")]
       
        public string Position { get; set; }
        [Column(TypeName = "nvarchar(100)")]
        [DisplayName("Office Location")]
        public string OfficeLocation { get; set; }
    }
}
