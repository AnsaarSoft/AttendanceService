using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceService.Models
{
    public class LeaveStructure
    {
        public int ID { get; set; }
        public string LeaveType { get; set; }
        public decimal Balance { get; set; }
        public int Priority { get; set; }
    }
}
