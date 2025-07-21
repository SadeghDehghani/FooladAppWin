using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FooladAppWin.Classes
{
    public class PersonnelRecord
    {
        public string PersonnelNumber { get; set; }
        public string FullName { get; set; }
        public string  Date { get; set; }
        public string Day { get; set; }
        public string Time { get; set; } // به‌صورت رشته (مثل "08:15")
        public string Status { get; set; }
    }
}
