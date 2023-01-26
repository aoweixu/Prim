using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.Models.CCOImporting
{
    class ErrorLog
    {
        public string Type { get; set; }
        public string MachineType { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }

    }
}
