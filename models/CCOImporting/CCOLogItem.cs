using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.Models.CCOImporting
{
    class CCOLogItem
    {
        public string WorkOrder { get; set; }
        public string ControlNo { get; set; }
        public int Duration { get; set; }
        public DateTime DateCompleted { get; set; }
        public string Machine { get; set; }

    }
}
