using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.Models
{
    public class SearchInfo : Spool
    {
        public SearchInfo()
        {
            DisplayInfo = this.WorkOrder;
        }
        public bool IsJobComplete { get; set; }
        public string DisplayInfo { get; set; }

    }

}
