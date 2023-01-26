using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ReportGen.Models
{
    /// <summary>
    /// Gathers information about individual RFIs
    /// </summary>
    public class RfiInfo
    {

        public string Subject { get; set; }
        public string Name { get; set; }
        public bool IsAnswered { get; set; }
        public DateTime? AnswerDate { get; set; }
        public DateTime? SentDate { get; set; }

        /// <summary>
        /// Constructor, set up specific job
        /// </summary>
        public RfiInfo()
        {

        }

    }
}
