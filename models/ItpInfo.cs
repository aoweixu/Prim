using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportGen.Attributes;
using System.IO;
using System.Reflection;
using System.ComponentModel;

namespace ReportGen.Models
{

    public class ItpInfo
    {
        public ItpInfo(string workorder)
        {
            GetItpStatusAndDate(workorder);
        }

        public string ItpStatusStr { get; set; }
        public DateTime? ItpDate { get; set; }

        public void GetItpStatusAndDate(string workorder)
        {
            string WorkOrder = workorder;

            if (WorkOrder == null)
            {
                ItpStatusStr = ItpStatus.NotSent.GetDescription();
                ItpDate = null;
                return;
            }

            var activejobDir = new DirectoryInfo(@"I:\Fabrication\17th ST\03 - Work Orders\01 - Active Jobs");
            DirectoryInfo[] activeJobDirs = null;
            try
            {
                activeJobDirs = activejobDir.GetDirectories();
            }
            catch
            {
                ItpDate = null;
                ItpStatusStr = ItpStatus.NotFound.GetDescription();
                return;
            }


            DirectoryInfo itpDir = null;

            foreach (DirectoryInfo dir in activeJobDirs)
            {
                if (dir.Name.Contains(WorkOrder))
                {
                    itpDir = new DirectoryInfo(Path.Combine(dir.FullName, @"11.0 Quality\ITP"));
                    break;
                }
            }

            if (itpDir == null || !itpDir.Exists)
            {
                ItpDate = null;
                ItpStatusStr = ItpStatus.NotFound.GetDescription();
                return;
            }

            List<string> acceptStringList = new List<string>
            {
                //List of words in caps for accepted ITPs
                "APPROVED",
                "ACC",
                "SIGNED",
            };

            foreach (FileInfo file in itpDir.GetFiles())
            {
                if (acceptStringList.Any(s => file.Name.ToUpper().Contains(s)))
                {
                    ItpDate = GetItpDate(file);
                    ItpStatusStr = ItpStatus.Approved.GetDescription();
                    return;
                }

            }

            foreach (FileInfo file in itpDir.GetFiles())
            {
                if (file.Name.ToUpper().Contains("APPROVAL"))
                {
                    ItpDate = GetItpDate(file);
                    ItpStatusStr = ItpStatus.Sent.GetDescription();
                    return;
                }
            }

            ItpStatusStr = ItpStatus.NotSent.GetDescription();
        }

        private static DateTime GetItpDate(FileInfo file)
        {
            return file.LastWriteTime;
        }

        public enum ItpStatus
        {
            [Description("Approved")]
            Approved,
            [Description("Sent/Not Approved")]
            Sent,
            [Description("Not Sent")]
            NotSent,
            [Description("Null")]
            NoWorkOrder,
            [Description("N/A")]
            NotFound,
        }
    }
}
