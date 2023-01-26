using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;

namespace ReportGen.Models
{
    public static class JobNav
    {
        public static string ACTIVE_JOBS_PATH = @"I:\Fabrication\17th ST\03 - Work Orders\01 - Active Jobs";
        static readonly string RfiPathAppend = @"3.0 Document Control\3.5 RFI's";
        static readonly string MrPathAppend = @"6.0 Material & Reqs\7.2 Material Requisitions";
        static readonly string MrPathAppend2 = @"6.0 Material & Reqs\7.2 Material Requisitions\Materials Requisition";
        static readonly string ItpPathAppend = @"11.0 Quality\ITP";
        static readonly string ActiveJobPath = @"I:\Fabrication\17th ST\03 - Work Orders\01 - Active Jobs";
        static readonly string CompletedJobPath = @"I:\Fabrication\17th ST\03 - Work Orders\02 - Completed Jobs";
        static readonly string oldJobSearchPattern = "18*";
        
        /*
        //Home Variables
        static readonly string RfiPathAppend = @"3.0 Document Control\3.5 RFI's";
        static readonly string MrPathAppend = @"6.0 Material & Reqs\7.2 Material Requisitions";
        static readonly string MrPathAppend2 = @"6.0 Material & Reqs\7.2 Material Requisitions\Materials Requisition";
        static readonly string ItpPathAppend = @"11.0 Quality\ITP";
        static readonly string ActiveJobPath = @"D:\Drive\PD\work\Willbros\9-10-18";
        static readonly string CompletedJobPath = @"I:\Fabrication\17th ST\03 - Work Orders\02 - Completed Jobs";
        static readonly string oldJobSearchPattern = "18*";
        */

        public static List<RfiInfo> GetRfiInfos(string workorder, bool searchOldJobs = true)
        {
            if (GetJobDir(workorder, searchOldJobs) == null)
                return null;

            string jobPath = GetJobDir(workorder, searchOldJobs).FullName;

            DirectoryInfo rfiDir = new DirectoryInfo(Path.Combine(jobPath, RfiPathAppend));

            if (!rfiDir.Exists)
                return null;

            DirectoryInfo[] rfiFolders = rfiDir.GetDirectories("*" + workorder + "*");

            if(rfiFolders.Length == 0)
            {
                return null;
            }

            List<RfiInfo> rfiInfos = new List<RfiInfo>();

            //Loop through every legit RFI and append information
            foreach(DirectoryInfo rfi in rfiFolders)
            {

                rfiInfos.Add(new RfiInfo()
                {
                    Name = rfi.Name,
                    IsAnswered = rfi.EnumerateFiles("*.*", SearchOption.TopDirectoryOnly).Where(s=>(s.Name.ToUpper().Contains("SIGNED") || s.Name.ToUpper().Contains("APPROVED") || s.Name.ToUpper().Contains("RESPONSE"))).Count() > 0,
                    AnswerDate = rfi.EnumerateFiles("*.*", SearchOption.TopDirectoryOnly).Where(s => (s.Name.ToUpper().Contains("SIGNED") || s.Name.ToUpper().Contains("APPROVED") || s.Name.ToUpper().Contains("RESPONSE"))).Count() > 0  ?
                    (DateTime?)rfi.EnumerateFiles("*.*", SearchOption.TopDirectoryOnly).Where(s => (s.Name.ToUpper().Contains("SIGNED") || s.Name.ToUpper().Contains("APPROVED") || s.Name.ToUpper().Contains("RESPONSE"))).First().LastWriteTime : null,
                    SentDate = rfi.GetFiles().Where(f => (!f.Name.ToUpper().Contains("RESPONSE")) && (f.Name.Contains(".pdf")) && (f.Name.Contains(workorder))).FirstOrDefault() == null ?
                    (DateTime?)null : rfi.GetFiles().Where(f => (!f.Name.ToUpper().Contains("RESPONSE")) && (f.Name.Contains(".pdf")) && (f.Name.Contains(workorder)) && (!f.Name.ToUpper().Contains("SIGNED"))).First().LastWriteTime,
                });

            }

            return rfiInfos;
   
        }

        public static DirectoryInfo GetItpDir(string workorder, bool searchOldJobs = true)
        {
            string jobPath = GetJobDir(workorder, searchOldJobs).FullName;
            DirectoryInfo itpDir = new DirectoryInfo(Path.Combine(jobPath, ItpPathAppend));

            return itpDir;
        }

        public static DirectoryInfo GetJobDir(string workorder, bool searchOldJobs = true)
        {
            //Search for job folder in Active job
            if (!Directory.Exists(ActiveJobPath))
                return null;

            DirectoryInfo[] activeJobDirs = new DirectoryInfo(ActiveJobPath).GetDirectories();

            DirectoryInfo jobDir = activeJobDirs.Where(j => j.Name.Contains(workorder)).FirstOrDefault();

            if (jobDir != null || searchOldJobs == false)
            {
                return jobDir;
            }
            else
            {
                //Search for job in Completed jobs if not found in active (SLOW AF)
                DirectoryInfo[] oldJobDirs = new DirectoryInfo(CompletedJobPath).GetDirectories(oldJobSearchPattern, SearchOption.AllDirectories);

                jobDir = oldJobDirs.FirstOrDefault(j => j.Name.Contains(workorder));

                return jobDir;
            }
        }

        public static List<string> GetActiveJobsList()
        {
            List<string> activeJobsList = new List<string>();
            var activeJobsDir = new DirectoryInfo(JobNav.ACTIVE_JOBS_PATH);
            var activeJobs = activeJobsDir.GetDirectories();

            foreach (DirectoryInfo job in activeJobs)
            {

                if (String.IsNullOrWhiteSpace(Regex.Match(job.Name.Trim(), @"^\d{4,}").ToString()))
                {
                    continue;
                }
                else
                {
                    string workOrder = Regex.Match(job.Name.Trim(), @"^\d{4,}").ToString();
                    activeJobsList.Add(workOrder);
                }

            }

            return activeJobsList;

        }
    }
}
