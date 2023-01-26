using Caliburn.Micro;
using ReportGen.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.ViewModels
{
    public class StatusViewOneViewModel : Screen
    {
        public StatusViewOneViewModel(JobInfo job)
        {
            Dc = new DataConnection();
            this.SelectedJobInfo = job;
            if (job != null)
            {
                SpoolList = new BindableCollection<Spool>();
                SpoolList.AddRange(DataConnection.AssignSpoolPropertiesFromData(Dc.GetSpoolsData(job.WorkOrder)));
            }

        }

        private JobInfo _selectedJobInfo;
        public BindableCollection<Spool> SpoolList { get; set; }
        public DataConnection Dc { get; set; }

        public JobInfo SelectedJobInfo
        {
            get { return _selectedJobInfo; }
            set
            {
                _selectedJobInfo = value;
                NotifyOfPropertyChange(() => SelectedJobInfo);
            }
        }
    }
}
