using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Caliburn.Micro;
using ReportGen.Models;
using System.Data;
using System.Windows;
using MoreLinq;

namespace ReportGen.ViewModels
{
    public class ShellViewModel : Conductor<object>
    {
        #region Private Members
        private BindableCollection<JobInfo> _jobList = new BindableCollection<JobInfo>();
        private BindableCollection<JobInfo> _selectedJobList = new BindableCollection<JobInfo>();
        private BindableCollection<SearchInfo> _intermJobCollection = new BindableCollection<SearchInfo>();
        private List<SearchInfo> _searchInfoList = new List<SearchInfo>();
        private BindableCollection<Spool> _selectedJobSpoolList;
        private DataConnection _dataConnection;
        private JobInfo _selectedJobInfo;
        private Spool _selectedSpoolInfo;
        private SearchInfo _selectedSearchInfo;
        private JobInfo _selectedJobInfoFromReport;
        private string _woMatchText;
        private string _poMatchText;
        private string _dwgMatchText;
        private string _spoolMatchText;
        private bool _excludeOldJobs;
        private BindableCollection<string> _clientList = new BindableCollection<string>();
        private string _selectedClient;
        private Report _report;

        #endregion

        #region Public Properties


        public List<string> ActiveJobsList { get; set; }

        public JobInfo SelectedJobInfoFromReport
        {
            get { return _selectedJobInfoFromReport; }
            set
            {
                _selectedJobInfoFromReport = value;
                NotifyOfPropertyChange(() => SelectedJobInfoFromReport);
                NotifyOfPropertyChange(() => CanRemoveJobFromList);
            }
        }

        public BindableCollection<Spool> SelectedJobSpoolList
        {
            get { return _selectedJobSpoolList; }
            set
            {
                _selectedJobSpoolList = value;
                NotifyOfPropertyChange(() => SelectedJobSpoolList);
            }
        }


        public BindableCollection<JobInfo> SelectedJobList
        {
            get { return _selectedJobList; }
            set
            {
                _selectedJobList = value;
                NotifyOfPropertyChange(() => SelectedJobList);
                NotifyOfPropertyChange(() => CanCreateReport);
            }
        }


        public BindableCollection<JobInfo> JobList
        {
            get { return _jobList; }
            set
            {
                _jobList = value;
            }
        }

        public BindableCollection<SearchInfo> InterimJobCollection
        {
            get { return _intermJobCollection; }
            set
            {
                _intermJobCollection = value;
                NotifyOfPropertyChange(() => InterimJobCollection);
                NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
            }
        }

        public List<SearchInfo> SearchInfoList
        {
            get { return _searchInfoList; }
            set
            {
                _searchInfoList = value;
                NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
            }
        }

        public DataConnection DataConnection
        {
            get { return _dataConnection; }
            set { _dataConnection = value; }
        }

        public JobInfo SelectedJobInfo
        {
            get { return _selectedJobInfo; }
            set
            {
                _selectedJobInfo = value;
                if (value != null)
                {
                    SelectedJobSpoolList = new BindableCollection<Spool>();
                    SelectedJobSpoolList.AddRange(DataConnection.AssignSpoolPropertiesFromData(DataConnection.GetSpoolsData(value.WorkOrder)));
                }


                NotifyOfPropertyChange(() => SelectedJobInfo);
                NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
                NotifyOfPropertyChange(() => CanAddJobToReportList);
                NotifyOfPropertyChange(() => CanRemoveJobFromInterimList);
            }
        }

        public Spool SelectedSpoolInfo
        {
            get { return _selectedSpoolInfo; }
            set
            {
                _selectedSpoolInfo = value;
                NotifyOfPropertyChange(() => SelectedSpoolInfo);
            }
        }

        public SearchInfo SelectedSearchInfo
        {
            get { return _selectedSearchInfo; }
            set
            {
                if (value == null)
                {
                    SelectedJobInfo = null;
                    return;
                }
                else
                {
                    _selectedSearchInfo = value;
                    SelectedJobInfo = JobList.First(x => x.WorkOrder == SelectedSearchInfo.WorkOrder);
                }

                NotifyOfPropertyChange(() => SelectedSearchInfo);
                NotifyOfPropertyChange(() => SelectedJobInfo);
                NotifyOfPropertyChange(() => CanRemoveJobFromInterimList);

            }
        }

        //Conducts search by WO number
        public string WoMatchText
        {
            get { return _woMatchText; }
            set
            {
                _woMatchText = value;

                InterimJobCollection.Clear();

                if(String.IsNullOrWhiteSpace(value) || value.Length < 3)
                {
                    SelectedJobInfo = null;
                }
                else
                {               
                    if (ExcludeOldJobs)
                        InterimJobCollection.AddRange((from c in SearchInfoList.Where(x => x.WorkOrder.Contains(value.ToString())) join j in ActiveJobsList on c.WorkOrder equals j select c).DistinctBy(x => x.WorkOrder));
                    else
                        InterimJobCollection.AddRange(SearchInfoList.Where(x => x.WorkOrder.Contains(value.ToString())).DistinctBy(x => x.WorkOrder).ToList());

                    //Set DisplayMember to be workorder
                    foreach (SearchInfo searchInfo in InterimJobCollection)
                    {
                        searchInfo.DisplayInfo = searchInfo.WorkOrder;
                    }

                    if (InterimJobCollection.Count == 0)
                        SelectedJobInfo = null;
                    else
                    {
                        SelectedJobInfo = JobList.FirstOrDefault(x=>x.WorkOrder.Trim() == InterimJobCollection.First().WorkOrder.Trim());
                    }

                    NotifyOfPropertyChange(() => SelectedSearchInfo);
                    NotifyOfPropertyChange(() => SelectedJobInfo);
                    NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
                    NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
                    NotifyOfPropertyChange(WoMatchText);
                }
            }
        }

        //Conducts search by PO number
        public string PoMatchText
        {
            get { return _poMatchText; }
            set
            {
                _poMatchText = value;

                InterimJobCollection.Clear();

                if (String.IsNullOrWhiteSpace(value) || value.Length<3)
                {
                    SelectedJobInfo = null;
                }
                else
                {
                    if (ExcludeOldJobs)
                        InterimJobCollection.AddRange(SearchInfoList.Where(x => x.PoNumber.Contains(value.ToString())).Where(x => !x.IsJobComplete).DistinctBy(x => x.WorkOrder).ToList());
                    else
                        InterimJobCollection.AddRange(SearchInfoList.Where(x => x.PoNumber.Contains(value.ToString())).DistinctBy(x => x.WorkOrder).ToList());

                    //Set DisplayMember to be workorder
                    foreach (SearchInfo searchInfo in InterimJobCollection)
                    {
                        searchInfo.DisplayInfo = searchInfo.WorkOrder;
                    }

                    if (InterimJobCollection.Count == 0)
                        SelectedJobInfo = null;
                    else
                    {
                        SelectedJobInfo = JobList.FirstOrDefault(x => x.WorkOrder.ToString() == InterimJobCollection.FirstOrDefault().WorkOrder.ToString());
                    }
                }
                NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
                NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
                NotifyOfPropertyChange(() => PoMatchText);
            }
        }
        
        //Conducts search by Dwg No
        public string DwgMatchText
        {
            get { return _dwgMatchText; }
            set
            {
                _dwgMatchText = value;

                InterimJobCollection.Clear();
                if (String.IsNullOrWhiteSpace(value) || value.Length < 4)
                {
                    SelectedJobInfo = null;
                }
                else
                {
                    if(ExcludeOldJobs)
                        InterimJobCollection.AddRange(SearchInfoList.Where(x => x.DrawingNo.ToLower().Contains(value.ToString())).Where(x=>!x.IsJobComplete).DistinctBy(x => new { x.WorkOrder, x.ControlNo }).ToList());
                    else
                        InterimJobCollection.AddRange(SearchInfoList.Where(x => x.DrawingNo.ToLower().Contains(value.ToString())).DistinctBy(x => new { x.WorkOrder, x.ControlNo }).ToList());

                    //Set DisplayMember to be workorder + control no
                    foreach (SearchInfo searchInfo in InterimJobCollection)
                    {
                        searchInfo.DisplayInfo = $"WO: {searchInfo.WorkOrder} CN: {searchInfo.ControlNo}";
                    }

                    if (InterimJobCollection.Count == 0)
                        SelectedJobInfo = null;
                    else
                        SelectedJobInfo = JobList.FirstOrDefault(x => x.WorkOrder.ToString() == InterimJobCollection.First().WorkOrder.ToString());
                }

                NotifyOfPropertyChange(() => DwgMatchText);
                NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
                NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
            }
        }

        //COnducts search by Spool No
        public string SpoolMatchText
        {
            get { return _spoolMatchText; }
            set
            {
                _spoolMatchText = value;

                InterimJobCollection.Clear();
                if(String.IsNullOrWhiteSpace(value) || value.Length < 4)
                {
                    SelectedJobInfo = null;
                }
                else
                {
                    if (ExcludeOldJobs)
                    {
                        InterimJobCollection.AddRange(SearchInfoList.Where(x => (x.SpoolNumber.ToLower().Contains(value.ToString())) ||
                                                 (x.SpoolNumber2.ToLower().Contains(value.ToString()))).Where(x=>!x.IsJobComplete).DistinctBy(x => new { x.WorkOrder, x.ControlNo }).ToList());
                    }
                    else
                    {
                        InterimJobCollection.AddRange(SearchInfoList.Where(x => (x.SpoolNumber.ToLower().Contains(value.ToString())) ||
                                                 (x.SpoolNumber2.ToLower().Contains(value.ToString()))).DistinctBy(x => new { x.WorkOrder, x.ControlNo }).ToList());
                    }

                    //Set DisplayMember to be workorder + control no
                    foreach (SearchInfo searchInfo in InterimJobCollection)
                    {
                        searchInfo.DisplayInfo = $"WO: {searchInfo.WorkOrder} CN: {searchInfo.ControlNo}";
                    }

                    if (InterimJobCollection.Count == 0)
                        SelectedJobInfo = null;
                    else
                        SelectedJobInfo = JobList.FirstOrDefault(x => x.WorkOrder.ToString() == InterimJobCollection.First().WorkOrder.ToString());
                }

                NotifyOfPropertyChange(() => SpoolMatchText);
                NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
                NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
            }
        }

        public bool ExcludeOldJobs
        {
            get { return _excludeOldJobs; }
            set
            {
                if (_excludeOldJobs == value)
                    return;

                _excludeOldJobs = value;

                NotifyOfPropertyChange(() => ExcludeOldJobs);
            }
        }

        public BindableCollection<string> ClientList
        {
            get { return _clientList; }
            set
            {
                _clientList = value;
                NotifyOfPropertyChange(() => ClientList);
            }
        }

        public string SelectedClient
        {
            get { return _selectedClient; }
            set
            {
                _selectedClient = value;
                NotifyOfPropertyChange(() => SelectedClient);
            }
        }


        public Report Report
        {
            get { return _report; }
            set { _report = value; }
        }
        #endregion

        /// <summary>
        /// Initialize list of all active jobs
        /// </summary>
        public ShellViewModel()
        {
            DataConnection = new DataConnection();
            //Initialize Active Job list
            JobList.AddRange(DataConnection.AssignJobPropertiesFromData(DataConnection.GetUniqueWorkOrderData()));

            //Initialize Client list from active job list
            ClientList.AddRange(JobList.Select(j => j.ClientName).ToList().Distinct());

            //Initialize searcheable list
            SearchInfoList.AddRange(DataConnection.GetSearcheableList());

            //Acquire active jobs list
            ActiveJobsList = new List<string>();
            ActiveJobsList = JobNav.GetActiveJobsList();
        }

        #region Command Clear PO text field

        public void ClearPoText(string poMatchText)
        {
            PoMatchText = "";
        }

        public bool CanClearPoText(string poMatchText)
        {
            return !String.IsNullOrWhiteSpace(poMatchText);
        }

        #endregion

        #region Command Add Job to Report list

        public void AddJobToReportList()
        {
            SelectedJobList.Add(SelectedJobInfo);
            NotifyOfPropertyChange(() => CanCreateReport);
        }

        public void AddAllJobsToReportList()
        {
            foreach(SearchInfo si in InterimJobCollection)
            {
                SelectedJobList.Add(JobList.First(x => x.WorkOrder == si.WorkOrder));
            }
            NotifyOfPropertyChange(() => CanCreateReport);
        }

        public void RemoveAllJobsFromReportList()
        {
            if (this.SelectedJobList.Count > 0)
                SelectedJobList.Clear();
        }



        public bool CanAddJobToReportList
        {
            get
            {
                return this.SelectedJobInfo != null;
            }
        }

        public bool CanAddAllJobsToReportList
        {
            get
            {
                return this.InterimJobCollection.Count > 0;
            }
        }

        #endregion

        #region Command Remove job from Report list

        public void RemoveJobFromReportList()
        {
            if (SelectedJobInfoFromReport != null)
            {
                SelectedJobList.Remove(SelectedJobInfoFromReport);
            }
                
        }

        public bool CanRemoveJobFromList
        {
            get
            {
                return this.SelectedJobInfoFromReport != null;
            }
        }

        public void RemoveJobFromInterimList()
        {
            if(SelectedSearchInfo != null)
            {
                InterimJobCollection.Remove(SelectedSearchInfo);
                SelectedSearchInfo = null;
                NotifyOfPropertyChange(() => CanRemoveJobFromInterimList);
            }
        }

        public bool CanRemoveJobFromInterimList
        {
            get
            {
                return this.SelectedSearchInfo != null;
            }
        }

        public void RemoveAllJobsFromInterimList()
        {
            InterimJobCollection.Clear();
            SelectedSearchInfo = null;
            NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
        }

        public bool CanRemoveAllJobsFromInterimList
        {
            get
            {
                return this.InterimJobCollection.Count > 0;
            }
        }

        #endregion

        #region Command Create report

        public async Task CreateReport()
        {
            SelectedJobSpoolList = new BindableCollection<Spool>();

            List<Task> tasks1 = new List<Task>();

            

            foreach(JobInfo job in SelectedJobList)
            {
                tasks1.Add(Task.Run(()=>SelectedJobSpoolList.AddRange(DataConnection.AssignSpoolPropertiesFromData(DataConnection.GetSpoolsData(job.WorkOrder)))));
            }

            await Task.WhenAll(tasks1);

            //Read CCO Loading
            var CCOLogList = await Task.Run(()=> ReportGen.Models.CCOImporting.CCOOutput.GetCCOLogList());

            foreach (Spool sp in SelectedJobSpoolList)
            {
                //Find CCO completion date for spools
                if(CCOLogList.Where(x => (x.WorkOrder == sp.WorkOrder && x.ControlNo.PadLeft(6,'0') == sp.ControlNo)) != null && CCOLogList.Where(x => (x.WorkOrder == sp.WorkOrder && x.ControlNo.PadLeft(6, '0') == sp.ControlNo)).Count() > 0)
                    sp.EstimatedCCOCompleteDate = CCOLogList.FirstOrDefault(x => (x.WorkOrder == sp.WorkOrder) && (x.ControlNo.PadLeft(6,'0') == sp.ControlNo)).DateCompleted.AddDays(2);

                if (sp.EstimatedCCOCompleteDate == null && sp.PtCcoProgress == 0.9m)
                    sp.EstimatedCCOCompleteDate = DateTime.UtcNow.AddDays(3);

            }
            Report = new Report();
            Report.CreateSummary(SelectedJobList.OrderBy(j=>j.WorkOrder).ToList());
            await Task.Run(() => Report.CreateDetailed(SelectedJobSpoolList.OrderBy(s=>s.WorkOrder).ThenBy(s=>s.ControlNo).ToList()));

            //Create Facing sheet
            await Task.Run(() => Report.CreateFacingSheet());

            MessageBox.Show("Report Generated");

        }

        public bool CanCreateReport
        {
            get
            {
                return SelectedJobList.ToList().Count() > 0;
            }
        }

        #endregion
        /*
        public void AddSelectionToInterimList()
        {
            InterimJobList.Clear();
            InterimJobList.Add(SelectedJobInfo);
        }
        */

        #region Multi Selection Commands

        public void SelectAllClientJob()
        {
            InterimJobCollection.Clear();

            foreach (SearchInfo searchInfo in SearchInfoList)
                searchInfo.DisplayInfo = searchInfo.WorkOrder;

            if(ExcludeOldJobs)
                InterimJobCollection.AddRange((from c in SearchInfoList.Where(j => j.ClientName == SelectedClient) join j in ActiveJobsList on c.WorkOrder equals j select c).DistinctBy(x => x.WorkOrder));
            else
                InterimJobCollection.AddRange(SearchInfoList.Where(j => j.ClientName == SelectedClient).DistinctBy(x=>x.WorkOrder));

            NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
            NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
        }

        public void SelectAllCCOJob()
        {
            InterimJobCollection.Clear();

            foreach (SearchInfo searchInfo in SearchInfoList)
                searchInfo.DisplayInfo = searchInfo.WorkOrder;

            if (ExcludeOldJobs)
            {
                InterimJobCollection.AddRange((from c in SearchInfoList.Where(j => j.boolCCO) join j in ActiveJobsList on c.WorkOrder equals j select c).DistinctBy(x=>x.WorkOrder));
            }

            else
                InterimJobCollection.AddRange(SearchInfoList.Where(j => j.boolCCO).DistinctBy(x => x.WorkOrder));

            NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
            NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
        }

        public void SelectAllFabJob()
        {
            InterimJobCollection.Clear();

            foreach (SearchInfo searchInfo in SearchInfoList)
                searchInfo.DisplayInfo = searchInfo.WorkOrder;

            if (ExcludeOldJobs)
                InterimJobCollection.AddRange((from c in SearchInfoList.Where(j => !j.boolCCO) join j in ActiveJobsList on c.WorkOrder equals j select c).DistinctBy(x => x.WorkOrder));
            else
                InterimJobCollection.AddRange(SearchInfoList.Where(j => !j.boolCCO).DistinctBy(x => x.WorkOrder));

            NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
            NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
        }

        public void SelectAllHMSJob()
        {
            InterimJobCollection.Clear();

            foreach (SearchInfo searchInfo in SearchInfoList)
                searchInfo.DisplayInfo = searchInfo.WorkOrder;

            if (ExcludeOldJobs)
                InterimJobCollection.AddRange((from c in SearchInfoList.Where(j => j.boolHMS) join j in ActiveJobsList on c.WorkOrder equals j select c).DistinctBy(x => x.WorkOrder));
            else
                InterimJobCollection.AddRange(SearchInfoList.Where(j => j.boolHMS).DistinctBy(x => x.WorkOrder));

            NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
            NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
        }


        public void SelectEveryJob()
        {
            InterimJobCollection.Clear();

            foreach (SearchInfo searchInfo in SearchInfoList)
                searchInfo.DisplayInfo = searchInfo.WorkOrder;

            if (ExcludeOldJobs)
                InterimJobCollection.AddRange((from c in SearchInfoList join j in ActiveJobsList on c.WorkOrder equals j select c).DistinctBy(x => x.WorkOrder));
            else
                InterimJobCollection.AddRange(SearchInfoList.DistinctBy(x => x.WorkOrder));

            NotifyOfPropertyChange(() => CanRemoveAllJobsFromInterimList);
            NotifyOfPropertyChange(() => CanAddAllJobsToReportList);
        }

        #endregion
    }
}
