using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ReportGen.Attributes;

namespace ReportGen.Models
{
    /// <summary>
    /// Contains properties for individual jobs/workorders
    /// </summary>
    public class JobInfo
    {
        public JobInfo()
        {

        }

        #region Public Properties

        public string WorkOrder
        {
            get { return _workOrder; }
            set
            {
                _workOrder = value;
                //Fetch ITP and RFI info once workorder is set
                Itp = new ItpInfo(value);
            }
        }

        public string JobType
        {
            get
            {
                if (this.boolCCO)
                    return "CCO";
                if (this.boolHMS)
                    return "HMS";

                return "General Fab (No CCO)";
            }
        }

        //Represents whether the job exists in Acorn
        public bool InAcorn
        {
            get { return _inAcorn; }
            set { _inAcorn = value; }
        }

        public string PoNumber
        {
            get { return _poNumber; }
            set { _poNumber = value; }
        }

        public string ClientName
        {
            get { return _clientName; }
            set { _clientName = value; }
        }

        public string ProjDesc
        {
            get { return _projDesc; }
            set { _projDesc = value; }
        }

        public Nullable<DateTime> RasDate
        {
            get { return _rasDate; }
            set { _rasDate = value; }
        }

        public int? ItemCount
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                else
                    return _itemCount;
            }
            set { _itemCount = value; }
        }

        public decimal? CompletedFdi
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return _completedFdi;
            }
            set { _completedFdi = value; }
        }

        public decimal? PtCompletedFdi //Readonly
        {
            get
            {
                if(!this.InAcorn)
                    return null;
                if (FdiCount != Decimal.Zero)
                    return CompletedFdi / FdiCount;
                else
                    return Decimal.Zero;

            }

        }

        public decimal? ReleasedFdi
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return _releasedFdi;
            }
            set { _releasedFdi = value; }
        }

        public decimal? PtReleasedFdi
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                try
                {
                    return ReleasedFdi / FdiCount;
                }
                catch (DivideByZeroException)
                {
                    return Decimal.Zero;
                }

            }
        }
        //Total FDI
        public decimal? FdiCount
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return _fdiCount;
            }
            set { _fdiCount = value; }
        }

        public decimal? PercentageFab //ReadOnly
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                if (FdiCount == Decimal.Zero)
                    return Decimal.Zero;
                else
                    return CompletedFdi.GetValueOrDefault() / FdiCount.GetValueOrDefault();
            }
        }


        public Nullable<DateTime> ItpDate //Readonly
        {
            get
            {
                return Itp.ItpDate;
            }
        }


        public string ItpStatusStr //Readonly
        {
            get
            {
                return Itp.ItpStatusStr;
            }
        }

        public Nullable<DateTime> DateAdded
        {
            get
            {
                return _dateAdded;
            }
            set { _dateAdded = value; }
        }

        public int? NumSpoolsInFab
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return _numSpoolsInFab;
            }
            set { _numSpoolsInFab = value; }
        }

        public decimal? PtSpoolsInFab //Readonly
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return NumSpoolsInFab / ItemCount; 
            }
        }

        public int? NumSpoolsHydrotested
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return _numSpoolsHydrotested;
            }
            set { _numSpoolsHydrotested = value; }
        }

        public decimal? PtSpoolsHydrotested //Readonly
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return NumSpoolsInFab / ItemCount;
            }
        }

        public int? NumSpoolsShipped
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return _numSpoolsShipped;
            }
            set { _numSpoolsShipped = value; }
        }

        public decimal? PtSpoolsShipped //Readonly
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                if (this.ItemCount == null || this.ItemCount == 0)
                    return Decimal.Zero;

                return (decimal)NumSpoolsShipped / ItemCount;
            }
        }

        public decimal? PtAverageCCOProgress
        {
            get
            {
                if (!this.InAcorn)
                    return null;
                return _ptAverageCCOProgress;
            }
            set { _ptAverageCCOProgress = value; }
        }
        #endregion

        #region Private Members

        private Nullable<DateTime> _rasDate;
        private string _projDesc;
        private bool _inAcorn;
        private string _workOrder;
        private string _poNumber;
        private string _clientName;
        private int? _itemCount;
        private decimal? _completedFdi;
        private decimal? _releasedFdi;
        private decimal? _fdiCount;
        private Nullable<DateTime> _dateAdded;
        private int? _numSpoolsInFab;
        private int? _numSpoolsHydrotested;
        private int? _numSpoolsShipped;
        private decimal? _ptAverageCCOProgress;
        public ItpInfo Itp { get; set; }
        public RfiInfo Rfi { get; set; }
        public bool boolCCO { get; set; }
        public bool boolHMS { get; set; }

        #endregion


    }


}
