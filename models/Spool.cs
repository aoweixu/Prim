using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FluentDate;
using FluentDateTime;

namespace ReportGen.Models
{
    public class Spool : JobInfo
    {

        public int DurationFab { get; set; }
        public int DurationHydro { get; set; }
        public int DurationShipping { get; set; }

        public Spool()
        {
            this.InAcorn = true;
        }

        private string _controlNo;

        public string ControlNo
        {
            get { return _controlNo; }
            set { _controlNo = value; }
        }

        public string ConcactenatedName
        {
            get
            {
                return $"{this.WorkOrder}-{this.ControlNo.TrimStart('0')}";
            }
        }


        private string _spoolNumber;

        public string SpoolNumber
        {
            get { return _spoolNumber; }
            set { _spoolNumber = value; }
        }

        private string _spoolNumber2;

        public string SpoolNumber2
        {
            get { return _spoolNumber2; }
            set { _spoolNumber2 = value; }
        }


        private string _drawingNo;

        public string DrawingNo
        {
            get { return _drawingNo; }
            set { _drawingNo = value; }
        }

        private string _dwgRev;

        public string DwgRev
        {
            get { return _dwgRev; }
            set { _dwgRev = value; }
        }


        private string _priority;

        public string Priority
        {
            get { return _priority; }
            set
            {
                _priority = value;
            }
        }

        private Nullable<DateTime> _dwgCheckDate;

        public Nullable<DateTime> DwgCheckDate
        {
            get
            {
                if (_dwgCheckDate == DateTime.MinValue)
                {
                    return null;
                }

                return _dwgCheckDate;
            }
            set { _dwgCheckDate = value; }
        }

        private string _status;

        public string Status
        {
            get { return _status; }
            set { _status = value; }
        }

        private double _ptMaterialReceived;

        public double PtMaterialReceived
        {
            get { return _ptMaterialReceived; }
            set { _ptMaterialReceived = value; }
        }

        private decimal? _ptCcoProgress;

        public decimal? PtCcoProgress
        {
            get
            {
                if (_ptCcoProgress == Decimal.Zero)
                {
                    return null;
                }

                return _ptCcoProgress;
            }
            set { _ptCcoProgress = value; }
        }

        private string _dateCcoComplete;

        public string DateCcoComplete
        {
            get { return _dateCcoComplete; }
            set { _dateCcoComplete = value; }
        }

        private Nullable<bool> _isSpotRepair;

        public Nullable<bool> IsSpotRepair
        {
            get { return _isSpotRepair; }
            set { _isSpotRepair = value; }
        }

        private Nullable<DateTime> _dateInFab;

        public Nullable<DateTime> DateInFab
        {
            get
            {
                if (_dateInFab == DateTime.MinValue)
                {
                    return null;
                }

                return _dateInFab;
            }
            set { _dateInFab = value; }
        }

        public Nullable<decimal> PtFabProgress
        {
            get
            {
                if (DateInFab == null)
                {
                    return null;
                }
                if (this.FdiCount == null || this.FdiCount == Decimal.Zero)
                    return Decimal.Zero;
                return (decimal?)this.CompletedFdi.GetValueOrDefault() / this.FdiCount.GetValueOrDefault();
            }
        }

        private int _numFacesTotal;

        public int NumFacesTotal
        {
            get { return _numFacesTotal; }
            set { _numFacesTotal = value; }
        }

        private int _numFacesCompleted;

        public int NumFacesCompleted
        {
            get { return _numFacesCompleted; }
            set { _numFacesCompleted = value; }
        }


        private Nullable<DateTime> _dateHydrotested;

        public Nullable<DateTime> DateHydrotested
        {
            get
            {
                if (_dateHydrotested == DateTime.MinValue)
                {
                    return null;
                }

                return _dateHydrotested;
            }
            set { _dateHydrotested = value; }
        }

        private Nullable<DateTime> _dateNdeComp;

        public Nullable<DateTime> DateNdeComp
        {
            get
            {
                if (_dateNdeComp == DateTime.MinValue)
                    return null;

                return _dateNdeComp;
            }
            set { _dateNdeComp = value; }
        }


        private Nullable<DateTime> _dateShipped;

        public Nullable<DateTime> DateShipped
        {
            get
            {
                if (_dateShipped == DateTime.MinValue)
                {
                    return null;
                }

                return _dateShipped;
            }
            set { _dateShipped = value; }
        }

        private string _shipmentLoadNumber;

        public string ShipmentLoadNumber
        {
            get { return _shipmentLoadNumber; }
            set { _shipmentLoadNumber = value; }
        }

        private string _shippingCompany;

        public string ShippingCompany
        {
            get { return _shippingCompany; }
            set { _shippingCompany = value; }
        }

        private Nullable<bool> _hasPackingSlip;

        public Nullable<bool> HasPackingSlip
        {
            get { return _hasPackingSlip; }
            set { _hasPackingSlip = value; }
        }

        private DateTime? _estimatedShipDate;

        private int _workDuration;

        public int WorkDuration
        {
            get
            {
                TimeSpan? durationSpan = RasDate - DateTime.Now;
                return durationSpan.GetValueOrDefault().Days;
            }
            set { _workDuration = value; }
        }

        private string _hydroPressure;

        public string HydroPressure
        {
            get { return _hydroPressure; }
            set { _hydroPressure = value; }
        }

        private string _pipeSpec;

        public string PipeSpec
        {
            get { return _pipeSpec; }
            set { _pipeSpec = value; }
        }

        private decimal? _weight;

        public decimal? Weight
        {
            get { return _weight; }
            set { _weight = value; }
        }

        public string Comment
        {
            get
            {
                return String.Empty;
            }
        }


        #region Shipping Date Estimation

        public DateTime? EstimatedShipDate
        {
            get
            {
                DateTime shipDay;

                if (this.DateHydrotested!=null)
                {
                    shipDay = ((DateTime)DateHydrotested).AddBusinessDays(2);

                }
                else if (this.DateInFab != null)
                {
                    int TotalFabDays;
                    if (this.FdiCount < 21.0m)
                    {
                        TotalFabDays = 7;
                    }
                    else if (this.FdiCount > 70.0m)
                    {
                        TotalFabDays = 13;
                    }
                    else
                    {
                        TotalFabDays = 10;
                    }

                    int daysToFabComplete = (int)((1 - this.PtFabProgress) * TotalFabDays);

                    shipDay = ((DateTime)DateInFab).AddBusinessDays(daysToFabComplete);
                }
                else if(this.PtCcoProgress > 0.98m)
                {
                    shipDay = DateTime.Now.AddBusinessDays(13);
                }
                else if (this.EstimatedCCOCompleteDate != null)
                {
                    shipDay = ((DateTime)EstimatedCCOCompleteDate).AddBusinessDays(13);
                }
                else
                {
                    return ((DateTime)this.RasDate).AddBusinessDays(-1);
                }

                DateTime shipDayFriday;
                shipDayFriday = shipDay.AddDays(1);
                while (shipDayFriday.DayOfWeek != DayOfWeek.Friday)
                    shipDayFriday = shipDayFriday.AddDays(1);

                return shipDayFriday;
            }

            set
            {
                _estimatedShipDate = value;

            }
        }

        private DateTime? _estimatedFabDate;

        public DateTime? EstimatedFabDate
        {
            get { return _estimatedFabDate; }
            set { _estimatedFabDate = value; }
        }

        private DateTime? _estimatedHydroDate;

        public DateTime? EstimatedHydroDate
        {
            get { return _estimatedHydroDate; }
            set { _estimatedHydroDate = value; }
        }

        private DateTime? _estimatedCCOCompleteDate;

        public DateTime? EstimatedCCOCompleteDate
        {
            get { return _estimatedCCOCompleteDate; }
            set { _estimatedCCOCompleteDate = value; }
        }

        public DateTime? CalculateEstimatedShipDate()
        {
            return null;
        }

        #endregion

    }

}
