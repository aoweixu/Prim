using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGen.Models.CCOImporting
{
    public enum Straight
    {
        WO = 1,
        SCN,
        Length,
        Size,
        Days,
        RasDate,
        Date
    }

    public enum Elbow
    {
        WO = 1,
        SCN,
        Size,
        Bend,
        Days,
        RasDate,
        Date
    }

    public enum Reducer
    {
        WO = 1,
        SCN,
        Length,
        Size,
        Days,
        RasDate,
        Date
    }
}
