using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ReportGen.Models;
using System.Reflection;
using System.ComponentModel;

namespace ReportGen.Attributes
{
    public class ReportColumnsWidthAttribute : Attribute
    {
        internal ReportColumnsWidthAttribute(int width)
        {
            this.ColWidth = width;
        }

        public int ColWidth { get; private set; }

    }

    public class ReportColumnsHiddenAttribute : Attribute
    {
        internal ReportColumnsHiddenAttribute(bool isHidden)
        {
            this.IsHidden = IsHidden;
        }

        public bool IsHidden { get; private set; }
    }

    public enum ReportColumns
    {
        ///Enum names must be the same as the property names of Spool class 
        //in order for data to be written to excel properly
        [Description("Client"), ReportColumnsWidth(20)]
        ClientName = 1,
        [Description("PO #"), ReportColumnsWidth(15)]
        PoNumber,
        [Description("Work order#"), ReportColumnsWidth(13)]
        WorkOrder,
        [Description("Control #"), ReportColumnsWidth(10)]
        ControlNo,
        [Description("PrimID"),ReportColumnsWidth(13)]
        ConcactenatedName,
        [Description("Drawing #"), ReportColumnsWidth(20)]
        DrawingNo,
        [Description("Rev"), ReportColumnsWidth(10)]
        DwgRev,
        [Description("Item #"), ReportColumnsWidth(20)]
        SpoolNumber,
        [Description("Alternative Item #"), ReportColumnsWidth(20)]
        SpoolNumber2,
        [Description("FDI"),ReportColumnsWidth(10)]
        FdiCount,
        [Description("Priority"), ReportColumnsWidth(10)]
        Priority,
        [Description("ITP Status"), ReportColumnsWidth(18)]
        ItpStatusStr,
        [Description("ITP Date"), ReportColumnsWidth(18)]
        ItpDate,
        [Description("CCO Progress"), ReportColumnsWidth(20)]
        PtCcoProgress,
        [Description("Est. CCO Comp Date"), ReportColumnsWidth(25)]
        EstimatedCCOCompleteDate,
        [Description("In Fab Date"), ReportColumnsWidth(20)]
        DateInFab,
        [Description("Welding Progress"), ReportColumnsWidth(20)]
        PtFabProgress,
        [Description("Faces Total"), ReportColumnsWidth(10)]
        NumFacesTotal,
        [Description("Faces Done"), ReportColumnsWidth(10)]
        NumFacesCompleted,
        [Description("Hydro Date"), ReportColumnsWidth(18)]
        DateHydrotested,
        [Description("QC Complete Date"), ReportColumnsWidth(18)]
        DateNdeComp,
        [Description("Ship Date"), ReportColumnsWidth(18)]
        DateShipped,
        [Description("Load #"), ReportColumnsWidth(18)]
        ShipmentLoadNumber,
        [Description("Date Added"), ReportColumnsWidth(18), ReportColumnsHidden(true)]
        DateAdded,
        [Description("Est. Ready to Ship Date"), ReportColumnsWidth(20)]
        EstimatedShipDate,
        [Description("Client RAS Date"), ReportColumnsWidth(18)]
        RasDate,
        [Description("Delta"), ReportColumnsWidth(10)]
        WorkDuration,
        [Description("Comment"), ReportColumnsWidth(40)]
        Comment,
    }

    public enum SummaryReportColumns
    {
        [Description("Client"), ReportColumnsWidth(18)] ClientName = 1,
        [Description("PO#"), ReportColumnsWidth(15)] PoNumber,
        [Description("Work order#"), ReportColumnsWidth(13)] WorkOrder,
        [Description("Scope Description"), ReportColumnsWidth(40)] ProjDesc,
        [Description("Number of items/spools"), ReportColumnsWidth(25)] ItemCount,
        [Description("RAS Date"), ReportColumnsWidth(15)] RasDate,
        [Description("ITP Status"), ReportColumnsWidth(20)] ItpStatusStr,
        [Description("OVerall CCO Progress"), ReportColumnsWidth(30)] PtAverageCCOProgress,
        [Description("Percent work released to fabrication"), ReportColumnsWidth(30)] PtReleasedFdi,
        [Description("Percent work completed"), ReportColumnsWidth(30)] PtCompletedFdi,
        [Description("Percent items shipped"), ReportColumnsWidth(30)] PtSpoolsShipped,
        [Description("RFI log"), ReportColumnsWidth(40)] RfiLog,
        [Description("Date Sent"), ReportColumnsWidth(20)] RfiLog2,
        [Description("Date Answered"), ReportColumnsWidth(20)] RfiLog3,

    }

    public static class ReportColumnsExtensions
    {
        public static int GetColWidth(this ReportColumns c)
        {
            var attr = c.GetAttribute<ReportColumnsWidthAttribute>();
            return attr.ColWidth;
        }

        public static int GetColWidth(this SummaryReportColumns c)
        {
            var attr = c.GetAttribute<ReportColumnsWidthAttribute>();
            return attr.ColWidth;
        }

        public static bool GetHiddenStatus(this ReportColumns c)
        {
            var attr = c.GetAttribute<ReportColumnsHiddenAttribute>();
            if (attr == null)
                return false;
            else
                return true;
        }

        public static bool GetHiddenStatus(this SummaryReportColumns c)
        {
            var attr = c.GetAttribute<ReportColumnsHiddenAttribute>();
            if (attr == null)
                return false;
            else
                return true;
        }

        public static string GetDescription<T>(this T enumerationValue)
    where T : struct
        {
            Type type = enumerationValue.GetType();
            if (!type.IsEnum)
            {
                throw new ArgumentException("EnumerationValue must be of Enum type", "enumerationValue");
            }

            //Tries to find a DescriptionAttribute for a potential friendly name
            //for the enum
            MemberInfo[] memberInfo = type.GetMember(enumerationValue.ToString());
            if (memberInfo != null && memberInfo.Length > 0)
            {
                object[] attrs = memberInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);

                if (attrs != null && attrs.Length > 0)
                {
                    //Pull out the description value
                    return ((DescriptionAttribute)attrs[0]).Description;
                }
            }
            //If we have no description attribute, just return the ToString of the enum
            return enumerationValue.ToString();
        }
    }

    public static class EnumExtensions
    {
        public static TAttribute GetAttribute<TAttribute>(this Enum value)
            where TAttribute : Attribute
        {
            var type = value.GetType();
            var name = Enum.GetName(type, value);
            return type.GetField(name) // I prefer to get attributes this way
                .GetCustomAttribute<TAttribute>();
        }
    }
}