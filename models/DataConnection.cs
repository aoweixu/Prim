using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Globalization;

namespace ReportGen.Models
{
    public class DataConnection
    {
        #region Private Members
        private string _sqlConnStr;
        private string _workOrder;
        #endregion

        #region Public Properties
        public string SqlConnStr
        {
            get { return _sqlConnStr; }
            set { _sqlConnStr = value; }
        }

        public string WorkOrder
        {
            get { return _workOrder; }
            set { _workOrder = value; }
        }

        #endregion

        //Constructor
        public DataConnection()
        {
            //Check from where the program is launched
            if (Environment.UserName.ToUpper() == "HYDROPC")
            {
                //Connection string for local replicate DB for testing outside work
                SqlConnStr = @"Server=DESKTOP-7CG2OJV\TEW_SQLEXPRESS; " +
                             @"Database=master; " +
                             @"Trusted_Connection=true";
            }
            else
            {
                //Connection String to Huston Server [Core]
                SqlConnStr = @"Data Source =TULVCORESQL01; " +
                             @"Initial Catalog =ProjectControls; " +
                             @"User id =fabshop; " +
                             @"Password =WqRm9BGw!ZT^[jsr;";
            }
        }

        public DataTable GetSpoolsData(string workOrder)
        {
            SqlConnection con = new SqlConnection(SqlConnStr);

            #region Select String

            string selectStr = @"select tsc.varClient, tsc.intWorkOrderNumber, tsc.varControlNumber, tsc.varClient,  tsc.varContractNumber, tsc.varSpoolNumber, tsc.varRefDrawing, tsc.varRefRev,
	                                    tsc.varPriority, taw.decCompletedFdi, taw.intWeldCount, taw.intFacingCount, taw.intCompletedFaces, tsc.decDiaInch,
										tsc.dteChkDate, tsc.varUser14, tsc.varStatus, tsc.dteInFabDate, tsc.dteClientRAS, tsc.dteDateAdded, tsc.dteHydroDate, tsc.varPipeSpec, 
										tsc.decWeight, tsc.varHTPressure, tsc.varSpoolID, tcco.decCCOPercentComplete, tcco.bolSpotRepair, tship.dteShipmentDate_Acorn,
	                                    tship.varShipmentNumber, tship.varShippingCompany, tship.bolPackingSlip, tfabWO.bolCCO
                                From
                                    (
	                                    select sc.intWorkOrderNumber, sc.varControlNumber, sc.varClient, sc.varSpoolNumber, sc.varRefDrawing, sc.varRefRev, sc.varContractNumber, sc.dteDateAdded,
		                                       sc.varPriority, sc.decDiaInch, sc.dteChkDate, sc.varStatus, sc.dteInFabDate, sc.dteHydroDate, sc.varUser02 as varSpoolID, sc.dteClientRAS, sc.varHTPressure,
                                               sc.varPipeSpec, sc.decWeight, sc.varUser14
	                                    from vwAcornSCN sc
                                    ) tsc
                                    left join
                                    (
	                                    select aw.intWorkOrderNumber, aw.varControlNumber, sum(case when aw.dteWeldDate is not null then aw.decDiaInch else 0 end) as decCompletedFdi,
											   count(aw.varWeldLabel) as intWeldCount, sum(case when aw.varWeldLabel like '%CCO[1-9]%' then 1 else 0 end) as intFacingCount,
											   sum(case when aw.varWeldLabel like '%CCO[1-9]%' and aw.dteWeldDate is not null then 1 else 0 end) as intCompletedFaces
	                                    from vwAcornWeld aw
	                                    group by aw.intWorkOrderNumber, aw.varControlNumber
                                    ) taw
                                    on (tsc.intWorkOrderNumber = taw.intWorkOrderNumber and tsc.varControlNumber = taw.varControlNumber)
                                    left join
                                    (
	                                    select cco.intWorkOrderNumber, cco.varCCOControlNumber, cco.decPercentComplete as decCCOPercentComplete, cco.bolSpotRepair
	                                    from vwCCOProgress cco
                                    ) tcco
                                    on (tsc.intWorkOrderNumber = tcco.intWorkOrderNumber and tsc.varControlNumber = right('00000' +isnull(tcco.varCCOControlNumber, ''),6))
                                    left join
                                    (
	                                    select ship.intWorkOrderNumber, ship.varControlNumber, ship.dteShipmentDate_Acorn, ship.varShipmentNumber, ship.varShippingCompany, bolPackingSlip
	                                    from vwShipments ship
                                    ) tship
                                    on (tsc.intWorkOrderNumber = tship.intWorkOrderNumber and tsc.varControlNumber = tship.varControlNumber)
									left join
									(
										select fabWO.intWorkOrderNumber, fabWO.bolCCO, fabWO.bolHMS
										from vwFabricationWorkOrders fabWO
									) tfabWO
									on (tsc.intWorkOrderNumber = tfabWO.intWorkOrderNumber)
                                where (not (tsc.varControlNumber like '%PF%') and  tsc.intWorkOrderNumber = " + workOrder + ")";

            #endregion

            SqlDataAdapter da = new SqlDataAdapter(selectStr, con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            con.Close();
            return dt;
        }

        public DataTable GetUniqueWorkOrderData()
        {
            SqlConnection con = new SqlConnection(SqlConnStr);

            #region Select String

            string selectStr = @"select t4.intWorkOrderNumber as trueWorkOrder, *
                                from
                                    (select 
                                        intWorkOrderNumber,
                                        bolHasAcorn,
                                        bolCCO,
                                        bolFabricationComplete,
                                        dteFabricationComplete,
                                        varClient,
                                        varContractNumber,
                                        dteClientRAS,
                                        varProjectDescription
                                    from vwFabricationWorkOrders) t4
                                left join
                                	(select 
                                        intWorkOrderNumber,  
 
                                        count(case when not(varControlNumber like '%PF%') then 1 else null end) as intItemCount,  
                                        sum(decDiaInch) as decTotalFdi,
                                        sum(case when dteInFabDate is not null then decDiaInch else 0 end) as decReleasedFdi,
										sum(case when dteHydroDate is not null then 1 else 0 end) as intSpoolsHydroComplete,
										sum(case when varUser15 is not null and varUser15 <> '' then 1 else 0 end) as intSpoolsShipped,
                                        sum(case when dteInFabDate is not null then 1 else 0 end) as intSpoolsInFab,
										max(varUser15) as dteLastShippingDate
                                    from vwAcornSCN
									group by intWorkOrderNumber) t1
                                on
                                    t4.intWorkOrderNumber = t1.intWorkOrderNumber
                                left join
                                	(select 
                                        intWorkOrderNumber, 
                                        sum(case when dteWeldDate is not null then decDiaInch else 0 end) as decCompletedFdi
                                    from vwAcornWeld group by intWorkOrderNumber) t2
                                on
                                	t4.intWorkOrderNumber = t2.intWorkOrderNumber
                                left join
                                    (select
                                        intWorkOrderNumber,
                                        avg(coalesce(decPercentComplete, cast(0 as decimal))) as decAverageCcoComplete
                                    from vwCCOProgress group by intWorkOrderNumber) t3
                                on
                                    t4.intWorkOrderNumber = t3.intWorkOrderNumber
								order by t4.intWorkOrderNumber";

            #endregion

            SqlDataAdapter da = new SqlDataAdapter(selectStr, con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            con.Close();

            return dt;
        }

        public List<SearchInfo> GetSearcheableList()
        {
            SqlConnection con = new SqlConnection(SqlConnStr);

            #region Select String

            /*
            string selectStr = @"select s.intWorkOrderNumber,
                                        s.varControlNumber,
                                        s.varContractNumber,
                                        s.varSpoolNumber,
                                        s.varUser02,
                                        s.varRefDrawing,
                                        s.varClient,
                                        f.bolFabricationComplete,
                                        f.bolHMS,
                                        f.bolCCO
                                from vwAcornSCN s
                                left join vwFabricationWorkOrders f
                                on s.intWorkOrderNumber = f.intWorkOrderNumber";
                                */

            string selectStr = @"select f.intWorkOrderNumber as realWOrkOrderNumber,
	                                    s.varControlNumber,
	                                    f.varContractNumber,
	                                    s.varSpoolNumber,
	                                    s.varUser02,
	                                    s.varRefDrawing,
	                                    f.varClient,
	                                    f.bolFabricationComplete,
	                                    f.bolCCO,
	                                    f.bolHMS,
	                                    f.bolHasAcorn	
                                from vwFabricationWorkOrders f
                                left join
                                vwAcornSCN s
                                on f.intWorkOrderNumber = s.intWorkOrderNumber
                                order by f.intWorkOrderNumber";
            #endregion

            SqlDataAdapter da = new SqlDataAdapter(selectStr, con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            con.Close();

            var searchInfoList = (from rw in dt.AsEnumerable()
                                  select new SearchInfo()
                                  {
                                      WorkOrder = GetData(rw["realWorkOrderNumber"]),
                                      ControlNo = GetData(rw["varControlNumber"]),
                                      PoNumber = GetData(rw["varContractNumber"]),
                                      SpoolNumber = GetData(rw["varSpoolNumber"]),
                                      SpoolNumber2 = GetData(rw["varUser02"]),
                                      DrawingNo = GetData(rw["varRefDrawing"]),
                                      ClientName = GetData(rw["varClient"]),
                                      IsJobComplete = Convert.ToBoolean(rw["bolFabricationComplete"]),
                                      boolHMS = Convert.ToBoolean(rw["bolHMS"]),
                                      boolCCO = Convert.ToBoolean(rw["bolCCO"])

                                  }).ToList();

            return searchInfoList;

        }

        public static List<Spool> AssignSpoolPropertiesFromData(DataTable dt)
        {
            var spoolList = (from rw in dt.AsEnumerable()
                             select new Spool()
                             {
                                 ClientName = GetData(rw["varClient"]),
                                 WorkOrder = GetData(rw["intWorkOrderNumber"]),
                                 ControlNo = GetData(rw["varControlNumber"]),
                                 PoNumber = GetData(rw["varContractNumber"]),
                                 SpoolNumber = GetData(rw["varSpoolNumber"]),
                                 SpoolNumber2 = GetData(rw["varSpoolID"]),
                                 DrawingNo = GetData(rw["varRefDrawing"]),
                                 DwgRev = GetData(rw["varRefRev"]),
                                 Priority = GetData(rw["varPriority"]),
                                 DwgCheckDate = GetDateData(rw["dteChkDate"]),
                                 Status = GetData(rw["varStatus"]),
                                 DateInFab = GetDateData(rw["dteInFabDate"]),
                                 FdiCount = rw["decDiaInch"] == DBNull.Value ? 70m : GetDecimalData(rw["decDiaInch"]),
                                 CompletedFdi = GetDecimalData(rw["decCompletedFdi"]),
                                 DateHydrotested = GetDateData(rw["dteHydroDate"]),
                                 DateNdeComp = GetDateFromVar(rw["varUser14"]),
                                 PtCcoProgress = GetDecimalData(rw["decCCOPercentComplete"]) > 0.9M && !GetBooleanData(rw["bolSpotRepair"]).GetValueOrDefault() ? 0.9M : GetDecimalData(rw["decCCOPercentComplete"]),
                                 IsSpotRepair = GetBooleanData(rw["bolSpotRepair"]),
                                 DateShipped = GetDateData(rw["dteShipmentDate_Acorn"]),
                                 ShipmentLoadNumber = GetData(rw["varShipmentNumber"]),
                                 ShippingCompany = GetData(rw["varShippingCompany"]),
                                 HasPackingSlip = GetBooleanData(rw["bolPackingSlip"]),
                                 RasDate = GetDateData(rw["dteClientRAS"]),
                                 DateAdded = GetDateData(rw["dteDateAdded"]),
                                 HydroPressure = GetData(rw["varHTPressure"]),
                                 PipeSpec = GetData(rw["varPipeSpec"]),
                                 Weight = GetDecimalData(rw["decWeight"]),
                                 NumFacesCompleted = rw["intCompletedFaces"] == DBNull.Value ? 0 : Convert.ToInt32(rw["intCompletedFaces"]),
                                 NumFacesTotal = rw["intFacingCount"] == DBNull.Value ? 0 : Convert.ToInt32(rw["intFacingCount"]),
                                 boolCCO = GetBooleanData(rw["bolCCO"]).HasValue ? GetBooleanData(rw["bolCCO"]).Value : false

                             }).ToList();

            return spoolList;
        }

        public List<JobInfo> AssignJobPropertiesFromData(DataTable dt)
        {

            //Convert.ToBoolean(rw["bolHasAcorn"]),
            var jobList = (from rw in dt.AsEnumerable()
                           select new JobInfo()
                           {
                               WorkOrder = GetData(rw["trueWorkOrder"]),
                               InAcorn = true,
                               ProjDesc = GetData(rw["varProjectDescription"]),
                               ClientName = GetData(rw["varClient"]),
                               //ControlNo not added as this is Job level
                               ItemCount = GetIntData(rw["intItemCount"]),
                               CompletedFdi = GetDecimalData(rw["decCompletedFdi"]),
                               ReleasedFdi = GetDecimalData(rw["decReleasedFdi"]),
                               FdiCount = GetDecimalData(rw["decTotalFdi"]),
                               RasDate = GetDateData(rw["dteClientRAS"]),
                               PoNumber = GetData(rw["varContractNumber"]),
                               NumSpoolsHydrotested = GetIntData(rw["intSpoolsHydroComplete"]),
                               NumSpoolsInFab = GetIntData(rw["intSpoolsInFab"]),
                               NumSpoolsShipped = GetIntData(rw["intSpoolsShipped"]),
                               PtAverageCCOProgress = GetDecimalData(rw["decAverageCcoComplete"]),
                               boolCCO = Convert.ToBoolean(rw["bolCCO"]),

                           }).ToList();

            return jobList;
        }

        #region DB Error checking

        private static string GetData(object o)
        {
            if (o == DBNull.Value)
            {
                return String.Empty;
            }
            return o.ToString();
        }

        private static int GetIntData(object o)
        {
            return o == DBNull.Value ? 0 : Convert.ToInt32(o);
        }

        private static DateTime GetDateData(object o)
        {
            return o == DBNull.Value ? DateTime.MinValue : Convert.ToDateTime(o);
        }

        private static DateTime GetDateFromVar(object o)
        {
            return o == DBNull.Value || o.ToString() == String.Empty ? DateTime.MinValue : DateTime.Parse(o.ToString()); 
        }

        private static decimal? GetDecimalData(object o)
        {
            return o == DBNull.Value ? Decimal.Zero : Convert.ToDecimal(o);
        }

        private static Boolean? GetBooleanData(object o)
        {
            return o == DBNull.Value ? (Boolean?)null : Convert.ToBoolean(o);
        }

        #endregion
    }
}
