using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Data;

namespace ReportGen.Models.CCOImporting
{
    class CCOOutput
    {
        private static string CCODocPath = @"I:\Fabrication\X. Planning Folder\Planning Meeting\CCO Shop Loading.xlsx";
        //private static string CCODocPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"Test CCO Shop Loading-1.xlsx");
        private static int numStraightMachines = 4;
        private static int numElbowMachines = 5;
        private static int numReducerMachines = 2;

        public static List<CCOLogItem> GetCCOLogList()
        {
            FileInfo CCODoc = new FileInfo(CCODocPath);

            //Initiate CCO log item collection
            List<CCOLogItem> CCOLogList = new List<CCOLogItem>();

            //Initialize error log
            List<ErrorLog> Errors = new List<ErrorLog>();

            //Open CCO loading excel document for reading
            using (var p = new ExcelPackage(CCODoc))
            {
                var wsStraights = p.Workbook.Worksheets[1];
                var wsElbows = p.Workbook.Worksheets[2];
                var wsReducers = p.Workbook.Worksheets[3];
                var wsSpotRepair = p.Workbook.Worksheets[5];

                //Initialize machines
                List<StraightMachine> straightMachines = new List<StraightMachine>();
                List<ElbowMachine> elbowMachines = new List<ElbowMachine>();
                List<ReducerMachine> reducerMachines = new List<ReducerMachine>();

                for(int i = 1; i <= numStraightMachines; i++)
                {
                    straightMachines.Add(new StraightMachine(i));
                }

                for(int j = 1; j <= numElbowMachines; j++)
                {
                    elbowMachines.Add(new ElbowMachine(j));
                }

                for(int k = 1; k <= numReducerMachines; k++)
                {
                    reducerMachines.Add(new ReducerMachine(k));
                }

                //Populate output datatable
                
                foreach(StraightMachine sm in straightMachines)
                {
                    for(int inputRow = 0; inputRow <= sm.TableHeight; inputRow++)
                    {

                        string scns = null;
                        try
                        {
                            scns = wsStraights.Cells[sm.TableStartRow + inputRow, (int)Straight.SCN + sm.TableColOffset * (sm.IntMachineID - 1)].Value.ToString();
                        }
                        catch (NullReferenceException)
                        {
                            continue;
                        }

                        if (scns.Contains("?"))
                            continue;

                        //Try parse & split SCN text
                        List<string> scnList1 = scns.Split(' ').ToList();
                        foreach (string s1 in scnList1)
                        {
                            if (!s1.Contains("--"))
                            {
                                CCOLogItem ccoItem = new CCOLogItem
                                {
                                    WorkOrder = wsStraights.Cells[sm.TableStartRow + inputRow, (int)Straight.WO + sm.TableColOffset * (sm.IntMachineID - 1)].Value.ToString().Trim(),
                                    ControlNo = s1.Trim(),
                                    Duration = Convert.ToInt32(wsStraights.Cells[sm.TableStartRow + inputRow, (int)Straight.Days + sm.TableColOffset * (sm.IntMachineID - 1)].Value),
                                    DateCompleted = DateTime.FromOADate(Convert.ToDouble(wsStraights.Cells[sm.TableStartRow + inputRow, (int)Straight.Date + sm.TableColOffset * (sm.IntMachineID - 1)].Value)),
                                    Machine = sm.MachineID
                                };

                                CCOLogList.Add(ccoItem);
                            }
                            else
                            {
                                int startSCN = 0;
                                int endSCN = 0;

                                try
                                {
                                    startSCN = Convert.ToInt32(s1.Trim().Substring(0, s1.IndexOf('-')));
                                    endSCN = Convert.ToInt32(s1.Trim().Substring(s1.IndexOf('-') + 2));
                                }
                                catch (FormatException)
                                {
                                    continue;
                                }

                                for (int k = startSCN; k <= endSCN; k++)
                                {
                                    CCOLogItem ccoItem = new CCOLogItem
                                    {
                                        WorkOrder = wsStraights.Cells[sm.TableStartRow + inputRow, (int)Straight.WO + sm.TableColOffset * (sm.IntMachineID - 1)].Value.ToString().Trim(),
                                        ControlNo = k.ToString(),
                                        Duration = Convert.ToInt32(wsStraights.Cells[sm.TableStartRow + inputRow, (int)Straight.Days + sm.TableColOffset * (sm.IntMachineID - 1)].Value),
                                        DateCompleted = DateTime.FromOADate(Convert.ToDouble(wsStraights.Cells[sm.TableStartRow + inputRow, (int)Straight.Date + sm.TableColOffset * (sm.IntMachineID - 1)].Value)),
                                        Machine = sm.MachineID
                                    };

                                    CCOLogList.Add(ccoItem);
                                }

                            }
                        }

                    }
 

                }

                foreach(ElbowMachine em in elbowMachines)
                {
                    for(int inputRow = 0; inputRow <= em.TableHeight; inputRow++)
                    {
                        CCOLogItem ccoItem = new CCOLogItem();
                        try
                        {
                            ccoItem.WorkOrder = wsElbows.Cells[em.TableStartRow + inputRow, (int)Straight.WO + em.TableColOffset * (em.IntMachineID - 1)].Value.ToString();
                        }
                        catch (NullReferenceException) { continue; }

                        if (ccoItem.WorkOrder.ToString().Contains("?"))
                            continue;

                        ccoItem.ControlNo = wsElbows.Cells[em.TableStartRow + inputRow, (int)Straight.SCN + em.TableColOffset * (em.IntMachineID - 1)].Value.ToString();
                        ccoItem.Duration = Convert.ToInt32(wsElbows.Cells[em.TableStartRow + inputRow, (int)Straight.Days + em.TableColOffset * (em.IntMachineID - 1)].Value);
                        ccoItem.DateCompleted = DateTime.FromOADate(Convert.ToDouble(wsElbows.Cells[em.TableStartRow + inputRow, (int)Straight.Date + em.TableColOffset * (em.IntMachineID - 1)].Value));
                        ccoItem.Machine = em.MachineID;

                        CCOLogList.Add(ccoItem);
                    }
                }

                foreach(ReducerMachine rm in reducerMachines)
                {
                    for(int inputRow = 0; inputRow<=rm.TableHeight; inputRow++)
                    {
                        CCOLogItem ccoItem = new CCOLogItem();
                        try
                        {
                            ccoItem.WorkOrder = wsReducers.Cells[rm.TableStartRow + inputRow, (int)Straight.WO + rm.TableColOffset * (rm.IntMachineID - 1)].Value.ToString();
                        }
                        catch (NullReferenceException) { continue; }

                        if (ccoItem.WorkOrder.ToString().Contains("?"))
                            continue;
                        if (wsReducers.Cells[rm.TableStartRow + inputRow, (int)Straight.SCN + rm.TableColOffset * (rm.IntMachineID - 1)].Value != null)
                            ccoItem.ControlNo = wsReducers.Cells[rm.TableStartRow + inputRow, (int)Straight.SCN + rm.TableColOffset * (rm.IntMachineID - 1)].Value.ToString();
                        else
                        {
                            Errors.Add(new ErrorLog
                            {
                                Type = "No Value",
                                MachineType = "Reducer",
                                Row = rm.TableStartRow + inputRow,
                                Column = (int)Straight.SCN + rm.TableColOffset * (rm.IntMachineID - 1)

                            });

                            continue;
                        }

                        ccoItem.Duration = Convert.ToInt32(wsReducers.Cells[rm.TableStartRow + inputRow, (int)Straight.Days + rm.TableColOffset * (rm.IntMachineID - 1)].Value);
                        ccoItem.DateCompleted = DateTime.FromOADate(Convert.ToDouble(wsReducers.Cells[rm.TableStartRow + inputRow, (int)Straight.Date + rm.TableColOffset * (rm.IntMachineID - 1)].Value));
                        ccoItem.Machine = rm.MachineID;

                        CCOLogList.Add(ccoItem);
                    }
                }

            }

            return CCOLogList;

            /*
            Console.WriteLine($"{outputDataTable.Rows.Count} Lines parsed");
            Console.WriteLine($"{Errors.Count} Errors encountered:");
            foreach(ErrorLog error in Errors)
            {
                Console.WriteLine($"Machine {error.MachineType} Line {error.Row} Column {error.Column} Error type {error.Type}");
            }
            Console.ReadLine();
            */
        }

        public static void PrintRow(DataRow dr)
        {
            //For logging
            //Console.WriteLine($"WO:{dr["WorkOrder"]} CN:{dr["ControlNo"]} Duration: {dr["Duration"]} Completion Date: {dr["DateComplete"]} Machine ID: {dr["MachineID"]}");
        }

        public static void PrintError(Exception ex)
        {
            //Console.WriteLine(ex);
        }
    }
}
