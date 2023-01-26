using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;
using ReportGen.Attributes;
using MoreLinq;

namespace ReportGen.Models
{
    /// <summary>
    /// Write detailed or summary report from job/spool data
    /// </summary>
    public class Report
    {
        #region Static variables
        static readonly System.Drawing.Color colorGenericGreen = System.Drawing.Color.FromArgb(102, 204, 102);
        static readonly System.Drawing.Color colorGenericRed = System.Drawing.Color.Red;
        static readonly System.Drawing.Color colorGenericOrange = System.Drawing.Color.Orange;
        static readonly System.Drawing.Color colorHeader = System.Drawing.Color.FromArgb(46, 46, 31);
        static readonly System.Drawing.Color colorPrimorisOne = System.Drawing.Color.FromArgb(214, 214, 194);
        static readonly System.Drawing.Color colorPrimorisTwo = System.Drawing.Color.FromArgb(184, 184, 148);

        #endregion

        public ExcelPackage p { get; set; }
        public System.IO.FileInfo ExcelReportDoc { get; set; }
        private bool ClientHasPicture { get; set; }
        private int DataRow { get; set; }


        public Report()
        {
            ExcelReportDoc = new System.IO.FileInfo(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "testreport.xlsx"));

            if (ExcelReportDoc.Exists)
                ExcelReportDoc.Delete();

            
        }

        public void CreateFacingSheet()
        {
            using (var p = new ExcelPackage(ExcelReportDoc))
            {
                var ws2 = p.Workbook.Worksheets[2];
                var ws1 = p.Workbook.Worksheets.Add("shtFacing", ws2);

                ws1.Column((int)ReportColumns.ClientName).Hidden = true;
                ws1.Column((int)ReportColumns.PoNumber).Hidden = true;
                ws1.Column((int)ReportColumns.ConcactenatedName).Hidden = true;
                ws1.Column((int)ReportColumns.DrawingNo).Hidden = true;
                ws1.Column((int)ReportColumns.DwgRev).Hidden = true;
                ws1.Column((int)ReportColumns.SpoolNumber).Hidden = true;
                ws1.Column((int)ReportColumns.SpoolNumber2).Hidden = true;
                ws1.Column((int)ReportColumns.FdiCount).Hidden = false;
                ws1.Column((int)ReportColumns.ItpDate).Hidden = true;
                ws1.Column((int)ReportColumns.ItpStatusStr).Hidden = true;
                ws1.Column((int)ReportColumns.EstimatedCCOCompleteDate).Hidden = true;
                ws1.Column((int)ReportColumns.PtCcoProgress).Hidden = true;
                ws1.Column((int)ReportColumns.DateShipped).Hidden = false;
                ws1.Column((int)ReportColumns.DateNdeComp).Hidden = true;
                ws1.Column((int)ReportColumns.ShipmentLoadNumber).Hidden = true;
                ws1.Column((int)ReportColumns.EstimatedShipDate).Hidden = true;
                ws1.Column((int)ReportColumns.RasDate).Hidden = true;
                ws1.Column((int)ReportColumns.WorkDuration).Hidden = true;
                ws1.Column((int)ReportColumns.Comment).Hidden = false;

                //Insert Date
                ws1.InsertRow(1, 1);
                ws1.Cells[1, (int)ReportColumns.ClientName, 1, (int)ReportColumns.Comment].Merge = true;
                ws1.Cells[1, 1].Value = $"Report Date: {DateTime.Now.ToString("MM/dd/yyyy")}";
                ws1.Row(1).Height = 20;
                ws1.Cells[1, 1].Style.Font.Size = 16;
                ws1.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                ws1.Cells[2, 1, ws1.Dimension.End.Row, ws1.Dimension.End.Column].AutoFilter = true;

                //Printer settings
                ws1.PrinterSettings.RepeatRows = new ExcelAddress("$1:$2");
                ws1.PrinterSettings.PaperSize = ePaperSize.Standard11_17;
                ws1.PrinterSettings.BlackAndWhite = false;
                ws1.PrinterSettings.VerticalCentered = true;
                ws1.PrinterSettings.FitToWidth = 1;
                ws1.PrinterSettings.FitToHeight = 0;
                ws1.PrinterSettings.Orientation = eOrientation.Landscape;

                p.Save();
            }
        }

        /// <summary>
        /// Create Summary Report
        /// </summary>
        public void CreateSummary(List<JobInfo> jobList)
        {

            DataRow = 2;

            //Determine Client
            string client = jobList.First().ClientName.ToUpper();
            System.Drawing.Bitmap HeaderBitmap = null;
            System.Drawing.Bitmap PrimorisBitmap = Properties.Resources.PSC_Willbros_Canada;

            if (client.Contains("SUNCOR"))
            {
                HeaderBitmap = Properties.Resources._284px_Suncor_Energy_logo_svg;
                DataRow = 3;
            }
            else if (client.Contains("CNRL"))
            {
                HeaderBitmap = Properties.Resources._1200px_Canadian_Natural_Logo_svg;
                DataRow = 3;
            }
            else if (client.Contains("IMPERIAL"))
            {
                HeaderBitmap = Properties.Resources.Imperial_Oil2_svg;
                DataRow = 3;
            }

            //Create new Excel sheet for writting
            using (var p = new ExcelPackage(ExcelReportDoc))
            {
                //Create Summary worksheet
                var ws = p.Workbook.Worksheets.Add("shtSummary");

                int i = DataRow;
                int i2 = 0;

                int pscLogoLocation = 0;
                int logoHeight = 50;

                //Add Job data
                foreach (JobInfo job in jobList)
                {
                    

                    //Add RFI info
                    var rfiList = JobNav.GetRfiInfos(job.WorkOrder, false);
                    int k = 0;
                    if(rfiList != null)
                    {
                        foreach (RfiInfo rfi in rfiList)
                        {
                            ws.Cells[i + k, (int)SummaryReportColumns.RfiLog].Value = rfi.Name;
                            ws.Cells[i + k, (int)SummaryReportColumns.RfiLog2].Value = String.Format("{0: M/d/yyyy}", rfi.SentDate);
                            ws.Cells[i + k, (int)SummaryReportColumns.RfiLog3].Value = rfi.IsAnswered ? String.Format("{0: M/d/yyyy}", rfi.AnswerDate) : "Not answered";
                            ws.Cells[i + k, (int)SummaryReportColumns.RfiLog, i + k, (int)SummaryReportColumns.RfiLog3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            k++;
                        }
                        //Row height to be uniform
                        if (k > 3) { } else
                        {
                            double singleRowHeight = (double)60 / (double)k;
                            for(int subK = 0; subK <=k; subK++)
                            {
                                ws.Row(i + subK).Height = singleRowHeight;
                            }
                        }

                        k--;

                    }
                    else
                    {
                        ws.Cells[i, (int)SummaryReportColumns.RfiLog].Value = "No RFI's";
                        ws.Row(i).Height = 60;
                        //Merge all RFI cols
                        ws.Cells[i, (int)SummaryReportColumns.RfiLog, i, (int)SummaryReportColumns.RfiLog3].Merge = true;
                        ws.Cells[i, (int)SummaryReportColumns.RfiLog].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        ws.Cells[i, (int)SummaryReportColumns.RfiLog].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    }

                    ExcelRange rowAddress = ws.Cells[i+k, 1, i+k, Enum.GetNames(typeof(SummaryReportColumns)).Length];

                    int j = 0;
                    //Add Job info
                    foreach (SummaryReportColumns src in (SummaryReportColumns[])Enum.GetValues(typeof(SummaryReportColumns)))
                    {
                        if (job.GetType().GetProperty(src.ToString()) == null)
                            continue;

                        ws.Cells[i, (int)src].Value = job.GetType().GetProperty(src.ToString()).GetValue(job, null);

                        if (src.ToString().ToUpper().Contains("DATE"))
                        {
                            ws.Cells[i, (int)src].Style.Numberformat.Format = "MMM-dd-yyyy";
                        }

                        //Merge & format
                        var mergedCellAddress = ws.Cells[i, (int)src, i + k, (int)src];
                        mergedCellAddress.Merge = true;
                        ws.Cells[i, (int)src].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        ws.Cells[i, (int)src].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        //Format non-CCO jobs
                        if (src == SummaryReportColumns.PtAverageCCOProgress && job.boolCCO == false)
                        {
                            ws.Cells[i, (int)src].Style.Border.DiagonalDown = true;
                            ws.Cells[i, (int)src].Style.Border.DiagonalUp = true;
                            ws.Cells[i, (int)src].Style.Border.Diagonal.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                            ws.Cells[i, (int)src].Value = String.Empty;
                        }

                        j++;
                    }

                    //Format - border after job
                    rowAddress.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; //bottom border
                    ws.Cells[i, Enum.GetNames(typeof(SummaryReportColumns)).Length, i+k, Enum.GetNames(typeof(SummaryReportColumns)).Length].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin; //right broder

                    //Incrementers
                    i = i + k;
                    i++;
                    i2 = i2 + i;
                }

                //Format - Final

                //Format - Worksheet level
                ws.Row(DataRow-1).Height = 25;
                var columnHeadersStyle = ws.Cells[DataRow-1, 1, DataRow-1, Enum.GetNames(typeof(SummaryReportColumns)).Length].Style;

                columnHeadersStyle.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                columnHeadersStyle.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                columnHeadersStyle.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                columnHeadersStyle.Fill.BackgroundColor.SetColor(colorHeader);
                columnHeadersStyle.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                columnHeadersStyle.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                columnHeadersStyle.Font.Color.SetColor(System.Drawing.Color.FromArgb(255, 255, 255));


                //Add headers
                foreach (SummaryReportColumns src in (SummaryReportColumns[])Enum.GetValues(typeof(SummaryReportColumns)))
                {
                    ExcelRange columnAddress = ws.Cells[DataRow - 1, (int)src, i-1, (int)src];

                    //Header - Name
                    ws.Cells[DataRow - 1, (int)src].Value = src.GetDescription();

                    //Assign Column width
                    int colWidth = src.GetColWidth();
                    ws.Column((int)src).Width = colWidth;

                    //Inner border
                    columnAddress.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    //Format decimal progress data
                    if (src.ToString().StartsWith("Pt"))
                    {
                        columnAddress.Style.Numberformat.Format = "#0.00%";

                        var bar = ws.ConditionalFormatting.AddDatabar(columnAddress, colorGenericGreen);
                        bar.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        bar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                        bar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                        bar.HighValue.Value = 1;
                        bar.LowValue.Value = 0;

                    }

                    #region Conditional Formatting
                    //Conditional Formatting - ITP 
                    if (src.ToString().ToUpper().Contains("ITP"))
                    {
                        //Conditional Formatting Accept
                        string cellAddress = ws.Cells[DataRow, (int)src].Address;
                        var expandedItpAddress = ws.Cells[DataRow, (int)src, i-1, (int)src];
                        string statementAccept = $"AND(ISNUMBER(SEARCH(\"APPR\",${cellAddress})),NOT(ISNUMBER(SEARCH(\"NOT\",${cellAddress}))))";
                        string statementNotAccept = $"AND(NOT({statementAccept}),NOT(ISNUMBER(SEARCH(\"ITP\",${cellAddress}))))";

                        //Accept
                        var itpConditionAccept = ws.ConditionalFormatting.AddExpression(expandedItpAddress);
                        itpConditionAccept.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        itpConditionAccept.Style.Fill.BackgroundColor.Color = colorGenericGreen;
                        itpConditionAccept.Formula = statementAccept;

                        //Not Accepted
                        var itpConditionNotAccept = ws.ConditionalFormatting.AddExpression(expandedItpAddress);
                        itpConditionNotAccept.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        itpConditionNotAccept.Style.Fill.BackgroundColor.Color = colorGenericOrange;
                        itpConditionNotAccept.Formula = statementNotAccept;

                    }
                    //Conditional Formatting - RFI
                    if(src == SummaryReportColumns.RfiLog3)
                    {
                        string cellAddress = ws.Cells[DataRow, (int)src].Address;
                        var expandedRfiAddress = ws.Cells[DataRow, (int)src-2, i - 1, (int)src];
                        string statementResponded = $"NOT(ISERR(DATEVALUE(${cellAddress})))";
                        string statementNotResponded = $"${cellAddress}=\"Not answered\"";

                        //Responded
                        var rfiResponded = ws.ConditionalFormatting.AddExpression(expandedRfiAddress);
                        rfiResponded.Style.Font.Color.Color = System.Drawing.Color.DarkSeaGreen;
                        rfiResponded.Formula = statementResponded;

                        //Not Responded
                        var rfiNotResponded = ws.ConditionalFormatting.AddExpression(expandedRfiAddress);
                        rfiNotResponded.Style.Font.Color.Color = colorGenericRed;
                        rfiNotResponded.Formula = statementNotResponded;
                    }
                    #endregion

                }

                //Create Client Logo Picture;
                if (HeaderBitmap != null)
                {
                    ExcelPicture clientLogo = ws.Drawings.AddPicture("clientLogo", HeaderBitmap);
                    double ratio = (double)HeaderBitmap.Width / (double)HeaderBitmap.Height;
                    clientLogo.SetPosition(0, 10, 0, 10);
                    clientLogo.SetSize((int)(logoHeight * ratio), logoHeight);
                    ws.Row(1).Height = logoHeight;
                    
                    //Bump PSC logo over
                    pscLogoLocation = 1;
                }

                //Create PSC logo
                ExcelPicture PSCLogo = ws.Drawings.AddPicture("PSCLogo", PrimorisBitmap);
                double ratio2 = (double)PrimorisBitmap.Width / (double)PrimorisBitmap.Height;
                PSCLogo.SetPosition(0,10,pscLogoLocation,10);
                PSCLogo.SetSize((int)(logoHeight * ratio2), logoHeight);

                SetDataBarFormat(ws);
                p.Save();
            }
        }

        /// <summary>
        /// Create Detailed Report
        /// </summary>
        /// <param name="spoolList"></param>
        public void CreateDetailed(List<Spool> spoolList)
        {
            int numWorkOrders = spoolList.DistinctBy(s => s.WorkOrder).Count();
            int woCounter = 0;

            using (var p = new ExcelPackage(ExcelReportDoc))
            {
                //Create schedule worksheet
                var ws = p.Workbook.Worksheets.Add("shtDetailedSchedule");
                
                //Assign Column Headers and Format Columns
                foreach (ReportColumns reportColumns in (ReportColumns[])Enum.GetValues(typeof(ReportColumns)))
                {
                    ExcelRange columnAddress = ws.Cells[1, (int)reportColumns, spoolList.Count() + 1, (int)reportColumns];
                    
                    //Header - Name
                    ws.Cells[1, (int)reportColumns].Value = reportColumns.GetDescription();

                    //Assign Column width
                    int colWidth = reportColumns.GetColWidth();
                    ws.Column((int)reportColumns).Width = colWidth;

                    //Header - Format
                    ws.Row(1).Height = 20;

                    var columnHeadersStyle = ws.Cells[1, 1, 1, Enum.GetNames(typeof(ReportColumns)).Length].Style;

                    columnHeadersStyle.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                    columnHeadersStyle.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    columnHeadersStyle.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    columnHeadersStyle.Fill.BackgroundColor.SetColor(colorHeader);
                    columnHeadersStyle.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    columnHeadersStyle.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    columnHeadersStyle.Font.Color.SetColor(System.Drawing.Color.FromArgb(255, 255, 255));

                    //Format Facing Data
                    if(reportColumns.ToString() == "NumFacesCompleted")
                    {
                        string cellAddress = ws.Cells[1, (int)reportColumns].Address;
                        string totalFacesCellAddress = ws.Cells[1, (int)ReportColumns.NumFacesTotal].Address;
                        var expandedAddress = ws.Cells[1, (int)ReportColumns.NumFacesTotal, spoolList.Count() + 1, (int)reportColumns];
                        string statement = $"AND(ISNUMBER(${cellAddress}), ${cellAddress}>0, ${cellAddress}=${totalFacesCellAddress})";

                        var faceCondition = ws.ConditionalFormatting.AddExpression(expandedAddress);
                        faceCondition.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        faceCondition.Style.Fill.BackgroundColor.Color = colorGenericGreen;
                        faceCondition.Formula = statement;
                    }

                    //Format Hydro Data (for NON-req)
                    if(reportColumns.ToString() == "DateHydrotested")
                    {
                        string cellAddress = ws.Cells[1, (int)reportColumns].Address;
                        string totalFacesCellAddress = ws.Cells[1, (int)ReportColumns.NumFacesTotal].Address;
                        string completedFacesCellAddress = ws.Cells[1, (int)ReportColumns.NumFacesCompleted].Address;
                        string statement = $"AND(${cellAddress}=\"NOT REQ\", ISNUMBER(${completedFacesCellAddress}), ${completedFacesCellAddress}>0, ${completedFacesCellAddress}=${totalFacesCellAddress})";
                        var expandedAddress = ws.Cells[1, (int)reportColumns, spoolList.Count() + 1, (int)reportColumns];

                        var hydroCondition = ws.ConditionalFormatting.AddExpression(expandedAddress);
                        hydroCondition.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        hydroCondition.Style.Fill.BackgroundColor.Color = colorGenericGreen;
                        hydroCondition.Formula = statement;
                    }

                    //Format decimal progress data
                    if (reportColumns.ToString().StartsWith("Pt"))
                    {
                        columnAddress.Style.Numberformat.Format = "#0.00%";

                        var bar = ws.ConditionalFormatting.AddDatabar(columnAddress, colorGenericGreen);
                        bar.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        bar.HighValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                        bar.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
                        bar.HighValue.Value = 1;
                        bar.LowValue.Value = 0;

                    }
                    
                    //Format date data
                    if (reportColumns.ToString().Contains("Date"))
                    {
                        columnAddress.Style.Numberformat.Format = "MMM-dd-yyyy";
                        //Conditionally Format completion dates

                        List<string> noFormatColList = new List<string>
                        {
                            "RAS",
                            "ESTIMATED",
                            "ITP",
                            "ADDED",
                        };

                        if (!noFormatColList.Any(x => reportColumns.ToString().ToUpper().Contains(x)))
                        {
                            string cellAddress = ws.Cells[1, (int)reportColumns].Address;
                            string statement = $"ISNUMBER({cellAddress})";

                            var dateCondition = ws.ConditionalFormatting.AddExpression(columnAddress);
                            dateCondition.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            dateCondition.Style.Fill.BackgroundColor.Color = colorGenericGreen;
                            dateCondition.Formula = statement;
                        }
                        
                        //Format CCO Estimate Date
                        if (reportColumns.ToString().Contains("EstimatedCCOCompleteDate"))
                        {
                            string cellAddress = ws.Cells[1, (int)reportColumns].Address;
                            var expandedAddress = ws.Cells[1, (int)reportColumns, spoolList.Count() + 1, (int)reportColumns];
                            string statement = $"(${cellAddress}=\"CCO Complete\")";

                            var CCOCompleteCondition = ws.ConditionalFormatting.AddExpression(expandedAddress);
                            CCOCompleteCondition.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            CCOCompleteCondition.Style.Fill.BackgroundColor.Color = colorGenericGreen;
                            CCOCompleteCondition.Formula = statement;

                            expandedAddress.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                    }




                    //Format ITP data
                    if (reportColumns.ToString().ToUpper().Contains("ITPSTATUS"))
                    {
                        //Conditional Formatting Accept
                        string cellAddress = ws.Cells[1, (int)reportColumns].Address;
                        var expandedItpAddress = ws.Cells[1, (int)reportColumns, spoolList.Count() + 1, (int)reportColumns + 1];
                        string statementAccept = $"AND(ISNUMBER(SEARCH(\"APPR\",${cellAddress})),NOT(ISNUMBER(SEARCH(\"NOT\",${cellAddress}))))";
                        string statementNotAccept = $"AND(NOT({statementAccept}),NOT(ISNUMBER(SEARCH(\"ITP\",${cellAddress}))))";

                        //Accept
                        var itpConditionAccept = ws.ConditionalFormatting.AddExpression(expandedItpAddress);
                        itpConditionAccept.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        itpConditionAccept.Style.Fill.BackgroundColor.Color = colorGenericGreen;
                        itpConditionAccept.Formula = statementAccept;
                        
                        //Not Accepted
                        var itpConditionNotAccept = ws.ConditionalFormatting.AddExpression(expandedItpAddress);
                        itpConditionNotAccept.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        itpConditionNotAccept.Style.Fill.BackgroundColor.Color = colorGenericOrange;
                        itpConditionNotAccept.Formula = statementNotAccept;

                    }

                    //Assign Border
                    columnAddress.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dotted;

                    //Hide/Unhide Column
                    ws.Column((int)reportColumns).Hidden = reportColumns.GetHiddenStatus();
                }

                //Additional formatting
                SetDataBarFormat(ws);


                //ws.Cells[ws.Dimension.Address].AutoFitColumns();
                //Write Spool Data

                int i = 2; //Start input data on row 2
                string previousSpoolWo = null;
                int oneOrZero = 0;

                foreach(Spool spool in spoolList)
                {
                    ExcelRange rowAddress  = ws.Cells[i, 1, i, Enum.GetNames(typeof(ReportColumns)).Length];

                    //Fill Spool Data
                    int j = 0;
                    foreach(ReportColumns reportColumns in (ReportColumns[])Enum.GetValues(typeof(ReportColumns)))
                    {
                        ws.Cells[i, (int)reportColumns].Value = spool.GetType().GetProperty(reportColumns.ToString()).GetValue(spool, null);

                        //Facing Rule
                        if(spool.NumFacesTotal == 0 && reportColumns == ReportColumns.NumFacesCompleted)
                        {
                            ws.Cells[i, (int)ReportColumns.NumFacesTotal, i, (int)reportColumns].Value = 0;
                        }

                        //Hydro Rule
                        if(Regex.IsMatch(spool.HydroPressure.ToUpper(), "SERVICE|NO") && reportColumns == ReportColumns.DateHydrotested)
                        {
                            ws.Cells[i, (int)reportColumns].Value = "NOT REQ";
                        }

                        //CCO Rule
                        if ((reportColumns == ReportColumns.PtCcoProgress || reportColumns == ReportColumns.EstimatedCCOCompleteDate) && !spool.boolCCO)
                        {
                            ws.Cells[i, (int)reportColumns].Style.Border.DiagonalDown = true;
                            ws.Cells[i, (int)reportColumns].Style.Border.DiagonalUp = true;
                            ws.Cells[i, (int)reportColumns].Style.Border.Diagonal.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

                        #region Overwrites

                        if (reportColumns.ToString().Contains("EstimatedCCOCompleteDate"))
                        {
                            if(ws.Cells[i, (int)ReportColumns.PtCcoProgress].Value != null && (decimal)ws.Cells[i,(int)ReportColumns.PtCcoProgress].Value > (decimal)0.98)
                            {
                                ws.Cells[i, (int)reportColumns].Value = "CCO Complete";
                            }
                        }

                        #endregion

                        j++;
                    }

                    //Formula & CF for over/under delivery date
                    var overUnderAddress = ws.Cells[i, (int)ReportColumns.WorkDuration];
                    string shipAddress = ws.Cells[i, (int)ReportColumns.EstimatedShipDate].Address.ToString();
                    string rasAddress = ws.Cells[i, (int)ReportColumns.RasDate].Address.ToString();
                    ws.Cells[i, (int)ReportColumns.WorkDuration].Formula = $"{rasAddress}-{shipAddress}";


                    if ((((DateTime)spool.RasDate).Subtract((DateTime)spool.EstimatedShipDate)).Days > 0)
                    {
                        var cficon = ws.ConditionalFormatting.AddThreeIconSet(overUnderAddress, eExcelconditionalFormatting3IconsSetType.Arrows);
                        cficon.Icon1.Type = eExcelConditionalFormattingValueObjectType.Num;
                    }
                    else if ((((DateTime)spool.RasDate).Subtract((DateTime)spool.EstimatedShipDate)).Days < 0)
                    {
                        var cficon = ws.ConditionalFormatting.AddThreeIconSet(overUnderAddress, eExcelconditionalFormatting3IconsSetType.Arrows);
                        cficon.Reverse = true;


                    }



                    //format rowcolor (Alternate Color between work orders)
                    if (spool.WorkOrder != previousSpoolWo)
                    {
                        oneOrZero = 1 - oneOrZero; //Toggle workorder change
                        //Thick job separator border
                        rowAddress.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        woCounter++;
                    }
                    else
                    {
                        //Thin border
                        rowAddress.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                    //Alternate filling color
                    var hue = (double)woCounter / numWorkOrders;
                    ColorWork.ColorRGB c = ColorWork.HSL2RGB(hue, 0.5, 0.5);
                    var fillColor = c;
                    var shortRowAddress = ws.Cells[i, (int)ReportColumns.PoNumber, i, (int)ReportColumns.WorkOrder];

                    rowAddress.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    rowAddress.Style.Fill.BackgroundColor.SetColor(colorPrimorisOne);

                    shortRowAddress.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    shortRowAddress.Style.Fill.BackgroundColor.SetColor(fillColor);




                    previousSpoolWo = spool.WorkOrder;
                    i++;

                }

                //Format special priorities
                CustomHighlight("2019", ws);

                p.Save();
            }
        }

        private void CustomHighlight(string value, ExcelWorksheet ws)
        {
            for(int i = 1; i < ws.Dimension.End.Row; i++)
            {
                if(ws.Cells[i, (int)ReportColumns.Priority].Value.ToString().Contains(value))
                {
                    ws.Row(i).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Row(i).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.HotPink);
                }
            }
        }


        private void SetDataBarFormat(ExcelWorksheet ws)
        {
            //Get reference to the worksheet xml for proper namespace
            var xdoc = ws.WorksheetXml;
            var nsm = new XmlNamespaceManager(xdoc.NameTable);
            nsm.AddNamespace("default", xdoc.DocumentElement.NamespaceURI);

            //Create the conditional format extension list entry
            var cfNodes = xdoc.SelectNodes("/default:worksheet/default:conditionalFormatting/default:cfRule", nsm);
            foreach (XmlNode cfnode in cfNodes)
            {
                var extLstCfNormal = xdoc.CreateNode(XmlNodeType.Element, "extLst", xdoc.DocumentElement.NamespaceURI);
                extLstCfNormal.InnerXml = @"<ext uri=""{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"" 
                                xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"">
                                <x14:id>{3F3F0E19-800E-4C9F-9CAF-1E3CE014ED86}</x14:id></ext>";

                cfnode.AppendChild(extLstCfNormal);
            }

            //Create the extension list content for the worksheet
            var extLstWs = xdoc.CreateNode(XmlNodeType.Element, "extLst", xdoc.DocumentElement.NamespaceURI);
            extLstWs.InnerXml = @"<ext uri=""{78C0D931-6437-407d-A8EE-F0AAD7539E65}"" 
                                            xmlns:x14=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"">
                                        <x14:conditionalFormattings>
                                        <x14:conditionalFormatting xmlns:xm=""http://schemas.microsoft.com/office/excel/2006/main"">
                                        <x14:cfRule type=""dataBar"" id=""{3F3F0E19-800E-4C9F-9CAF-1E3CE014ED86}"">
                                            <x14:dataBar minLength=""0"" maxLength=""100"" gradient=""0"">
                                            <x14:cfvo type=""num"">
                                                <xm:f>0</xm:f>
                                            </x14:cfvo>
                                            <x14:cfvo type=""num"">
                                                <xm:f>100</xm:f>
                                            </x14:cfvo>
                                            <x14:negativeFillColor rgb=""FFFF0000""/><x14:axisColor rgb=""FF000000""/>
                                            </x14:dataBar>
                                        </x14:cfRule>
                                        <xm:sqref>A1:A20</xm:sqref>
                                        </x14:conditionalFormatting>
                                        <x14:conditionalFormatting xmlns:xm=""http://schemas.microsoft.com/office/excel/2006/main"">
                                            <x14:cfRule type=""dataBar"" id=""{3F3F0E19-800E-4C9F-9CAF-1E3CE014ED86}"">
                                            <x14:dataBar minLength=""0"" maxLength=""100"" gradient=""0"">
                                                <x14:cfvo type=""num"">
                                                <xm:f>0</xm:f>
                                                </x14:cfvo><x14:cfvo type=""num"">
                                                <xm:f>200</xm:f>
                                                </x14:cfvo><x14:negativeFillColor rgb=""FFFF0000""/>
                                                <x14:axisColor rgb=""FF000000""/>
                                            </x14:dataBar>
                                            </x14:cfRule>
                                            <xm:sqref>B1:B20</xm:sqref>
                                        </x14:conditionalFormatting>
                                        </x14:conditionalFormattings>
                                    </ext>";
            var wsNode = xdoc.SelectSingleNode("/default:worksheet", nsm);
            wsNode.AppendChild(extLstWs);
        }


    }
}
