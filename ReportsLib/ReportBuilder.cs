using FastReport;
using FastReport.Utils;
using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using FastReport.Table;
using System.Text.RegularExpressions;
using System.Drawing.Printing;
using System.IO;
using System.Net;
using System.Text;
using FastReport.Data;
using System.Data;
using FastReport.Design;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ReportsLib
{
    class ReportBuilder
    {
        FastReport.Report report = null;
        string ConnectionString;
        bool ShowPrint = true;
        int IdRes = 0;
        int ReqId = 0;
        bool ShowFrDialog = true;
        bool NeedClear = true;
        List<ResultInfo> Results = null;

        // for MatrixManualBuilding
        FirebirdSql.Data.FirebirdClient.FbDataReader fb_reader = null;
        List<string> fields = null;
        List<string> fieldCaptions = null;
        List<string> fieldWidths = null;
        FastReport.Matrix.MatrixObject matrix = null;
        int FontSize = 10;
        string FontName = "Arial";


        public ReportBuilder(string aConnectionString, int aIdRes, int aReqId, bool aShowPrint, bool aShowFrDialog, bool aNeedClear)
        {
            if (!RegisteredObjects.IsTypeRegistered(typeof(FirebirdDataConnection)))
                RegisteredObjects.AddConnection(typeof(FirebirdDataConnection));

            ConnectionString = aConnectionString;
            ShowPrint = aShowPrint;
            IdRes = aIdRes;
            ReqId = aReqId;
            ShowFrDialog = aShowFrDialog;
            NeedClear = aNeedClear;
            Results = null;
            report = new FastReport.Report();
            // Set report option to shor or to hide progress window
            Config.ReportSettings.ShowProgress = aShowPrint;
            // Set report preview option to hide Edit button 
            Config.PreviewSettings.Buttons = PreviewButtons.All ^ PreviewButtons.Edit ^ PreviewButtons.PageSetup;
            Config.PreviewSettings.ShowInTaskbar = true;
            Config.DesignerSettings.ShowInTaskbar = true;

            report.PrintSettings.Printer = this.GetDefaultPrinter();
            report.PrintSettings.ShowDialog = false;
        }

        public ReportBuilder(string aConnectionString, List<ResultInfo> aResults, int aReqId, bool aShowPrint, bool aShowFrDialog, bool aNeedClear)
        {
            if (!RegisteredObjects.IsTypeRegistered(typeof(FirebirdDataConnection)))
                RegisteredObjects.AddConnection(typeof(FirebirdDataConnection));

            ConnectionString = aConnectionString;
            ShowPrint = aShowPrint;
            IdRes = 0;
            ReqId = aReqId;
            ShowFrDialog = aShowFrDialog;
            NeedClear = aNeedClear;
            Results = aResults;
            report = new FastReport.Report();
            // Set report option to shor or to hide progress window
            Config.ReportSettings.ShowProgress = aShowPrint;
            // Set report preview option to hide Edit button 
            Config.PreviewSettings.Buttons = PreviewButtons.All ^ PreviewButtons.Edit ^ PreviewButtons.PageSetup;
            Config.PreviewSettings.ShowInTaskbar = true;
            Config.DesignerSettings.ShowInTaskbar = true;

            report.PrintSettings.Printer = this.GetDefaultPrinter();
            report.PrintSettings.ShowDialog = false;
        }

        /*
         * 
         * */
         public void SomeAction(object sender, EventArgs e)
        {

        }

        /*
         * The function builds a reports in accordation of its structure
         * It can show or print report, if the given report is a report of the top level
         * 
         * The possibility of printin/showing of the report is used within user's applications ONLY!
         * 
         * The function is recoursive: it calls itself with Secondary == true to prepare subreports
         * for the reports of the top levels
         * 
         * @param idReport - ID of the report's form to build
         * @param Secondary - if the report is a report of the top level, then this parameter is equal to TRUE,
         * other way it's equal to FALSE
         * 
         * */

        public void BuildReport(int idReport, bool Secondary = false)
        {
            Log.Write(idReport, "START report building");
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }
            Log.Write(idReport, "BuildReport - DB connected");

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand SelectReport = null;
            FbCommand SubReports = null;
            FbDataReader readerForm = null;

            List<int> SubRportIds = new List<int>();

            if (!Secondary)
                report.Clear();

            try
            {
                try
                {
                    SelectReport = new FbCommand("select form from reportforms where id = " + Convert.ToString(idReport), fb_con, fb_trans);
                    Log.Write(idReport, "BuildReport - form fetched" + report.ToString());

                    readerForm = SelectReport.ExecuteReader();

                    if (readerForm.Read())
                    {
                        byte[] reportForm = null;
                        string reportStr = "";
                        try
                        {
                            reportForm = (byte[])readerForm["form"];
                            reportStr = System.Text.Encoding.UTF8.GetString(reportForm);
                        }
                        catch
                        {
                        }

                        if (reportStr != "")
                        {
                            report.LoadFromString(reportStr);
                            report.Dictionary.Connections[0].ConnectionString = ConnectionString;
                            report.Dictionary.Connections[0].Name = "VoteDB";
                            report.Dictionary.Connections[0].Enabled = true;
                            report.SetParameterValue("IDRES", IdRes);
                            report.SetParameterValue("REQID", ReqId);
                            report.SetParameterValue("FRSELECT", ShowFrDialog);
                            report.SetParameterValue("NEEDCLEAR", NeedClear);
                            report.Prepare(true);
                            Log.Write(idReport, "BuildReport - report prepared ");
                        }
                    }

                    SubReports = new FbCommand("select subreport from subreportlink where report = " + Convert.ToString(idReport) +" order by id"
                        , fb_con, fb_trans);
                    readerForm = SubReports.ExecuteReader();

                    SubRportIds.Clear();

                    Log.Write(idReport, "BuildReport - subreports fetched ");

                    while (readerForm.Read())
                    {
                        Log.Write(idReport, "BuildReport - add subreport "+ readerForm[0].ToString());
                        SubRportIds.Add((int)readerForm[0]);
                    }

                    // fb_trans.Commit();

                    for (int i = 0; i < SubRportIds.Count; i++)
                    {
                        Log.Write(idReport, "BuildReport - recoursive call " + SubRportIds[i].ToString());
                        BuildReport((int)SubRportIds[i], true);
                    }

                    if (!Secondary)
                    {
                        if (ShowPrint)
                        {
                            report.ShowPrepared(false,null);
                        }
                        else
                        {
                            report.PrintPrepared();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Write(idReport, " =  REPORT error building ", ex);
                }
            }
            finally
            {
                if (readerForm != null)
                    readerForm.Close();

                if (SelectReport != null)
                    SelectReport.Dispose();

                if (SubReports != null)
                    SubReports.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();

                if (SubRportIds != null)
                {
                    SubRportIds.Clear();
                }
            }

        }

        /*
         *  The function is used to prepare data for summary report creation
         *  Then it calls BuildReport function to build a report and show/print it
         *  
         *  This function is for user's applications only!
         * */

        public void BuildSummaryReport(int idReport)
        {
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            String DelSQL = "execute procedure ClearReportReq(" + ReqId.ToString() + ",1)";
            String InsSQL = "insert into RepRequestResult(reqid, result) values ("+ReqId.ToString()+", @RESID)";
            String DecodeRes = "execute procedure DecodeReqResult(" + ReqId.ToString() + ", 1)";
            // String PrepareFract = "execute procedure PrepareFrList("+ReqId.ToString()+")";
            FbCommand fb_com = new FbCommand(InsSQL, fb_con, fb_trans);

            FbParameter par_resid = new FbParameter("@RESID", FbDbType.Integer);

            try
            {
                try
                {
                    // Delete old results
                    fb_com.CommandText = DelSQL;
                    fb_com.ExecuteNonQuery();

                    // Insert new results
                    fb_com.CommandText = InsSQL;
                    for (int i = 0; i < Results.Count; i++)
                    {
                        fb_com.Parameters.Clear();
                        par_resid.Value = Results[i].IdRes;
                        fb_com.Parameters.Add(par_resid);
                        fb_com.ExecuteNonQuery();
                    }

                    // Decode inserted results
                    fb_com.CommandText = DecodeRes;
                    fb_com.Connection = fb_con;
                    fb_com.Transaction = fb_trans;

                    fb_com.ExecuteNonQuery();

                    fb_trans.Commit();

                    BuildReport(idReport);

                }
                catch (Exception ex)
                {
                    fb_trans.Rollback();
                    MessageBox.Show("Error during preparing results for the summary report = "+ex.ToString());
                }

            }
            finally
            {

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_com != null)
                    fb_com.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();

            }
        }

        /*
         * The function creates a report according its structure,
         * but doesn't show nor print it
         * 
         * It returns the report object as a result
         * */

        public void PrepareReport(int idReport, bool Secondary = false)
        {

            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand SelectReport = null;
            FbCommand SubReports = null;
            FbDataReader readerForm = null;

            List<int> SubRportIds = new List<int>();

            if (!Secondary)
                 report.Clear();

            try
            {
                try
                {
                    SelectReport = new FbCommand("select form from reportforms where id = " + Convert.ToString(idReport), fb_con, fb_trans);

                    readerForm = SelectReport.ExecuteReader();

                    if (readerForm.Read())
                    {
                        byte[] reportForm = null;
                        string reportStr = "";
                        try
                        {
                            Log.Write(idReport, " Form extracted = "+ Convert.ToString(idReport));
                            reportForm = (byte[])readerForm["form"];
                            reportStr = System.Text.Encoding.UTF8.GetString(reportForm);
                        }
                        catch
                        {
                        }

                        if (reportStr != "")
                        {
                            report.LoadFromString(reportStr);
                            report.Dictionary.Connections[0].ConnectionString = ConnectionString;
                            report.Dictionary.Connections[0].Name = "VoteDB";
                            report.Dictionary.Connections[0].Enabled = true;
                            report.SetParameterValue("IDRES", IdRes);
                            report.SetParameterValue("REQID", ReqId);
                            report.SetParameterValue("FRSELECT", ShowFrDialog);
                            report.SetParameterValue("NEEDCLEAR", NeedClear);
                            report.Prepare(true);
                        }
                    }

                    SubReports = new FbCommand("select subreport from subreportlink where report = " + Convert.ToString(idReport) + " order by id"
                        , fb_con, fb_trans);
                    readerForm = SubReports.ExecuteReader();

                    SubRportIds.Clear();

                    while (readerForm.Read())
                    {
                        SubRportIds.Add((int)readerForm[0]);
                    }

                    // fb_trans.Commit();

                    for (int i = 0; i < SubRportIds.Count; i++)
                    {
                        PrepareReport((int)SubRportIds[i], true);
                    }


                }
                catch (Exception ex)
                {
                    Log.Write(this.IdRes, " =  REPORT error building ", ex);
                    throw new Exception(ex.Message);
                }
            }
            finally
            {
                if (readerForm != null)
                    readerForm.Close();

                if (SelectReport != null)
                    SelectReport.Dispose();

                if (SubReports != null)
                    SubReports.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();

                if (SubRportIds != null)
                {
                    SubRportIds.Clear();
                }
            }

        }

        /*
         * Function for preparing of the aggregative report
         * */

        public void PrepareSummaryReport(int idReport, DateTime ResDate)
        {
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            String DelSQL = "execute procedure ClearReportReq(" + ReqId.ToString() + ",1)";
            String InsSQL = "insert into RepRequestResult(reqid, result) values (" + ReqId.ToString() + ", @RESID)";
            String DecodeRes = "execute procedure DecodeReqResult(" + ReqId.ToString() + ", 1)";
            FbCommand fb_com = new FbCommand(InsSQL, fb_con, fb_trans);

            // Parameters for insert results' IDs into temporary table
            // for results
            FbParameter par_resid = new FbParameter("@RESID", FbDbType.Integer);

            // Prepare parameters for result select query
            DateTime StartDate = new DateTime(ResDate.Year, ResDate.Month, ResDate.Day);
            DateTime EndDate = new DateTime((ResDate.AddDays(1)).Year, (ResDate.AddDays(1)).Month, (ResDate.AddDays(1).Day));

            Log.Write(0, "Strat Date = "+StartDate.ToString());
            Log.Write(0, "End Date = " + EndDate.ToString());


            String GetResults = "select id from ViewResults(0, 0, -1, '"+
                                    StartDate.ToString("dd.MM.yyyy") + "', '"+
                                    EndDate.ToString("dd.MM.yyyy") + "', 0)";


            FbDataReader fb_reader = null;

            try
            {
                try
                {
                    // Get new results

                    if (this.Results == null)
                        this.Results = new List<ResultInfo>();
                    else
                        this.Results.Clear();

                    fb_com.CommandText = GetResults;
                    fb_reader = fb_com.ExecuteReader();

                    // Add results' IDs into the internal results' list
                    while (fb_reader.Read())
                    {
                        this.Results.Add(new ResultInfo(Convert.ToInt32(fb_reader["id"]), 0, 0));
                        Log.Write(Convert.ToInt32(fb_reader["id"]), "added ");
                    }
                    
                    // Remove old data
                    fb_com.CommandText = DelSQL;
                    fb_com.ExecuteNonQuery();

                    // Insert new results
                    fb_com.CommandText = InsSQL;
                    for (int i = 0; i < Results.Count; i++)
                    {
                        fb_com.Parameters.Clear();
                        par_resid.Value = Results[i].IdRes;
                        fb_com.Parameters.Add(par_resid);
                        fb_com.ExecuteNonQuery();
                    }

                    // Decode inserted results
                    fb_com.CommandText = DecodeRes;
                    fb_com.ExecuteNonQuery();

                    fb_trans.Commit();

                    PrepareReport(idReport);

                }
                catch (Exception ex)
                {
                    fb_trans.Rollback();
                    Log.Write(idReport, "Error during preparing results for the summary report = ",ex);
                }

            }
            finally
            {
                if (fb_reader != null && !fb_reader.IsClosed)
                    fb_reader.Close();

                if (fb_com != null)
                    fb_com.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();

            }
        }
        /*
         * The function replaces the template symbols within the file name template
         * with the values, which are obtained from the result's date
         * */

        private string ReplaceFormatSymbols(string aFileName, DateTime aResDate)
        {
            string result = aFileName;

            result = result.Replace("yyyy", aResDate.Year.ToString("0000"));
            result = result.Replace("mm", aResDate.Month.ToString("00"));
            result = result.Replace("dd", aResDate.Day.ToString("00"));
            result = result.Replace("hh", aResDate.Hour.ToString("00"));
            result = result.Replace("nn", aResDate.Minute.ToString("00"));
            result = result.Replace("ss", aResDate.Second.ToString("00"));
            result = result.Replace("%\"D", "");
            result = result.Replace("D\"0:s", "");

            return result;
        }

        /*
         * The function call PrepareReport to generate the report's object
         * Then it export the report into the file of the determinated format
         * */


        public void ExportReportFile(int idReport, int idFormat, string aPath, string aFileName, DateTime aResDate)
        {
            try
            {
                // FastReport.Report loc_report = PrepareReport(idReport);
                Log.Write(this.IdRes, " = REPORT Export started " + idReport.ToString());
                PrepareReport(idReport);

                string FileName = aPath + ReplaceFormatSymbols(aFileName, aResDate);

                switch (idFormat)
                {
                    case 1: // TXT
                        FileName += ".txt";
                        FastReport.Export.Text.TextExport txt_export = new FastReport.Export.Text.TextExport();
                        txt_export.Export(report, FileName);
                        break;
                    case 2: // HTML
                        FileName += ".html";
                        FastReport.Export.Html.HTMLExport html_export = new FastReport.Export.Html.HTMLExport();
                        html_export.Export(report, FileName);
                        break;
                    case 3: // CSV
                        FileName += ".csv";
                        FastReport.Export.Csv.CSVExport csv_export = new FastReport.Export.Csv.CSVExport();
                        csv_export.Export(report, FileName);
                        break;
                    case 4: // RTF
                        FileName += ".rft";
                        FastReport.Export.RichText.RTFExport rtf_export = new FastReport.Export.RichText.RTFExport();
                        rtf_export.Export(report, FileName);
                        break;
                    case 5: // PDF
                    default:
                        FileName += ".pdf";
                        FastReport.Export.Pdf.PDFExport pdf_export = new FastReport.Export.Pdf.PDFExport();
                        pdf_export.Export(report, FileName);
                        break;
                }
                Log.Write(this.IdRes, " = FileName " + FileName);
                Log.Write(this.IdRes, " = REPORT exported " + report.ToString());
            }
            catch (Exception e)
            {
                Log.Write(this.IdRes, "EXPORT internal exception " + idFormat.ToString(), e);
                throw new Exception(e.Message);
            }
        }

        /*
         * For automatical saving an aggregative report into the file
         * */
        public void ExportSumReportFile(int idReport, int idFormat, string aPath, string aFileName, DateTime aResDate)
        {
            try
            {
                // FastReport.Report loc_report = PrepareReport(idReport);
                Log.Write(this.IdRes, " = REPORT Export started " + idReport.ToString());
                PrepareSummaryReport(idReport, aResDate);

                string FileName = aPath + ReplaceFormatSymbols(aFileName, aResDate);

                switch (idFormat)
                {
                    case 1: // TXT
                        FileName += ".txt";
                        FastReport.Export.Text.TextExport txt_export = new FastReport.Export.Text.TextExport();
                        txt_export.Export(report, FileName);
                        break;
                    case 2: // HTML
                        FileName += ".html";
                        FastReport.Export.Html.HTMLExport html_export = new FastReport.Export.Html.HTMLExport();
                        html_export.Export(report, FileName);
                        break;
                    case 3: // CSV
                        FileName += ".csv";
                        FastReport.Export.Csv.CSVExport csv_export = new FastReport.Export.Csv.CSVExport();
                        csv_export.Export(report, FileName);
                        break;
                    case 4: // RTF
                        FileName += ".rft";
                        FastReport.Export.RichText.RTFExport rtf_export = new FastReport.Export.RichText.RTFExport();
                        rtf_export.Export(report, FileName);
                        break;
                    case 5: // PDF
                    default:
                        FileName += ".pdf";
                        FastReport.Export.Pdf.PDFExport pdf_export = new FastReport.Export.Pdf.PDFExport();
                        pdf_export.Export(report, FileName);
                        break;
                }
                Log.Write(this.IdRes, " = FileName " + FileName);
                Log.Write(this.IdRes, " = REPORT exported " + report.ToString());
            }
            catch (Exception e)
            {
                Log.Write(this.IdRes, "EXPORT internal exception " + idFormat.ToString(), e);
                throw new Exception(e.Message);
            }
        }

        /*
         * The function for printing of the report during the AutoPrintService's job
         * Reloaded function!!!
         * @param idReport - the ID of the report, that must be printed
         * @param idAuto - the ID of the automatiacal task
         * 
         * In this version of the function, the printer's parameters must be obtained 
         * during the function's execution using the ID of the automatical task
         * 
         * */

        public void PrintReport(int idReport, int idAuto)
        {
            try
            {
                // Prepare the report object
                Log.Write(this.IdRes, " PRINT Start creation ");

                // FastReport.Report loc_report = PrepareReport(idReport);
                PrepareReport(idReport);
                Log.Write(this.IdRes, " PRINT report created ");

                // Printer's parameters
                string aPrinter = "";
                string aPort = "";
                byte[] aPrinterSettings = null;

                // Get printer's parameters
                int idPrinter = GetPrinerParameters(idAuto, out aPrinter, out aPort, out aPrinterSettings);
                Log.Write(this.IdRes, " PRINT Printer detected ");

                Log.Write(this.IdRes, " PRINT Printer selected " + idPrinter.ToString() + " for IDAUTO " + idAuto.ToString());
                // Set the printing parameters to the report
                if (idPrinter != -1)
                {
                    Log.Write(this.IdRes, " PRINT Printer selected " + aPrinter);
                    SetPrinterMode(aPrinter, aPrinterSettings);
                    report.PrintSettings.Printer = aPrinter;
                    Log.Write(this.IdRes, " PRINT Printer is set " + report.PrintSettings.Printer);
                }
                else
                {
                    report.PrintSettings.Printer = GetDefaultPrinter();
                    Log.Write(this.IdRes, " PRINT DEFAULT Printer is set " + report.PrintSettings.Printer);
                }


                report.PrintSettings.ShowDialog = false;
                report.PrintPrepared();
                Log.Write(this.IdRes, " PRINT report printed ");
            }
            catch (Exception e)
            {
                Log.Write(this.IdRes, "PRINT "+idReport.ToString()+" "+idAuto.ToString(), e);
                throw new Exception(e.Message);
            }
        }

        /*
         * For automatical print of the aggregative report
         * */
        public void PrintSumReport(int idReport, int idAuto, DateTime aResDate)
        {
            try
            {
                // Prepare the report object
                Log.Write(this.IdRes, " PRINT Start creation ");

                // FastReport.Report loc_report = PrepareReport(idReport);
                PrepareSummaryReport(idReport, aResDate);
                Log.Write(this.IdRes, " PRINT report created ");

                // Printer's parameters
                string aPrinter = "";
                string aPort = "";
                byte[] aPrinterSettings = null;

                // Get printer's parameters
                int idPrinter = GetPrinerParameters(idAuto, out aPrinter, out aPort, out aPrinterSettings);
                Log.Write(this.IdRes, " PRINT Printer detected ");

                Log.Write(this.IdRes, " PRINT Printer selected " + idPrinter.ToString() + " for IDAUTO " + idAuto.ToString());
                // Set the printing parameters to the report
                if (idPrinter != -1)
                {
                    Log.Write(this.IdRes, " PRINT Printer selected " + aPrinter);
                    SetPrinterMode(aPrinter, aPrinterSettings);
                    report.PrintSettings.Printer = aPrinter;
                    Log.Write(this.IdRes, " PRINT Printer is set " + report.PrintSettings.Printer);
                }
                else
                {
                    report.PrintSettings.Printer = GetDefaultPrinter();
                    Log.Write(this.IdRes, " PRINT DEFAULT Printer is set " + report.PrintSettings.Printer);
                }

                report.PrintSettings.ShowDialog = false;
                report.PrintPrepared();
                Log.Write(this.IdRes, " PRINT report printed ");
            }
            catch (Exception e)
            {
                Log.Write(this.IdRes, "PRINT " + idReport.ToString() + " " + idAuto.ToString(), e);
                throw new Exception(e.Message);
            }
        }

        /*
         * The function gets an information about printer in dependence on the ID of the automatical task
         * */

        private int GetPrinerParameters(int idAuto, out string aPrinter, out string aPort, out byte[] aPrinterSettings)
        {
            int result = -1;

            aPrinter = "";
            aPort = "";
            aPrinterSettings = null;

            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand fb_com = new FbCommand("select printers.* from printers join devoptions on printers.id = devoptions.printer "+
                " join autoresults on devoptions.id = autoresults.device "+
                " where autoresults.id = "+idAuto.ToString(),
                fb_con,
                fb_trans);

            FbDataReader fb_reader = null;

            try
            {
                try
                {
                    fb_reader = fb_com.ExecuteReader();
                    if(fb_reader.Read())
                    {
                        if ((int)fb_reader["id"] != -1)
                        {
                            result = (int)fb_reader["id"];
                            aPrinter = (string)fb_reader["device"];
                            aPort = (string)fb_reader["port"];
                            aPrinterSettings = (byte[])fb_reader["bindata"];

                        }
                    }

                    fb_reader.Close();
                    fb_trans.Commit();

                }
                catch 
                {
                    fb_trans.Rollback();
                }

            }
            finally
            {
                if (fb_reader != null)
                    fb_reader.Close();

                if (fb_com != null)
                    fb_com.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();
            }

            return result;
        }


        /*
         * Get default printer
         * */

        private string GetDefaultPrinter()
        {
            PrinterSettings settings = new PrinterSettings();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                settings.PrinterName = printer;
                if (settings.IsDefaultPrinter)
                    return printer;
            }
            return string.Empty;
        }

        /*
         * Get printer settings from system registry
         * */

        private void GetPrinterSettingsFromRegistry(string PrinterRegKey, string PrinterRegSection, int PrintRegItem,
                                                    out string aPrinter, out byte[] aPrinterSettings)
        {
            aPrinter = "";
            aPrinterSettings = null;

            Log.Write(0, "Registry key = " + PrinterRegKey);
            Log.Write(0, "Registry section = " + PrinterRegSection);

            try
            {
                using (RegistryKey printerKey = Registry.CurrentUser.OpenSubKey(PrinterRegKey))
                {
                    if (printerKey != null)
                    {
                        Log.Write(0, PrinterRegSection + Convert.ToString(PrintRegItem) + " Device");
                        aPrinter = (string)printerKey.GetValue(PrinterRegSection + Convert.ToString(PrintRegItem) + " Device", "");
                        Log.Write(0, "Registry printer = " + aPrinter);
                        aPrinterSettings = (byte[])printerKey.GetValue(PrinterRegSection + Convert.ToString(PrintRegItem) + " Data", null);
                    }
                }
            }
            catch (Exception e)
            {
                Log.Write(0, "Registry read exception = ", e);
            }

        }

        /*
         * Set printer parameters readed form database for given task
         * 
         * */

        private PrinterSettings SetPrinterMode(string aPrinter, byte[] aPrinterSettings)
        {
            PrinterSettings settings = new PrinterSettings();
            settings.PrinterName = aPrinter;
            IntPtr pDevMode = Marshal.AllocHGlobal(aPrinterSettings.Length);
            try
            {
                Marshal.Copy(aPrinterSettings, 0, pDevMode, aPrinterSettings.Length);
                settings.SetHdevmode(pDevMode);
            }
            catch (Exception e)
            {
                Log.Write(0, "Printer settings error = ", e);
            }
            finally
            {
                Marshal.FreeHGlobal(pDevMode);
            }
            return settings;
        }
        
        /*
         * The function converts a report into the determinated format and 
         * then sends it to the given FTP-server
         * 
         * */

        private void SendtoFTP(string Location, string UserName, string Password, MemoryStream stream)
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Location);
                request.Method = WebRequestMethods.Ftp.UploadFile;
                Log.Write(this.IdRes, "FTP request created ");
                request.Credentials = new NetworkCredential(UserName, Password);
                request.UsePassive = true;
                request.UseBinary = true;
                request.KeepAlive = true;
                request.Timeout = 1000;


                int len = (int)stream.Length;
                byte[] buffer = new byte[len];
                stream.Position = 0;
                int read = stream.Read(buffer, 0, len);
                Log.Write(this.IdRes, "FTP byte read " + len);
                request.ContentLength = read;

                Stream requestStream = request.GetRequestStream();
                requestStream.Write(buffer, 0, read);
                Log.Write(this.IdRes, "FTP sent " + read.ToString());
                requestStream.Close();
                Log.Write(this.IdRes, "FTP closed ");
            }
            catch (Exception ex)
            {
                Log.Write(0, "FTP during sending ", ex);
            }
        }

        public void SendToFtp(int idReport, 
            int idFormat, string aFileName, DateTime aResDate,
            string aServer, string aLogin, string aPassword, string aServerPath
            )
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                // FastReport.Report loc_report = PrepareReport(idReport);
                Log.Write(this.IdRes, "FTP start " + idReport.ToString());
                PrepareReport(idReport);
                Log.Write(this.IdRes, "FTP report prepared " + idReport.ToString());
                string FileName = ReplaceFormatSymbols(aFileName, aResDate);

                // Set the parameters of the FTP-server   
                Log.Write(this.IdRes, "FTP start create connection " + idReport.ToString());
                FastReport.Cloud.StorageClient.Ftp.FtpStorageClient ftp = new FastReport.Cloud.StorageClient.Ftp.FtpStorageClient();
                ftp.Server = aServer + "/" + aServerPath;
                ftp.Username = aLogin;
                ftp.Password = aPassword;

                Log.Write(this.IdRes, "FTP connection established ");

                // Export report and send it to FTP
                switch (idFormat)
                {
                    case 1: // TXT
                        FileName += ".txt";
                        FastReport.Export.Text.TextExport txt_export = new FastReport.Export.Text.TextExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, txt_export);
                        break;
                    case 2: // HTML
                        FileName += ".html";
                        FastReport.Export.Html.HTMLExport html_export = new FastReport.Export.Html.HTMLExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, html_export);
                        break;
                    case 3: // CSV
                        FileName += ".csv";
                        FastReport.Export.Csv.CSVExport csv_export = new FastReport.Export.Csv.CSVExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, csv_export);
                        break;
                    case 4: // RTF
                        FileName += ".rft";
                        FastReport.Export.RichText.RTFExport rtf_export = new FastReport.Export.RichText.RTFExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, rtf_export);
                        break;
                    case 5: // PDF
                    default:
                        FileName += ".pdf";
                        FastReport.Export.Pdf.PDFExport pdf_export = new FastReport.Export.Pdf.PDFExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, pdf_export);
                        break;
                }
                Log.Write(this.IdRes, "FTP SENT " + FileName);
                Log.Write(this.IdRes, "FTP SENT " + idReport.ToString());

            }
            catch (Exception e)
            {
                Log.Write(this.IdRes, "SEND TO FTP " + idFormat.ToString(), e);
                throw new Exception(e.Message);
            }
        }

        /* 
         * For automatical sending of the aggregative report to the FTP server
         * */

        public void SendSumToFtp(int idReport,
            int idFormat, string aFileName, DateTime aResDate,
            string aServer, string aLogin, string aPassword, string aServerPath
            )
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                // FastReport.Report loc_report = PrepareReport(idReport);
                Log.Write(this.IdRes, "FTP start " + idReport.ToString());
                PrepareSummaryReport(idReport, aResDate);
                Log.Write(this.IdRes, "FTP report prepared " + idReport.ToString());
                string FileName = ReplaceFormatSymbols(aFileName, aResDate);

                // Set the parameters of the FTP-server   
                Log.Write(this.IdRes, "FTP start create connection " + idReport.ToString());
                FastReport.Cloud.StorageClient.Ftp.FtpStorageClient ftp = new FastReport.Cloud.StorageClient.Ftp.FtpStorageClient();
                ftp.Server = aServer + "/" + aServerPath;
                ftp.Username = aLogin;
                ftp.Password = aPassword;

                Log.Write(this.IdRes, "FTP connection established ");

                // Export report and send it to FTP
                switch (idFormat)
                {
                    case 1: // TXT
                        FileName += ".txt";
                        FastReport.Export.Text.TextExport txt_export = new FastReport.Export.Text.TextExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, txt_export);
                        break;
                    case 2: // HTML
                        FileName += ".html";
                        FastReport.Export.Html.HTMLExport html_export = new FastReport.Export.Html.HTMLExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, html_export);
                        break;
                    case 3: // CSV
                        FileName += ".csv";
                        FastReport.Export.Csv.CSVExport csv_export = new FastReport.Export.Csv.CSVExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, csv_export);
                        break;
                    case 4: // RTF
                        FileName += ".rft";
                        FastReport.Export.RichText.RTFExport rtf_export = new FastReport.Export.RichText.RTFExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, rtf_export);
                        break;
                    case 5: // PDF
                    default:
                        FileName += ".pdf";
                        FastReport.Export.Pdf.PDFExport pdf_export = new FastReport.Export.Pdf.PDFExport();
                        report.FileName = FileName;
                        ftp.SaveReport(report, pdf_export);
                        break;
                }
                Log.Write(this.IdRes, "FTP SENT " + FileName);
                Log.Write(this.IdRes, "FTP SENT " + idReport.ToString());

            }
            catch (Exception e)
            {
                Log.Write(this.IdRes, "SEND TO FTP " + idFormat.ToString(), e);
                throw new Exception(e.Message);
            }
        }


        /*
         * New and universal realization of the ShowGrid
         * @param Query - SQL-query obtained from the application. This string contains SQL written using the Firebird
         *                syntax, so the parameters are preceded by ":" symbol. Before using this string, it must be
         *                corrected via replacement the ":" symbols with "@"
         * @param Parameters - a list of the Query parameters' names
         * @paran ParamValues - a list of the Query parameters' values. Both of the lists, linked with query parameters, 
         *                must be ordered in the way, thus a parameter's name and its value would have had the same position
         *                in the lists
         * @param Fields - a list of the visible fields of the Query (according the condition of the data grid in the application)
         * @param FieldCaptions - a list of columns titles of the grid
         * @param FieldWidths - a list of the columns widths of the grid
         * @param FontName - a name of the font, which wsa used for the grid in the application
         * @param FontSize - a size of the given font
         * 
         * */
        public void ShowGrid(String Query,
            String Parameters,
            String ParamValues,
            String Fields,
            String FieldCaptions,
            String FieldWidths,
            String FontName,
            int FontSize,
            string ReportCaption,
            string PrinterRegKey = "",
            string PrinterRegSection = "",
            int PrintRegItem = 0)
        {

            // Preparing of the database connection and transaction
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand fb_com = null;
            // In this case we use the global fb_reader
            // because we'll need to pass this object into the
            // event handler for building the report in the manual mode
            fb_reader = null;

            this.FontName = FontName;
            this.FontSize = FontSize;

            // Preparing of the SQL query
            string sql = Query.Replace(':','@');

            try
            {
                try
                {
                    fb_com = new FbCommand(sql, fb_con, fb_trans);
                    fb_com.Parameters.Clear();

                    // Set parameters values
                    List<string> parameters = new List<string>(Parameters.Split(';'));
                    List<string> paramValues = new List<string>(ParamValues.Split(';'));
                    for (int i = 0; i<parameters.Count; i++)
                    {
                        FbParameter param = new FbParameter(parameters[i], (object)(paramValues[i]));
                        fb_com.Parameters.Add(param);
                    }

                    // Read data from the query
                    fb_reader = fb_com.ExecuteReader();

                    // Define the field list according the structure of the grid
                    fields = new List<string>(Fields.Split(';'));
                    fieldCaptions = new List<string>(FieldCaptions.Split(';'));
                    fieldWidths = new List<string>(FieldWidths.Split(';'));

                    // Report building
                    report.Clear();

                    ReportPage page = new ReportPage();
                    page.Name = "Page1";
                    page.PaperWidth = 210;
                    page.PaperHeight = 297;
                    // create report title
                    page.ReportTitle = new FastReport.ReportTitleBand();
                    page.ReportTitle.Name = "ReportTitle1";
                    page.ReportTitle.Height = Units.Millimeters * 10;
                    page.ReportTitle.CanGrow = true;
                    // show title
                    TextObject text = new TextObject();
                    text.Name = "Text1";
                    // text.Bounds = new System.Drawing.RectangleF(0, 0, Units.Millimeters * 100, Units.Millimeters * 5);
                    text.Dock = DockStyle.Fill;
                    text.Font = new System.Drawing.Font(FontName, FontSize);
                    text.Text = ReportCaption; // here must be caption !!!!
                    text.CanGrow = true;
                    page.ReportTitle.Objects.Add(text);

                    //  Create page header
                    page.PageHeader = new FastReport.PageHeaderBand();
                    page.PageHeader.Name = "PageHeader1";
                    page.PageHeader.Height = Units.Millimeters * 10;
                    page.PageHeader.CanGrow = false;
                    // Page number show
                    TextObject text_page = new TextObject();
                    text_page.Name = "Text2";
                    text_page.Dock = DockStyle.Fill;
                    text_page.Text = "[Page#]";
                    text_page.CanGrow = false;
                    text_page.HorzAlign = HorzAlign.Center;
                    page.PageHeader.Objects.Add(text_page);

                    // create data band
                    FastReport.DataBand data = new FastReport.DataBand();
                    data.Name = "Data1";
                    data.Height = Units.Millimeters * 10;
                    data.CanBreak = true;
                    // add data band to the page
                    page.Bands.Add(data);
                    // add page to the report
                    report.Pages.Add(page);

                    // ... create a matrix
                    // FastReport.Matrix.MatrixObject 
                    matrix = new FastReport.Matrix.MatrixObject();
                    matrix.Name = "Table1";
                    matrix.AutoSize = true;
                    matrix.Parent = data;

                    // ... create a column
                    FastReport.Matrix.MatrixHeaderDescriptor col1 = new FastReport.Matrix.MatrixHeaderDescriptor("", false);
                    col1.Sort = FastReport.SortOrder.None;
                    matrix.Data.Columns.Add(col1);

                    // ... create a row
                    FastReport.Matrix.MatrixHeaderDescriptor row1 = new FastReport.Matrix.MatrixHeaderDescriptor("", false);
                    row1.Sort = FastReport.SortOrder.None;
                    matrix.Data.Rows.Add(row1);

                    // ... create a cell
                    FastReport.Matrix.MatrixCellDescriptor cell1 = new FastReport.Matrix.MatrixCellDescriptor("",
                        FastReport.Matrix.MatrixAggregateFunction.None);
                    matrix.Data.Cells.Add(cell1);

                    // ... perform manual building of the matrix 
                    // ... using this event we can put our data into matrix as we want to
                    matrix.ManualBuild += Matrix_ManualBuild;
                    // ... define a MadifyResult event
                    // ... via this event we'll bring to rise other event - AlterCalcBounds
                    // ... of the result table of our matrix
                    matrix.ModifyResult += Matrix_ModifyResults;

                    matrix.BuildTemplate();

                    if (PrinterRegKey == "")
                    {
                        report.PrintSettings.Printer = this.GetDefaultPrinter(); 
                    }
                    else
                    {
                        string aPrinter = "";
                        byte[] aPrinterSettings = null;
                        GetPrinterSettingsFromRegistry(PrinterRegKey, PrinterRegSection, PrintRegItem, out aPrinter, out aPrinterSettings); 
                        if (aPrinter == "")
                        {
                            report.PrintSettings.Printer = this.GetDefaultPrinter();
                        }
                        else
                        {
                            PrinterSettings settings = SetPrinterMode(aPrinter, aPrinterSettings);
                            report.PrintSettings.Printer = aPrinter;
                            page.Landscape = settings.DefaultPageSettings.Landscape;
                            page.RawPaperSize = settings.DefaultPageSettings.PaperSize.RawKind;
                        }
                    }

                    report.Prepare(false);

                    report.ShowPrepared(false);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Grid show error = " + ex.ToString());
                }


            }
            finally
            {
                if (fields != null)
                    fields.Clear();

                if (fb_reader != null)
                    fb_reader.Close();

                if (fb_com != null)
                    fb_com.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();

            }

        }

        /*
         * We try to recognize date and time value via the fields' names
         * If a field name contains 'TIME' or 'DATE' substring, we're attempting     
         * to format the value
         * */
        public object FieldValueConverter(string fieldname, object value)
        {
            object result = value;
            if (fieldname.Contains("DATA") || fieldname.Contains("data") || fieldname.Contains("date") || fieldname.Contains("DATE"))
            {
                result = Convert.ToDateTime(value.ToString()).ToString("dd.MM.yyyy");
            }
            else if (fieldname.Contains("TIME") || fieldname.Contains("time"))
                result = Convert.ToDateTime(value.ToString()).ToString("HH:mm:ss");


            return result;
        }

        /*
         * Event handler for manual building of the matrix
         * */
        public void Matrix_ManualBuild(object sender, EventArgs e)
        {
            // throw new NotImplementedException();

            int j = 0;

            while (fb_reader.Read())
            {
                j++;

                for (int i = 0; i < fields.Count; i++)
                {
                    string fieldname = fields[i];
                    object value = fb_reader[fields[i]];

                    value = FieldValueConverter(fieldname, value);

                    ((FastReport.Matrix.MatrixObject)sender).Data.AddValue(
                        new object[] { fieldCaptions[i] },
                        new object[] { j.ToString() },
                        new object[] { value }
                        );

                }
            }
        }

        /*
         * Event on modify results in matrix
         * */

        public void Matrix_ModifyResults(object sender, EventArgs e)
        {
            // throw new NotImplementedException();
            ((FastReport.Matrix.MatrixObject)sender).ResultTable.AfterCalcBounds += Matrix_AfterCalcBounds;
        }

        private bool IsDate(string value)
        {
            Regex r = new Regex(@"^\d{2}\.\d{2}\.\d{4}$", RegexOptions.IgnoreCase);
            return r.Match(value).Success;
        }

        private bool IsTime(string value)
        {
            Regex r = new Regex(@"^\d{2}:\d{2}:\d{2}$", RegexOptions.IgnoreCase);
            return r.Match(value).Success;
        }

        private bool IsFigure(string value)
        {
            Regex r = new Regex(@"^\d+(\.|\,)?\d*\D{0}$", RegexOptions.IgnoreCase);
            return r.Match(value).Success;
        }



        private HorzAlign DetectCellAlign(object value)
        {
            HorzAlign result = HorzAlign.Left;

            if (IsFigure(value.ToString()))
                result = HorzAlign.Right;
            else if (IsDate(value.ToString()) || IsTime(value.ToString()))
                result = HorzAlign.Center;
                
            return result;
        }


        /*
         * Event to set table sizes
         * */
        public void Matrix_AfterCalcBounds(object sender, EventArgs e)
        {
            // throw new NotImplementedException();
            TableResult resultTable = sender as TableResult;

            // configurate table cells
            
            for (int i = 0; i < resultTable.ColumnCount; i++)
            {
                for (int j = 0; j < resultTable.RowCount; j++)
                {
                    resultTable[i, j].Font = new System.Drawing.Font(FontName, FontSize);
                    // Actually I'd like to calculate the align acording the cell's value
                    // but the result table is shown using the align of the very last cell 
                    // The function DetectCellAlign WORKS! but its using is useless
                    resultTable[i, j].HorzAlign = HorzAlign.Left; //  DetectCellAlign(resultTable[i, j].Text);
                    resultTable[i, j].VertAlign = VertAlign.Top;
                    resultTable[i, j].Border.Color = System.Drawing.Color.FromArgb(0xAA, 0xAA, 0xAA);
                    resultTable[i, j].Border.Style = LineStyle.Solid;
                    resultTable[i, j].Border.Lines = BorderLines.All;
                }
            }

            // configurate table columns
            foreach (TableColumn column in resultTable.Columns)
            {
                if (column.Index > 0)
                {
                    column.AutoSize = false;
                    column.Width = (float)Convert.ToDouble(fieldWidths[column.Index - 1]) + 6; // * FontSize;
                }
            }


            // configurate table rows
            foreach(TableRow row in resultTable.Rows)
            {
                row.AutoSize = true;
            }

            // recalc table sizes
            resultTable.CalcWidth();
            resultTable.CalcHeight();
        }

        /*
         * The function for building a report, which contains a tree og the given agenda
         * */
        public void BuildTreeReport(DateTime aDate)
        {
            // Preparing of the database connection and transaction
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand fb_com = new FbCommand("select * from AgendaTree(1,'"+aDate.ToString("dd.MM.yyyy")+"')",
                                              fb_con, fb_trans);
            FbDataReader  fb_reader = fb_com.ExecuteReader();

            List<Agenda> agenda = new List<Agenda>();

            try
            {
                try
                {
                    // Create a list of objects
                    agenda.Clear();
                    while (fb_reader.Read())
                    {
                        agenda.Add(new Agenda((int)fb_reader["id"], (int)fb_reader["aitem"], (string)fb_reader["comment"],
                            (string)fb_reader["fulltext"], (string)fb_reader["number"], (string)fb_reader["info"]));
                    }

                    report.Clear();
                    // Make a report
                    // Register the list above as a data source https://www.fast-report.com/en/forum/index.php?showtopic=5584
                    // report.RegisterData(agenda, "Agenda", BOConverterFlags.AllowFields, 1);
                    report.Dictionary.RegisterBusinessObject(agenda, "Agenda", 1, true);
                    // create report page
                    FastReport.ReportPage page = new FastReport.ReportPage();
                    page.Name = "Page1";
                    page.PaperWidth = 210;
                    page.PaperHeight = 297;
                    
                    // create report title
                    page.ReportTitle = new FastReport.ReportTitleBand();
                    page.ReportTitle.Name = "ReportTitle1";
                    page.ReportTitle.Height = Units.Millimeters * 10;
                    page.ReportTitle.CanGrow = true;
                    // show title
                    TextObject text = new TextObject();
                    text.Name = "Text1";
                    text.Dock = DockStyle.Fill;
                    text.Text = "Порядок работы от " + aDate.ToString("dd.MM.yyyy");
                    text.CanGrow = true;
                    page.ReportTitle.Objects.Add(text);
                    //  Create page header
                    page.PageHeader = new FastReport.PageHeaderBand();
                    page.PageHeader.Name = "PageHeader1";
                    page.PageHeader.Height = Units.Millimeters * 10;
                    page.PageHeader.CanGrow = false;
                    // Page number show
                    TextObject text_page = new TextObject();
                    text_page.Name = "Text2";
                    text_page.Dock = DockStyle.Fill;
                    text_page.Text = "[Page#]";
                    text_page.CanGrow = false;
                    text_page.HorzAlign = HorzAlign.Center;
                    page.PageHeader.Objects.Add(text_page);
                    // create data band
                    DataBand data = new DataBand();
                    data.Name = "Data1";
                    data.DataSource = report.GetDataSource("Agenda");
                    data.CanGrow = true;
                    data.StartNewPage = false;
                    page.Bands.Add(data);

                    // create textblock for the data band
                    // ... to show the agenda item's number
                    TextObject text_num = new TextObject();
                    text_num.Parent = data;
                    text_num.Name = "Text_Num";
                    text_num.Bounds = new System.Drawing.RectangleF(0, 0, Units.Centimeters * 2, Units.Centimeters * 1);
                    text_num.CanGrow = true;
                    text_num.CanShrink = true;
                    text_num.Text = "[Agenda.Number]";
                    // ... set event handler to change font style during report building
                    text_num.BeforePrint += Text_BeforePrint;

                    // create textblock for the data band
                    // ... to show the agenda item's text
                    TextObject text_text = new TextObject();
                    text_text.Parent = data;
                    text_text.Name = "Text_Text";
                    text_text.CanGrow = true;
                    text_text.CanShrink = true;
                    // ... calculate bounds of the text object as a difference between the page's width,
                    // ... its margins and the width of the privious textobjext
                    text_text.Bounds = new System.Drawing.RectangleF(Units.Centimeters * 2, 0,
                         Units.Millimeters * page.PaperWidth - Units.Millimeters * page.LeftMargin - Units.Millimeters * page.RightMargin - Units.Millimeters * 20, 
                         Units.Centimeters * 1);
                    text_text.Text = "[Agenda.FullText]";
                    // ... set event handler to change font style during report building
                    text_text.BeforePrint += Text_BeforePrint;
                    // add page to the report

                    report.Pages.Add(page);

                    report.Prepare();
                    report.ShowPrepared(false);
                    // report.Design();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }

            }
            finally
            {
                if (fb_reader != null)
                    fb_reader.Close();

                if (fb_com != null)
                    fb_com.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                {
                    fb_con.Close();
                    fb_con.Dispose();
                }

                if (agenda != null)
                {
                    agenda.Clear();
                }

            }
        }

        public void Text_BeforePrint(object sender, EventArgs e)
        {
            if ((Int32)(report.GetColumnValue("Agenda.Id")) == (Int32)(report.GetColumnValue("Agenda.AItem")))
            {
                ((TextObject)sender).Font = new System.Drawing.Font(((TextObject)sender).Font, System.Drawing.FontStyle.Bold);
            }
            else
            {
                ((TextObject)sender).Font = new System.Drawing.Font(((TextObject)sender).Font, System.Drawing.FontStyle.Regular);
            }
        }

        /*
         * The method to show a designer of the report
         * 
         * */
         public void DesignReport(int IdForm)
        {
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand SelectReport = null;
            FbCommand SubReports = null;
            FbDataReader readerForm = null;

            List<int> SubRportIds = new List<int>();


            report.Clear();

            try
            {
                try
                {
                    SelectReport = new FbCommand("select name, form from reportforms where id = " + Convert.ToString(IdForm), fb_con, fb_trans);

                    readerForm = SelectReport.ExecuteReader();

                    if (readerForm.Read())
                    {
                        byte[] reportForm = null;
                        string reportStr = "";
                        try
                        {
                            reportForm = (byte[])readerForm["form"];
                            reportStr = System.Text.Encoding.UTF8.GetString(reportForm);
                        }
                        catch
                        {
                        }

                        if (reportStr != "")
                        {
                            report.LoadFromString(reportStr);
                        }
                        else
                        {
                            report.Dictionary.Connections.Add(new FastReport.Data.FirebirdDataConnection());
                        }
                        report.Dictionary.Connections[0].ConnectionString = ConnectionString;
                        report.Dictionary.Connections[0].Name = "VoteDB";
                        report.Dictionary.Connections[0].Enabled = true;
                        report.SetParameterValue("IDRES", 0);
                        report.SetParameterValue("REQID", 0);
                        report.SetParameterValue("FRSELECT", false);
                        report.SetParameterValue("NEEDCLEAR", true);
                        
                        // fb_reader["name"].ToString() + ".frx";

                        fb_trans.Commit();
                        Config.DesignerSettings.CustomSaveDialog += CustomSaveDialog_Handler;
                        Config.DesignerSettings.CustomSaveReport += CustomSaveReport_Handler;
                        Config.DesignerSettings.DesignerClosed += DesignerClosed_Handler;
                        Config.DesignerSettings.DesignerLoaded += DesignerLoaded_Handler;

                        if (report.Design(true)) // Open designer in modal mode
                        {
                            // when designer is closed - save a report
                            SaveReportForm(IdForm, report.SaveToString());
                        }
                    }
                    else
                    {
                        fb_trans.Commit();
                        MessageBox.Show(IdForm.ToString() + " doesn't exist!");
                    }

                }
                catch (Exception ex)
                {
                    fb_trans.Rollback();
                    MessageBox.Show("Error during the report designer opening " + ex.ToString());
                }
            }
            finally
            {
                if (readerForm != null)
                    readerForm.Close();

                if (SelectReport != null)
                    SelectReport.Dispose();

                if (SubReports != null)
                    SubReports.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();

                if (SubRportIds != null)
                {
                    SubRportIds.Clear();
                }
            }
        }

        /*
         *  When we try to call a save dialog within the report's designer
         *  we call this method instead of
         *  
         * */
        private void CustomSaveDialog_Handler(object sender, OpenSaveDialogEventArgs e)
        {
            e.Cancel = false;
        }

        /*
         * When we try to save a report after call and confirm a save report dialog
         * we call this method
         * 
         * */
        private void CustomSaveReport_Handler(object sender, OpenSaveReportEventArgs e)
        {
            // save the report to the given e.FileName
            
        }

        /*
         * We handle the designer closed event to prevent the prompt dialog to show
         * */

        private void DesignerClosed_Handler(object sender, EventArgs e)
        {

        }

        /*
         * We handle the designer closed event to prevent the prompt dialog to show
         * */

        private void DesignerLoaded_Handler(object sender, EventArgs e)
        {

        }


        private void SaveReportForm(int IdReport, String ReportForm)
        {
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State != ConnectionState.Open)
                fb_con.Open();

            // Create a new transaction
            FbTransaction fb_trans = fb_con.BeginTransaction();

            // Create SELECT command to get a BLOB field FORM
            String sql = @"UPDATE REPORTFORMS SET " +
                 "FORM = @PAR_FORM " +
                 " WHERE ID = " + IdReport.ToString();

            FbCommand fb_com = new FbCommand(sql, fb_con, fb_trans);
            try
            {
                try
                {
                    // Convert the report's string into the array of bytes
                    byte[] btForm = System.Text.Encoding.UTF8.GetBytes(ReportForm);

                    // Create a new parameter for the query
                    FbParameter fb_par_form = new FbParameter("@PAR_FORM", FbDbType.Binary);
                    // Set the parameter's value
                    fb_par_form.Value = btForm;

                    // Add the parameter into the query
                    fb_com.Parameters.Add(fb_par_form);

                    // Execute the query
                    fb_com.ExecuteNonQuery();
                    fb_trans.Commit();
                }
                catch (Exception e)
                {
                    fb_trans.Rollback();
                    MessageBox.Show("Сохранение формы не удалось! " + e.ToString());
                }

            }
            finally
            {
                if (fb_com != null)
                    fb_com.Dispose();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();


            }

        }



    } // end of the class ReportBuilder


}
