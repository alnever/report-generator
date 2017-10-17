using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;


namespace ReportsLib
{

    public class Reports
    {
        /*
         * Create a connection string for the Firebird database using parameters obtained from the interface function
         * 
         * 
         * */
        private static String MakeConnectionString(String aHost, String aDatabase, String aLogin, String aPassword)
        {
            FbConnectionStringBuilder connectionString = new FbConnectionStringBuilder();
            connectionString.Charset = "WIN1251";
            connectionString.UserID = aLogin;
            connectionString.Password = aPassword;
            connectionString.Database = aDatabase;
            connectionString.DataSource = aHost;
            connectionString.ServerType = 0;

            return connectionString.ConnectionString;
        }

        /*
         * This function creates a connection string and passes it with some result parameters to the form,
         * which shows a list of aviable reports for given result's type and its subtype
         * 
         * Creation of the report is implemented within the ReportSelectForm class
         * 
         * */
        private static void GetReport(String aHost, String aDatabase, String aLogin, String aPassword, int idRes, int Rtype, int RSubtype, int ReqId, bool ShowPrint = true)
        {
            String ConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
            ReportSelectForm rs_form = new ReportSelectForm(ConnectionString, idRes, Rtype, RSubtype, ReqId, ShowPrint);
            rs_form.ShowDialog();
            rs_form.Dispose();
            rs_form = null;
        }

        /*
         * Reloading of the GetReport function to use within AutoPrintService 
         * for oper == 1: Export Into The File of the given Format
         * 
         * */
        private static void GetReportForExport(String aConnectionString, int idRes, int idForm,
                                            int idFormat, string aPath, string aFileName, DateTime aResDate,
                                            int ReqId)
        {
            ReportBuilder builder = new ReportBuilder(aConnectionString, idRes, ReqId, false, false, true);
            builder.ExportReportFile(idForm, idFormat, aPath, aFileName, aResDate);
        }

        /*
         * Save agregative report into the file
         * 
         * */

        private static void GetSumReportForExport(String aConnectionString, int idForm,
                                            int idFormat, string aPath, string aFileName, DateTime aResDate,
                                            int ReqId)
        {
            ReportBuilder builder = new ReportBuilder(aConnectionString, null, ReqId, false, false, true);
            builder.ExportSumReportFile(idForm, idFormat, aPath, aFileName, aResDate);
        }


        /*
         * Printing within the AutoPrintService
         * 
         * This variant of the function contains the name, the port, and setting for the printer,
         * defined within the automatical task
         * 
         * This variant is usefull for application and services, which were written using .NET 
         * 
         * */

        private static void GetReportForPrint(String aConnectionString, int idRes, int idForm,
                                                int idAuto,
                                                int ReqId)
        {
            Log.Write(idRes, " PRINT builder creation ");
            ReportBuilder builder = new ReportBuilder(aConnectionString, idRes, ReqId, false, false, true);
            Log.Write(idRes, " PRINT builder creatied ");
            builder.PrintReport(idForm, idAuto);
            Log.Write(idRes, " PRINT builder finished its work ");
        }

        /*
         * Automatical print of the aggregative report
         * */

        private static void GetSumReportForPrint(String aConnectionString, DateTime ResDate, int idForm, int idAuto,
                                                int ReqId)
        {
            ReportBuilder builder = new ReportBuilder(aConnectionString, null, ReqId, false, false, true);
            builder.PrintSumReport(idForm, idAuto, ResDate);
        }


        /*
         * Sending converted report's file to the given FTP-sever
         * 
         * */
        private static void GetReportForFTP(String aConnectionString, int idRes, int idForm,
                                            int idFormat, string aFileName, DateTime aResDate,
                                            string aServer, string aLogin, string aPassword, string aServerPath,
                                            int ReqId)
        {
            ReportBuilder builder = new ReportBuilder(aConnectionString, idRes, ReqId, false, false, true);
            builder.SendToFtp(idForm, idFormat, aFileName, aResDate,
                aServer, aLogin, aPassword, aServerPath);
        }

        /*
        * Sending agregative report to the given FTP-sever
        * 
        * */
        private static void GetSumReportForFTP(String aConnectionString, int idForm,
                                            int idFormat, string aFileName, DateTime aResDate,
                                            string aServer, string aLogin, string aPassword, string aServerPath,
                                            int ReqId)
        {
            ReportBuilder builder = new ReportBuilder(aConnectionString, null, ReqId, false, false, true);
            builder.SendSumToFtp(idForm, idFormat, aFileName, aResDate,
                aServer, aLogin, aPassword, aServerPath);
        }


        /*
         * This function create a connection string and passes it with a list of the given results to the form,
         * which shows a list of reports families according the given types of the reports
         * 
         * Definition of the allowed families and creation of the reports is implemented in the FamilySelectForm
         * 
         * */
        private static void GetReportFamily(String aHost, String aDatabase, String aLogin, String aPassword, List<ResultInfo> Results, int aReqId)
        {
            String ConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
            FamilySelectForm fam_form = new FamilySelectForm(ConnectionString, Results, aReqId);
            fam_form.ShowDialog();
            fam_form.Dispose();
            fam_form = null;
        }
        
        /*
         * This function allows to select a type of the summaral report and build the corresponding report
         * for the given set of the results
         * 
         * */
        private static void GetSummaryReport(String aHost, String aDatabase, String aLogin, String aPassword, List<ResultInfo> Results, int aReqId)
        {
            String ConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
            SummaryReportSelect sum_form = new SummaryReportSelect(ConnectionString, Results, aReqId);
            sum_form.ShowDialog();
            sum_form.Dispose();
            sum_form = null;
        }

        /*
         * The interface function of the DLL, which supports to create a report and shows its preview
         * 
         * 
         * */

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static void ShowReport(String aHost, String aDatabase, String aLogin, String aPassword, int idRes, int Rtype, int RSubtype, int ReqId)
        {
            GetReport(aHost, aDatabase, aLogin, aPassword, idRes, Rtype, RSubtype, ReqId);
        }

        /*
         * The interface function of the DLL, which supports to create a report and allows to print it
         * without showing of the preview
         * 
         * 
         * */

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static void PrintReport(String aHost, String aDatabase, String aLogin, String aPassword, int idRes, int Rtype, int RSubtype, int ReqId)
        {
            GetReport(aHost, aDatabase, aLogin, aPassword, idRes, Rtype, RSubtype, ReqId, false);
        }


        /*
         * Overloadin the function PrintReport for the AutoPrint service
         * Insteed of passing the connection string, it uses database's parameters:
         * aHost, aDatabase, aLogin, aPassword, and then constructs a new 
         * connection string
         * 
         * In addition, this function takes an ID of the automatical task insteed of the 
         * printer's parameters
         * 
         * */

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static bool PrintReportAuto(String aHost, String aDatabase, String aLogin, String aPassword,
                                            int idRes, int idForm, int idAuto,
                                            int aCopies,
                                            int ReqId = -100)
        {
            bool result = false;
            try
            {
                String aConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                for (int i = 0; i < aCopies; i++)
                {
                    GetReportForPrint(aConnectionString, idRes, idForm, idAuto, ReqId);
                }
                result = true;
            }
            catch (Exception e)
            {
                Log.Write(idRes, "PRINT FORM = "+idForm.ToString() ,e);
            }
            return result;
        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static bool PrintSumReportAuto(String aHost, String aDatabase, String aLogin, String aPassword,
                                            int idForm, int idAuto, string aResDate,
                                            int aCopies,
                                            int ReqId = -100)
        {
            bool result = false;
            DateTime dResDate = Convert.ToDateTime(aResDate);
            try
            {
                String aConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                for (int i = 0; i < aCopies; i++)
                {
                    GetSumReportForPrint(aConnectionString, dResDate, idForm, idAuto, ReqId);
                }
                result = true;
            }
            catch (Exception e)
            {
                Log.Write(0,"PRINT FORM = " + idForm.ToString(), e);
            }
            return result;
        }


        /*
         * The function for export report into the file
         * For using by the AutoPrintService ONLY!!!
         * The test implementation: the ReqId == -100 for every instance of the service
         * */

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static bool ExportReport(String aHost, String aDatabase, String aLogin, String aPassword,
                                            int idRes, int idForm,
                                            int idFormat, string aPath, string aFileName, string aResDate,
                                            int aCopies,
                                            int ReqId = -100)
        {
            bool result = false;
            try
            {
                DateTime dResDate = Convert.ToDateTime(aResDate);
                String aConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                for (int i = 0; i < aCopies; i++)
                {
                    GetReportForExport(aConnectionString, idRes, idForm, idFormat, aPath, aFileName, dResDate, ReqId);
                }
                result = true;
            }
            catch(Exception e)
            {
                Log.Write(idRes, "EXPORT FORM = " + idForm.ToString() + " TO " +idFormat.ToString(), e);
            }
            return result;

        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static bool ExportSumReport(String aHost, String aDatabase, String aLogin, String aPassword,
                                            int idForm, int idFormat, 
                                            string aPath, string aFileName, string aResDate,
                                            int aCopies,
                                            int ReqId = -100)
        {
            bool result = false;
            try
            {
                DateTime dResDate = Convert.ToDateTime(aResDate);
                String aConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                for (int i = 0; i < aCopies; i++)
                {
                    GetSumReportForExport(aConnectionString, idForm, idFormat, aPath, aFileName, dResDate, ReqId);
                }
                result = true;
            }
            catch (Exception e)
            {
                Log.Write(0, "EXPORT FORM = " + idForm.ToString() + " TO " + idFormat.ToString(), e);
            }
            return result;

        }


        /*
         * The function for
         * ... export report to the given format
         * ... sending its file to the given FTP-server
         *          
         * The test implementation: the ReqId == -100 for every instance of the service
         * 
         * */

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static bool SendReport(String aHost, String aDatabase, String aLogin, String aPassword,
                                            int idRes, int idForm,
                                            int idFormat, string aFileName, string aResDate,
                                            string aServer, string aFtpLogin, string aFtpPassword, string aServerPath,
                                            int aCopies,
                                            int ReqId = -100)
        {
            bool result = false;
            try
            {
                String aConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                DateTime dResDate = Convert.ToDateTime(aResDate);
                for (int i = 0; i < aCopies; i++)
                {
                    GetReportForFTP(aConnectionString, idRes, idForm,
                        idFormat, aFileName, dResDate,
                        aServer, aFtpLogin, aFtpPassword, aServerPath,
                        ReqId);
                }
                result = true;
            } catch(Exception e)
            {
                Log.Write(idRes, "SEND TO FTP FORM = " + idForm.ToString() + " FORMAT " + idFormat.ToString(), e);
            }
            return result;
        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static bool SendSumReport(String aHost, String aDatabase, String aLogin, String aPassword,
                                            int idForm,
                                            int idFormat, string aFileName, string aResDate,
                                            string aServer, string aFtpLogin, string aFtpPassword, string aServerPath,
                                            int aCopies,
                                            int ReqId = -100)
        {
            bool result = false;
            try
            {
                String aConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                DateTime dResDate = Convert.ToDateTime(aResDate);
                for (int i = 0; i < aCopies; i++)
                {
                    GetSumReportForFTP(aConnectionString, idForm,
                        idFormat, aFileName, dResDate,
                        aServer, aFtpLogin, aFtpPassword, aServerPath,
                        ReqId);
                }
                result = true;
            }
            catch (Exception e)
            {
                Log.Write(0, "SEND TO FTP FORM = " + idForm.ToString() + " FORMAT " + idFormat.ToString(), e);
            }
            return result;
        }




        /*
         * The interface function of the DLL, which creates a list of the results for further creation of the 
         * reports families list according the given reports types.
         * 
         * The very first four parameters of the function are used to create a connection string.
         * Parameter n is a number of the results selected by the user in the main application.
         * Parameters idRes, Rtype and RSubtype are pointers at the arrays of the results IDs, its types, and  subtypes.
         * Factial parameters (in Delphi applications) have the type of ^integer
         * 
         * 
         * */

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        unsafe public static void PrintReports(String aHost, String aDatabase, String aLogin, String aPassword, 
            int n, Int32* idRes, Int32* Rtype, Int32* RSubtype, int ReqId)
        {
            List<ResultInfo> results = new List<ResultInfo>();

            for (int i = 0; i < n; i++)
            {
                results.Add(new ResultInfo((Int32)idRes[i], (Int32)Rtype[i], (Int32)RSubtype[i]));
            }

            GetReportFamily(aHost, aDatabase, aLogin, aPassword, results, ReqId);
        }

        /*
         * The function is used for printing of the aggregative reports
         * 
         * */

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        unsafe public static void SummaryReport(String aHost, String aDatabase, String aLogin, String aPassword, 
            int n, Int32* idRes, Int32* Rtype, Int32* RSubtype, int ReqId)
        {
            List<ResultInfo> results = new List<ResultInfo>();
            for (int i = 0; i < n; i++)
            {
                results.Add(new ResultInfo((Int32)idRes[i], (Int32)Rtype[i], (Int32)RSubtype[i]));
            }

            GetSummaryReport(aHost, aDatabase, aLogin, aPassword, results, ReqId);
        }

        /*
         * The function is used for printing of the aggregative reports
         * @param aHost - the database server name
         * @param aDatabase - the database file name
         * @paran aLogin - the login of the database user
         * @param aPassword - the database user's password
         * @param ResCount - a number of the results in the source query (within the application, which called
         *                   this function)
         * @param idRes - the array of the ID's of the results from the source query
         * @param ColCount - the number of the columns within th grid from the calling application
         * @param Fields - the sublimated list of the fields within the grid from the calling application
         * @param OrderBy - the order by clause from the source query
         * 
         * REMARK. It looks like it's impossible to pass a parameter of the type string*. Because of this, 
         * the parameter Fields contains the list of filed names, which are ...
         * ... displayed within the application's grid
         * ... sorted according the order of their presenting via the grif
         * ... devided using semicolon
         * 
         * During passing the parameter into further functions we should ...
         * ... split the source string by the semicolon
         * ... convert the obtained array of strings into the list of strings
         * 
         * */
        
        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static void ShowGrid(string aHost, string aDatabase, string aLogin, string aPassword,
            string Query,
            string Parameters,
            string ParamValues,
            string Fields,
            string FieldCaptions,
            string FieldWidths,
            string FontName,
            int FontSize,
            string ReportCaption,
            string PrinterRegKey,
            string PrinterRegSection,
            int PrinterRegItem
            )
        {
            String ConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
            ReportBuilder builder = new ReportBuilder(ConnectionString, null, 0, true, false, false);
            builder.ShowGrid(Query, Parameters, ParamValues, 
                             Fields, FieldCaptions, FieldWidths, 
                             FontName, FontSize, 
                             ReportCaption, 
                             PrinterRegKey, PrinterRegSection, PrinterRegItem);
        }

        /*
         * The function for agenda tree report creation
         * 
         * */
        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static void ShowTree(string aHost, string aDatabase, string aLogin, string aPassword, string aDate)
        {
            try
            {
                String ConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                ReportBuilder builder = new ReportBuilder(ConnectionString, null, 0, true, false, false);
                builder.BuildTreeReport(Convert.ToDateTime(aDate));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }


        /*
         * The function opens a report's designer
         * 
         * */
        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static void DesignReport(string aHost, string aDatabase, string aLogin, string aPassword, int idForm)
        {
            try
            {
                String ConnectionString = MakeConnectionString(aHost, aDatabase, aLogin, aPassword);
                ReportBuilder builder = new ReportBuilder(ConnectionString, null, 0, true, false, false);
                builder.DesignReport(idForm);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
    }
}
