using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportsLib
{
    public partial class FamilySelectForm : Form
    {
        String ConnectionString = "";
        List<ResultInfo> results = null;
        List<ReportFamily> families = null;
        int ReqId;

        public FamilySelectForm()
        {
            InitializeComponent();
        }

        public FamilySelectForm(string aConnectionString, List<ResultInfo> aResults, int aReqId)
        {
            InitializeComponent();
            ConnectionString = aConnectionString;
            results = aResults;
            families = new List<ReportFamily>();
            ReqId = aReqId;
        }

        /*
         * We need to get an index of the report's family from the families list
         * to calculate the number of the families, which fit the given result
         * 
         * */
        private int GetFamilyIndex(List<ReportFamily> aFamilies, int aFamId)
        {
            int res = -1;
            for(int i = 0; i < aFamilies.Count && res == -1; i++)
            {
                if (aFamilies[i].ID == aFamId)
                    res = i;
            }
            return res;
        }

        private void FamilySelectForm_Shown(object sender, EventArgs e)
        {

            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand SelectFamilies = new FbCommand();

            FbDataReader readerFamilies = null;

            try
            {
                try
                {
                    /* Select full list of the reports families */

                    SelectFamilies.CommandText = "select id, pos, name, displayname, visible, frselect from reportfamily where visible = 1 order by pos";
                    SelectFamilies.Connection = fb_con;
                    SelectFamilies.Transaction = fb_trans;

                    readerFamilies = SelectFamilies.ExecuteReader();

                    // ... Create a list of the families
                    while (readerFamilies.Read())
                    {
                        families.Add(new ReportFamily((int)readerFamilies["id"], (int)readerFamilies["pos"], (string)readerFamilies["name"],
                            (string)readerFamilies["displayname"], (int)readerFamilies["visible"], (int)readerFamilies["frselect"],0));
                    }

                    readerFamilies.Close();

                    /* Select families for every report  */
                    SelectFamilies.CommandText = "select famid from get_family_result(@idr)";
                    SelectFamilies.Connection = fb_con;
                    SelectFamilies.Transaction = fb_trans;
                    
                    // ... create new parameter
                    FbParameter par_resid = new FbParameter("@idr", FbDbType.Integer);

                    // ... for every result, which was passed into the form
                    for (int i=0; i<results.Count;i++)
                    {
                        // ... clear the parameters set of the query
                        SelectFamilies.Parameters.Clear();
                        // ... set new value of the parameter
                        par_resid.Value = results[i].IdRes;
                        // ... add current parameter to the query
                        SelectFamilies.Parameters.Add(par_resid);

                        readerFamilies = SelectFamilies.ExecuteReader();

                        // ... calculate numver of the families, we'd been able to build for given result
                        while (readerFamilies.Read())
                        {
                            int fam_id = GetFamilyIndex(families, (int)readerFamilies[0]);
                            if (fam_id > -1)
                            {
                                families[fam_id].RepCount++;
                            }
                        }

                        readerFamilies.Close();
                    }

                    fb_trans.Commit();

                    // Before showing of the families list, we need to clear the families list
                    // and delete families with RepCount == 0

                    for (int i = 0; i<families.Count;)
                    {
                        if (families[i].RepCount == 0)
                        {
                            families.RemoveAt(i); // if RepCount == 0, then we remove the Family
                        }
                        else
                        {
                            i++; // else we should take the next element of the families list
                        }
                    }

                    // And now we can show selected families
                    this.groupBox1.AutoSize = true;
                    this.groupBox1.Controls.Clear();

                    for (int i = 0; i < families.Count; i++)
                    {
                        RadioButton rb = new RadioButton();
                        rb.Text = families[i].DispName;
                        rb.Name = "RadioButton" + i.ToString();
                        rb.Location = new Point(5, 30 * (i + 1));
                        rb.AutoSize = true;
                        // checkedListBox1.Items.Add(readerReports.GetString(1), false);
                        groupBox1.Controls.Add(rb);
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Info = " + ex.ToString());
                    fb_trans.Rollback();
                }

            }
            finally
            {
                if (readerFamilies != null)
                    readerFamilies.Close();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (SelectFamilies != null)
                    SelectFamilies.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

            }

        }

        /*
         * Get actual Family ID from the families list according the checked element of the CheckedListBox
         * It takes the very 1st checked element. If there are more than one element checked, so other 
         * elements will be ignored 
         * 
         * */

        private int GetFamilyd()
        {
            int idFam = 0;

            for (int i = 0; i < groupBox1.Controls.Count && idFam == 0; i++)
            {
                if (((RadioButton)(groupBox1.Controls[i])).Checked)
                {
                    idFam = families[i].ID;
                }
            }

            

            return idFam;
        }

        private int GetFamilyIndex(int aIdFam)
        {
            int i;
            for (i = 0; i < families.Count && families[i].ID != aIdFam; i++) ;
            return i;
        }

        /*
         * Print reports selected via selection of the reports family for every result
         * 
         * */

        private void PrintFamily()
        {
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand SelectReports = new FbCommand();

            FbDataReader readerReports = null;

            int IdFam = GetFamilyd();
            int FamIndex = GetFamilyIndex(IdFam);
            int IdReport = 0;

            ReportBuilder rep_builder = null;

            try
            {
                try
                {
                    // Prepare the query, receiving the report's ID
                    SelectReports.CommandText = "select first 1 id from get_family_result_report(@idr, @famid)";
                    SelectReports.Connection = fb_con;
                    SelectReports.Transaction = fb_trans;
                    // ... and its parameters

                    FbParameter par_res = new FbParameter("@idr", FbDbType.Integer);

                    FbParameter par_fam = new FbParameter("@famid", FbDbType.Integer);
                    par_fam.Value = IdFam;

                    // For every result we'll print the report, that is corresponding the selected family
                    for (int i = 0; i < results.Count; i++)
                    {
                        // ... prepare query's parameters
                        SelectReports.Parameters.Clear();

                        par_res.Value = results[i].IdRes;
                        SelectReports.Parameters.Add(par_res);
                        SelectReports.Parameters.Add(par_fam);

                        // ... execute query and obtain the report's ID
                        readerReports = SelectReports.ExecuteReader();
                        if (readerReports.Read())
                        {
                            IdReport = (int)readerReports[0];
                            for (int copies = Convert.ToInt32(Copies.Value); copies > 0; copies--)
                            {
                                // Create a new report builder
                                // ConnectionString - parameters to connect to the database
                                // results[i].IdRes - the ID of the current result
                                // ReqId - the current value of the Requiered ID (ID of the session)
                                // ShowPrint == false (the 4th parameter) - in this case the builder will print the report, but don't show the report's preview
                                // ShowFrDialog (the 5th parameter) - the dialog for the selection of the fractions will be shown, if just ...
                                // ... the selected family requires to select the fractions and groups before the report's building
                                // ... this is the very 1st reports ought to be printed
                                // ... this is the very 1st cpoy of the report ought to be printed
                                // NeedClear (the 6th parameter) influes at the process of the database cleaning 
                                // (in part of the temporary data, which were used during the report creation)
                                // It'll be settled to the TRUE, if ...
                                // ... it is the very last report in the queue of the reports ought to be printed
                                // ... it is the last copy of the report (because of the parameters of the cycle the last copy will have the number 1)

                                rep_builder = new ReportBuilder(ConnectionString, results[i].IdRes, ReqId, false, 
                                    families[FamIndex].FrSelect == 1 && i == 0 && copies == Convert.ToInt32(Copies.Value), 
                                    i == results.Count - 1 && copies == 1);
                                rep_builder.BuildReport(IdReport);
                            }
                        }
                        readerReports.Close();

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Info = " + ex.ToString());
                    fb_trans.Rollback();
                }
            }
            finally
            {
                if (readerReports != null)
                    readerReports.Close();

                if (fb_trans != null)
                    fb_trans.Dispose();

                if (SelectReports != null)
                    SelectReports.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            PrintFamily();
            this.Close();                
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
