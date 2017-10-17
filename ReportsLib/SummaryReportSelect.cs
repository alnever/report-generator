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
    public partial class SummaryReportSelect : Form
    {
        String ConnectionString;
        int ReqId;
        List<ResultInfo> results;
        List<ReportsFit> ReportList;


        public SummaryReportSelect()
        {
            InitializeComponent();
        }

        public SummaryReportSelect(string aConnectionString, List<ResultInfo> aResults, int aReqId)
        {
            InitializeComponent();
            ConnectionString = aConnectionString;
            results = aResults;
            ReqId = aReqId;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            ReportBuilder rep_builder = new ReportBuilder(ConnectionString, results, ReqId, true, true, true);
            rep_builder.BuildSummaryReport(GetIdReport());
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SummaryReportSelect_Shown(object sender, EventArgs e)
        {
            
            
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand SelectReports = new FbCommand("select id, displayname from reportforms where mask = 2 and visible = 1 order by displayname",
                fb_con, fb_trans);

            FbDataReader readerReports = null;
            try
            {
                try
                {
                    readerReports = SelectReports.ExecuteReader();

                    if (ReportList == null)
                        ReportList = new List<ReportsFit>();
                    ReportList.Clear();

                    this.groupBox1.AutoSize = true;
                    this.groupBox1.Controls.Clear();

                    int i = 0;
                    while (readerReports.Read())
                    {
                        RadioButton rb = new RadioButton();
                        rb.Text = readerReports.GetString(1);
                        rb.Name = "RadioButton" + i.ToString();
                        rb.Location = new Point(5, 30 * (i + 1));
                        rb.AutoSize = true;
                        // checkedListBox1.Items.Add(readerReports.GetString(1), false);
                        groupBox1.Controls.Add(rb);

                        ReportList.Add(new ReportsFit((int)readerReports[0]));
                        i++;
                    }

                }
                catch (Exception ex)
                {
                    fb_trans.Rollback();
                    MessageBox.Show("Error info (summary) = " + ex.ToString());
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

        private int GetIdReport()
        {
            int idReport = 0;

            for (int i = 0; i < groupBox1.Controls.Count && idReport == 0; i++)
            {
                if (((RadioButton)(groupBox1.Controls[i])).Checked)
                {
                    idReport = ReportList[i].IdReport;
                }
            }

            return idReport;
        }

    }
}
