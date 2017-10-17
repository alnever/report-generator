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
    public partial class ReportSelectForm : Form
    {
        int Rtype = 0;
        int Rsubtype = 0;
        int IdRes = 0;
        int ReqId = 0;
        string ConnectionString = "";
        List<ReportsFit> ReportList = null;
        bool ShowPrint = false;

        public ReportSelectForm()
        {
            InitializeComponent();
        }

        public ReportSelectForm(String aConnectionString, int aIdRes, int aRtype, int aRsubtype, int aReqId, bool aShowPrint)
        {
            InitializeComponent();
            ConnectionString = aConnectionString;
            IdRes = aIdRes;
            Rtype = aRtype;
            Rsubtype = aRsubtype;
            ReqId = aReqId;
            ShowPrint = aShowPrint;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            ReportBuilder rep_builder = new ReportBuilder(ConnectionString, IdRes, ReqId, ShowPrint, true, true);
            rep_builder.BuildReport(GetIdReport());
            this.Close();
        }

        private void ReportSelectForm_Shown(object sender, EventArgs e)
        {
            FbConnection fb_con = new FbConnection(ConnectionString);
            if (fb_con.State == System.Data.ConnectionState.Closed)
            {
                fb_con.Open();
            }

            FbTransaction fb_trans = null;
            fb_trans = fb_con.BeginTransaction();

            FbCommand SelectReports = null;

            if (Rtype < 10)
            {
                SelectReports = new FbCommand("select id, displayname from reportforms where visible = 1 and rtype = " + Convert.ToString(Rtype) +
                    " and (rsubtype = " + Convert.ToString(Rsubtype) +
                    " or rsubtype is null) "+
                    " order by displayname", fb_con, fb_trans);
            }
            else
            {
                SelectReports = new FbCommand("select id, displayname from reportforms where visible = 1 and rtype = " + 
                    Convert.ToString(Rtype) +
                    " order by displayname", fb_con, fb_trans);
            }

            FbDataReader readerReports = null;

            try
            {

                this.groupBox1.AutoSize = true;
                this.groupBox1.Controls.Clear();

                if (ReportList == null)
                    ReportList = new List<ReportsFit>();
                ReportList.Clear();

                readerReports = SelectReports.ExecuteReader();

                int i = 0;
                while (readerReports.Read())
                {
                    RadioButton rb = new RadioButton();
                    rb.Text = readerReports.GetString(1);
                    rb.Name = "RadioButton" + i.ToString();
                    rb.Location = new Point(5, 30 * (i+1));
                    rb.AutoSize = true;
                    // checkedListBox1.Items.Add(readerReports.GetString(1), false);
                    groupBox1.Controls.Add(rb);

                    ReportList.Add(new ReportsFit(readerReports.GetInt32(0)));
                    i++;
                }
                
            }
            finally
            {
                if (readerReports != null)
                    readerReports.Close();

                if (fb_trans != null)
                    fb_trans.Commit();

                if (SelectReports != null)
                    SelectReports.Dispose();

                if (fb_con != null && fb_con.State == System.Data.ConnectionState.Open)
                    fb_con.Close();

                if (fb_con != null)
                    fb_con.Dispose();

            }
        }

        private int GetIdReport()
        {
            int idReport = 0;

/*            for(int i=0; i < checkedListBox1.Items.Count && idReport == 0; i++)
            {
                if (checkedListBox1.CheckedItems.IndexOf(checkedListBox1.Items[i]) >= 0)
                {
                    idReport = ReportList[i].IdReport;
                }
            }
            */
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
