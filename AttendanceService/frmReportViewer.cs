using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Telerik.WinControls.UI;

namespace AttendanceService.Properties
{
    public partial class frmReportViewer : RadForm
    {
        #region Variable
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        public int ReportCode = 0;
        public string PeriodCode = "173";
        public string EmpCode = "14";
        public string FromDate = "";
        public string ToDate = "";
        ReportDocument oDocument;
        #endregion

        #region Function

        #endregion

        #region Events
        public frmReportViewer()
        {
            InitializeComponent();
        }
        private void frmReportViewer_Load(object sender, EventArgs e)
        {
            try
            {
                this.TopMost = true;
                this.WindowState = FormWindowState.Maximized;
                oDocument = new ReportDocument();
                if(ReportCode == 1)
                {
                    string filepath = Path.Combine(Environment.CurrentDirectory.ToString(),"Report", "SaveandPostedAttendanceReport.rpt");
                    oDocument.Load(filepath);
                    oDocument.SetDatabaseLogon(Settings.Default.DBUser, Settings.Default.DBPassword, Settings.Default.ServerName, Settings.Default.Database);
                    oDocument.SetParameterValue("Critaria", $"WHERE  A1.PeriodID={PeriodCode} and A2.EmpID='{EmpCode}'");
                    rptViewer.ReportSource = oDocument;
                }
                else
                {
                    string filepath = Path.Combine(Environment.CurrentDirectory.ToString(), "Report", "TempAttendanceReport.rpt");
                    oDocument.Load(filepath);
                    oDocument.SetDatabaseLogon(Settings.Default.DBUser, Settings.Default.DBPassword, Settings.Default.ServerName, Settings.Default.Database);
                    oDocument.SetParameterValue("Critaria", $"WHERE A1.EmpID = '{EmpCode}' AND A4.PunchedDate BETWEEN '{FromDate}' AND '{ToDate}'");
                    rptViewer.ReportSource = oDocument;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                this.Close();
            }
        }
        #endregion
    }
}
