using DIHRMS;
using NLog;
using NLog.Layouts;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Telerik.WinControls.UI;
using AttendanceService.Models;
using System.Collections.Specialized;
using System.Data.SqlTypes;
using static Telerik.WinControls.VistaAeroTheme;
using System.Windows.Forms;
using System.Reflection;
using System.Data.SqlClient;
using Telerik.WinControls;
using AttendanceService.Properties;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Reflection.Emit;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using UFFU;

namespace AttendanceService
{
    public partial class frmMain : RadForm
    {

        #region Variables
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private string ConnectionString = "", AttConnectionString = "";
        private DataTable dtEmployees;
        private DataTable dtProcessed;
        int Serial = 0;
        private bool GridOneToggle = false;
        #endregion

        #region Functions

        void CreateGrid()
        {
            try
            {
                dtEmployees = new DataTable();
                dtEmployees.Columns.Add("Select", typeof(bool));
                dtEmployees.Columns.Add("EmpCode", typeof(string));
                dtEmployees.Columns.Add("EmpName", typeof(string));
                dtEmployees.Columns.Add("Department", typeof(string));
                dtEmployees.Columns.Add("Designation", typeof(string));
                dtEmployees.Columns.Add("Location", typeof(string));
                dtEmployees.Columns.Add("Branch", typeof(string));
                dtEmployees.Columns.Add("Payroll", typeof(string));
                grdEmployee.DataSource = dtEmployees;

                dtProcessed = new DataTable();
                dtProcessed.Columns.Add("ID");
                dtProcessed.Columns.Add("Serial");
                dtProcessed.Columns.Add("EmpCode");
                dtProcessed.Columns.Add("EmpName");
                dtProcessed.Columns.Add("Date", typeof(DateTime));
                dtProcessed.Columns.Add("Day");
                dtProcessed.Columns.Add("Shift");
                dtProcessed.Columns.Add("ShiftIn");
                dtProcessed.Columns.Add("ShiftOut");
                dtProcessed.Columns.Add("ShiftDuration");
                dtProcessed.Columns.Add("In");
                dtProcessed.Columns.Add("Out");
                dtProcessed.Columns.Add("WorkHour");
                dtProcessed.Columns.Add("EarlyIn");
                dtProcessed.Columns.Add("LateIn");
                dtProcessed.Columns.Add("EarlyOut");
                dtProcessed.Columns.Add("LateOut");
                dtProcessed.Columns.Add("LeaveHour");
                dtProcessed.Columns.Add("LeaveType");
                dtProcessed.Columns.Add("LeaveNew");
                dtProcessed.Columns.Add("LeaveCount", typeof(decimal));
                dtProcessed.Columns.Add("LTID");
                dtProcessed.Columns.Add("OTHour");
                dtProcessed.Columns.Add("OTType");
                dtProcessed.Columns.Add("OTID");
                dtProcessed.Columns.Add("Status");
                grdProcess.DataSource = dtProcessed;

            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        void FillCombo()
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    var oDepartments = (from a in odb.MstDepartment
                                        where a.FlgActive == true
                                        orderby a.DeptName
                                        select a).ToList();
                    cmbDepartment.Items.Add("All");
                    foreach (var dept in oDepartments)
                    {
                        cmbDepartment.Items.Add(dept.DeptName);
                    }
                    cmbDepartment.SelectedIndex = 0;

                    var oDesignation = (from a in odb.MstDesignation
                                        where a.FlgActive == true
                                        orderby a.Name
                                        select a).ToList();
                    cmbDesignation.Items.Add("All");
                    foreach (var desig in oDesignation)
                    {
                        cmbDesignation.Items.Add(desig.Name);
                    }
                    cmbDesignation.SelectedIndex = 0;

                    var oLocation = (from a in odb.MstLocation
                                     where a.FlgActive == true
                                     orderby a.Name
                                     select a).ToList();
                    cmbLocation.Items.Add("All");
                    foreach (var loc in oLocation)
                    {
                        cmbLocation.Items.Add(loc.Name);
                    }
                    cmbLocation.SelectedIndex = 0;

                    var oBranches = (from a in odb.MstBranches
                                     where a.FlgActive == true
                                     orderby a.Name
                                     select a).ToList();
                    cmbBranch.Items.Add("All");
                    foreach (var branch in oBranches)
                    {
                        cmbBranch.Items.Add(branch.Name);
                    }
                    cmbBranch.SelectedIndex = 0;

                    var oPayrolls = (from a in odb.CfgPayrollDefination
                                     orderby a.PayrollName
                                     select a).ToList();
                    foreach (var payroll in oPayrolls)
                    {
                        cmbPayroll.Items.Add(payroll.PayrollName);
                    }
                    cmbPayroll.SelectedIndex = 0;


                    string SelectedPayroll = cmbPayroll.SelectedItem.ToString();
                    cmbPeriod.Items.Clear();
                    if (!string.IsNullOrEmpty(SelectedPayroll))
                    {
                        var oPeriods = (from a in odb.CfgPeriodDates
                                        where a.FlgLocked == false
                                        && a.CfgPayrollDefination.PayrollName == SelectedPayroll
                                        orderby a.StartDate
                                        select a).ToList();
                        foreach (var period in oPeriods)
                        {
                            cmbPeriod.Items.Add(period.PeriodName);
                        }
                    }
                    cmbPeriod.SelectedIndex = 0;

                }

            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        void FillEmployees()
        {
            try
            {
                dtEmployees.Rows.Clear();
                using (var odb = new dbHRMS(ConnectionString))
                {
                    string DepartmentValue, DesignationValue, LocationValue, BranchValue, PayrollValue, FromEmp;
                    DepartmentValue = cmbDepartment.SelectedItem.ToString() == "All" ? string.Empty : cmbDepartment.SelectedItem.ToString();
                    DesignationValue = cmbDesignation.SelectedItem.ToString() == "All" ? string.Empty : cmbDesignation.SelectedItem.ToString();
                    LocationValue = cmbLocation.SelectedItem.ToString() == "All" ? string.Empty : cmbLocation.SelectedItem.ToString();
                    BranchValue = cmbBranch.SelectedItem.ToString() == "All" ? string.Empty : cmbBranch.SelectedItem.ToString();
                    PayrollValue = cmbPayroll.SelectedItem.ToString();
                    FromEmp = (txtFromEmployeeCode.Text == "0" || string.IsNullOrEmpty(txtFromEmployeeCode.Text)) ? string.Empty : txtFromEmployeeCode.Text;

                    IEnumerable<MstEmployee> oCollection;

                    oCollection = (from a in odb.MstEmployee
                                   where a.FlgActive == true
                                   && a.ResignDate == null
                                   && a.CfgPayrollDefination.PayrollName == PayrollValue
                                   && (string.IsNullOrEmpty(DepartmentValue) || a.DepartmentName == DepartmentValue)
                                   && (string.IsNullOrEmpty(DesignationValue) || a.DesignationName == DesignationValue)
                                   && (string.IsNullOrEmpty(LocationValue) || a.LocationName == LocationValue)
                                   && (string.IsNullOrEmpty(BranchValue) || a.BranchName == BranchValue)
                                   && (string.IsNullOrEmpty(FromEmp) || a.EmpID == FromEmp)
                                   orderby a.ID
                                   select a).ToList();

                    if (oCollection.Count() > 0)
                    {
                        foreach (var emp in oCollection)
                        {
                            dtEmployees.Rows.Add(false, emp.EmpID, $"{emp.FirstName} {emp.MiddleName} {emp.LastName}", emp.DepartmentName, emp.DesignationName, emp.LocationName, emp.BranchName, emp.PayrollName);
                        }
                    }
                }
                grdEmployee.DataSource = null;
                grdEmployee.DataSource = dtEmployees;
                grdEmployee.BestFitColumns();
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        void StartProcess()
        {
            try
            {
                grdEmployee.Visible = false;
                Serial = 1;
                dtProcessed.Rows.Clear();

                grdProcess.Visible = true;
                lblStart.Text = "Processing attendance started, Please wait.";
                using (dbHRMS odb = new dbHRMS(ConnectionString))
                {
                    if (dtEmployees.Rows.Count > 0)
                    {
                        var SelectedEmployees = (from a in dtEmployees.AsEnumerable()
                                                 where a.Field<bool>("Select") == true
                                                 select a).ToList();

                        foreach (var emp in SelectedEmployees)
                        {
                            var oEmp = (from a in odb.MstEmployee
                                        where a.EmpID == emp.Field<string>("EmpCode")
                                        select a).FirstOrDefault();
                            if (oEmp is null)
                            {
                                logger.Info($"employee {emp.Field<string>("EmpCode")} not found.");
                                continue;
                            }
                            var oPayroll = (from a in odb.CfgPayrollDefination
                                            where a.ID == oEmp.PayrollID
                                            select a).FirstOrDefault();
                            if (oPayroll is null)
                            {
                                logger.Info($"payroll not found. for employee {oEmp.EmpID}");
                                continue;
                            }
                            var oPeriod = (from a in odb.CfgPeriodDates
                                           where a.PayrollId == oPayroll.ID
                                           && a.PeriodName == cmbPeriod.SelectedItem.ToString()
                                           select a).FirstOrDefault();

                            if (oPeriod is null)
                            {
                                logger.Info($"period not found. for employee {oEmp.EmpID}");
                                continue;
                            }

                            var oCalendar = (from a in odb.MstCalendar
                                             where a.FlgActive == true
                                             select a).FirstOrDefault();

                            if (oCalendar is null)
                            {
                                logger.Info("calendar not active or define.");
                                continue;
                            }

                            var oEmpLeave = (from a in odb.MstEmployeeLeaves
                                             where a.EmpID == oEmp.ID
                                             && a.LeaveCalCode == oCalendar.Code
                                             && a.LeavesEntitled > 0
                                             select a).ToList();
                            List<LeaveStructure> oLeaves = new List<LeaveStructure>();
                            if (oEmpLeave is null)
                            {
                                logger.Info($"employee leave not assign, for employee {oEmp.EmpID}");
                            }
                            else
                            {
                                foreach (var leave in oEmpLeave)
                                {
                                    //if (leave.ID != 1 || leave.ID != 2 || leave.ID != 3 || leave.ID != 10) continue;
                                    LeaveStructure oDoc = new LeaveStructure();
                                    oDoc.ID = leave.LeaveType ?? 0;
                                    oDoc.LeaveType = leave.MstLeaveType.Description;
                                    oDoc.Balance = leave.LeavesEntitled + leave.LeavesCarryForward;
                                    switch (oDoc.ID)
                                    {
                                        case 3:
                                            oDoc.Priority = 3; break;
                                        case 1:
                                            oDoc.Priority = 2; break;
                                        case 2:
                                            oDoc.Priority = 1; break;
                                        case 7:
                                            continue;
                                        //oDoc.Priority = 4; break;
                                        case 6:
                                            continue;
                                        //oDoc.Priority = 5; break;
                                        case 5:
                                            continue;
                                        //oDoc.Priority = 6; break;
                                        case 4:
                                            continue;
                                        //oDoc.Priority = 7; break;
                                        case 10:
                                            oDoc.Priority = 10; break;
                                        case 8:
                                            continue;
                                        //oDoc.Priority = 11; break;
                                        case 9:
                                            continue;
                                        //oDoc.Priority = 12; break;
                                        case 11:
                                            continue;
                                        //oDoc.Priority = 13; break;
                                        case 12:
                                            continue;
                                            //oDoc.Priority = 14; break;

                                    }
                                    oLeaves.Add(oDoc);
                                }
                            }
                            ProcessEmployeeMonth(oEmp, oPeriod, oLeaves);
                            //Addsandwich leave conditions.
                            SandwichLeaves(oEmp);
                        }
                    }
                }

                grdProcess.DataSource = null;
                grdProcess.DataSource = dtProcessed;
                grdProcess.BestFitColumns();
                lblStart.Text = "Processing attendance Completed.";
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                lblStart.Text = "Processing attendance Completed, with error";
            }
        }
        void ProcessEmployeeMonth(MstEmployee oEmp, CfgPeriodDates oPeriod, List<LeaveStructure> oLeaveBal)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    //DateTime loopStart = oPeriod.StartDate.GetValueOrDefault();
                    DateTime loopStart = dtFrom.Value;
                    //DateTime loopEnd = oPeriod.EndDate.GetValueOrDefault();
                    DateTime loopEnd = dtTo.Value;
                    int ShortLeave = 0;
                    int PenaltyCount = 0;
                    int GracePeriod = 0;
                    var oAttendanceRule = (from a in odb.MstAttendanceRule
                                           select a).FirstOrDefault();
                    for (DateTime i = loopStart; i <= loopEnd; i = i.AddDays(1))
                    {
                        if (i < oEmp.JoiningDate.GetValueOrDefault().Date)
                        {
                            continue;
                        }
                        var oAttendance = (from a in odb.TrnsAttendanceRegister
                                           where a.EmpID == oEmp.ID
                                           && a.PeriodID == oPeriod.ID
                                           && a.Date == i.Date
                                           select a).FirstOrDefault();
                        if (oAttendance is null)
                        {
                            logger.Warn($"employee {oEmp.EmpID} shift is not assigned.");
                            continue;
                        }
                        bool flgSaved = oAttendance.Processed.GetValueOrDefault();
                        //Attendace not saved
                        if (!flgSaved)
                        {
                            #region not saved

                            logger.Info($"processing will be done attendance not saved.");
                            string ShiftIn, ShiftOut, ShiftDuration, StartBuffer, EndBuffer;
                            int iShiftIn, iShiftOut, iShiftDuration, iStartBuffer, iEndBuffer;
                            bool flgConfirmedOut, flgExpectedOut;
                            string TimeIn, TimeOut, WorkHour;
                            int iTimeIn, iTimeOut, iWorkHour;
                            bool flgEarlyIn = false, flgEarlyOut = false, flgLateIn = false, flgLateOut = false;
                            string EarlyIn = string.Empty, EarlyOut = string.Empty, LateIn = string.Empty, LateOut = string.Empty;
                            bool flgNewLeave = false;
                            bool flgFullDayLeave = false;
                            string LeaveHour = string.Empty, LeaveType = string.Empty;
                            int LeaveTypeID = 0;
                            decimal LeaveCount = 0;

                            int OTId = 0;
                            string OTHours = string.Empty, OTType = string.Empty;
                            bool flgOT = false;

                            #region Shift Data

                            ShiftIn = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().StartTime;
                            ShiftOut = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().EndTime;
                            ShiftDuration = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().Duration;
                            StartBuffer = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().BufferStartTime;
                            EndBuffer = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().BufferEndTime;
                            flgConfirmedOut = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().FlgOutOverlap.GetValueOrDefault();
                            flgExpectedOut = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().FlgExpectedOut.GetValueOrDefault();
                            iShiftIn = TimeConvert(ShiftIn);
                            iShiftOut = TimeConvert(ShiftOut);
                            iShiftDuration = TimeConvert(ShiftDuration);
                            iStartBuffer = string.IsNullOrEmpty(StartBuffer) ? 0 : TimeConvert(StartBuffer);
                            iEndBuffer = string.IsNullOrEmpty(EndBuffer) ? 0 : TimeConvert(EndBuffer);
                            #endregion

                            #region Time In/Out

                            if (!flgConfirmedOut)
                            {
                                if (flgExpectedOut)
                                {
                                    TimeIn = (from a in odb.TrnsTempAttendance
                                              where (a.In_Out == "1" || a.In_Out == "01" || a.In_Out == "In")
                                              && a.PunchedDate == oAttendance.Date
                                              && a.EmpID == oEmp.EmpID
                                              orderby a.ID ascending
                                              select a.PunchedTime).FirstOrDefault();
                                    if (string.IsNullOrEmpty(TimeIn))
                                    {
                                        TimeIn = string.Empty;
                                        iTimeIn = 0;
                                    }
                                    else
                                    {
                                        iTimeIn = TimeConvert(TimeIn);
                                    }
                                    TimeOut = (from a in odb.TrnsTempAttendance
                                               where (a.In_Out == "2" || a.In_Out == "02" || a.In_Out == "Out")
                                               && a.PunchedDate == oAttendance.Date
                                               && a.EmpID == oEmp.EmpID
                                               orderby a.ID descending
                                               select a.PunchedTime).FirstOrDefault();
                                    if (string.IsNullOrEmpty(TimeOut))
                                    {
                                        TimeOut = string.Empty;
                                        iTimeOut = 0;
                                    }
                                    else
                                    {
                                        iTimeOut = TimeConvert(TimeOut);
                                    }
                                    if (iTimeIn > 0 && iTimeOut == 0)
                                    {
                                        TimeOut = (from a in odb.TrnsTempAttendance
                                                   where (a.In_Out == "2" || a.In_Out == "02" || a.In_Out == "Out")
                                                   && a.PunchedDate == oAttendance.Date.Value.AddDays(1)
                                                   && a.EmpID == oEmp.EmpID
                                                   orderby a.ID ascending
                                                   select a.PunchedTime).FirstOrDefault();
                                        if (string.IsNullOrEmpty(TimeOut))
                                        {
                                            TimeOut = string.Empty;
                                            iTimeOut = 0;
                                        }
                                        else
                                        {
                                            iTimeOut = TimeConvert(TimeOut);
                                            int intShiftIn = TimeConvert(ShiftIn);
                                            if (iTimeOut > intShiftIn)
                                            {
                                                TimeOut = string.Empty;
                                                iTimeOut = 0;
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    TimeIn = (from a in odb.TrnsTempAttendance
                                              where (a.In_Out == "1" || a.In_Out == "01" || a.In_Out == "In")
                                              && a.PunchedDate == oAttendance.Date
                                              && a.EmpID == oEmp.EmpID
                                              orderby a.ID ascending
                                              select a.PunchedTime).FirstOrDefault();
                                    if (string.IsNullOrEmpty(TimeIn))
                                    {
                                        TimeIn = string.Empty;
                                        iTimeIn = 0;
                                    }
                                    else
                                    {
                                        iTimeIn = TimeConvert(TimeIn);
                                    }
                                    TimeOut = (from a in odb.TrnsTempAttendance
                                               where (a.In_Out == "2" || a.In_Out == "02" || a.In_Out == "Out")
                                               && a.PunchedDate == oAttendance.Date
                                               && a.EmpID == oEmp.EmpID
                                               orderby a.ID descending
                                               select a.PunchedTime).FirstOrDefault();
                                    if (string.IsNullOrEmpty(TimeOut))
                                    {
                                        TimeOut = string.Empty;
                                        iTimeOut = 0;
                                    }
                                    else
                                    {
                                        iTimeOut = TimeConvert(TimeOut);
                                    }
                                }

                                if ((iTimeIn == 0 && iTimeOut == 0)
                                    || (iTimeIn > 0 && iTimeOut == 0)
                                    || (iTimeIn == 0 && iTimeOut > 0))
                                {
                                    WorkHour = "00:00";
                                    iWorkHour = 0;
                                }
                                else
                                {
                                    WorkHour = CalculateWorkHour(TimeIn, TimeOut);
                                    iWorkHour = TimeConvert(WorkHour);
                                }


                            }
                            else
                            {
                                TimeIn = (from a in odb.TrnsTempAttendance
                                          where (a.In_Out == "1" || a.In_Out == "01" || a.In_Out == "In")
                                          && a.PunchedDate == oAttendance.Date
                                          && a.EmpID == oEmp.EmpID
                                          orderby a.ID ascending
                                          select a.PunchedTime).FirstOrDefault();
                                if (string.IsNullOrEmpty(TimeIn))
                                {
                                    TimeIn = string.Empty;
                                    iTimeIn = 0;
                                }
                                else
                                {
                                    iTimeIn = TimeConvert(TimeIn);
                                }
                                TimeOut = (from a in odb.TrnsTempAttendance
                                           where (a.In_Out == "2" || a.In_Out == "02" || a.In_Out == "Out")
                                           && a.PunchedDate == oAttendance.Date.Value.AddDays(1)
                                           && a.EmpID == oEmp.EmpID
                                           orderby a.ID ascending
                                           select a.PunchedTime).FirstOrDefault();
                                if (string.IsNullOrEmpty(TimeOut))
                                {
                                    TimeOut = string.Empty;
                                    iTimeOut = 0;
                                }
                                else
                                {
                                    iTimeOut = TimeConvert(TimeOut);
                                }
                                if ((iTimeIn == 0 && iTimeOut == 0)
                                    || (iTimeIn > 0 && iTimeOut == 0)
                                    || (iTimeIn == 0 && iTimeOut > 0))
                                {
                                    WorkHour = "00:00";
                                    iWorkHour = 0;
                                }
                                else
                                {
                                    WorkHour = CalculateWorkHour(TimeIn, TimeOut);
                                    iWorkHour = TimeConvert(WorkHour);
                                }
                            }

                            #endregion

                            #region Early/Late Ins

                            if (iTimeIn > 0)
                            {
                                if (iTimeIn - iShiftIn > 0)
                                {
                                    //outside graceperiod
                                    int iGracePeriod = TimeConvert(oAttendanceRule.GpAfterStartTime);
                                    if ((iTimeIn - iShiftIn) > iGracePeriod)
                                    {
                                        LateIn = TimeConvert(iTimeIn - iShiftIn);
                                        flgLateIn = true;
                                    }
                                }
                                if (iShiftIn - iTimeIn > 0)
                                {
                                    EarlyIn = TimeConvert(iShiftIn - iTimeIn);
                                    flgEarlyIn = true;
                                }
                            }
                            if (iTimeOut > 0)
                            {
                                if (iTimeOut - iShiftOut > 0)
                                {
                                    LateOut = TimeConvert(iTimeOut - iShiftOut);
                                    flgLateOut = true;
                                }
                                if (iShiftOut - iTimeOut > 0)
                                {
                                    EarlyOut = TimeConvert(iShiftOut - iTimeOut);
                                    flgEarlyOut = true;
                                }
                            }

                            #endregion

                            #region Penalty & Leaves

                            //full day leave
                            if (iTimeIn == 0 && iTimeOut == 0 && iShiftDuration > 0)
                            {
                                var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                                  where a.LeaveFrom <= i && a.LeaveTo >= i
                                                  && a.EmpID == oEmp.ID
                                                  select a).Count();
                                if (LeaveCheck == 0)
                                {
                                    var oLT = (from a in oLeaveBal
                                               where a.Balance > 0
                                               && a.LeaveType.ToLower() == "absent"
                                               select a).FirstOrDefault();
                                    if (oLT == null) { RadMessageBox.Show("Leave type not assign to employee Absent."); continue; }
                                    flgNewLeave = true;
                                    LeaveHour = ShiftDuration;
                                    LeaveType = oLT.LeaveType;
                                    LeaveTypeID = oLT.ID;
                                    LeaveCount = 1;
                                    oLT.Balance -= 1;
                                }
                                else
                                {
                                    var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                         where a.LeaveFrom <= i && a.LeaveTo >= i
                                                         && a.EmpID == oEmp.ID
                                                         select a).FirstOrDefault();
                                    flgNewLeave = false;
                                    LeaveHour = ShiftDuration;
                                    LeaveType = oLeaveRequest.MstLeaveType.Description;
                                    LeaveCount = 1;
                                    LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                                }
                                flgFullDayLeave = true;
                            }

                            //Missing timein or timeout
                            if ((iTimeIn > 0 && iTimeOut == 0 && iShiftDuration > 0) || (iTimeIn == 0 && iTimeOut > 0 && iShiftDuration > 0))
                            {
                                var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                                  where a.LeaveFrom <= i && a.LeaveTo >= i
                                                  && a.EmpID == oEmp.ID
                                                  select a).Count();
                                if (LeaveCheck == 0)
                                {
                                    var oLT = (from a in oLeaveBal
                                               where a.Balance > 0
                                               && a.LeaveType.ToLower() == "absent"
                                               select a).FirstOrDefault();
                                    if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); continue; }
                                    flgNewLeave = true;
                                    LeaveHour = ShiftDuration;
                                    LeaveType = oLT.LeaveType;
                                    LeaveTypeID = oLT.ID;
                                    LeaveCount = 1;
                                    oLT.Balance -= 1;
                                }
                                else
                                {
                                    var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                         where a.LeaveFrom <= i && a.LeaveTo >= i
                                                         && a.EmpID == oEmp.ID
                                                         select a).FirstOrDefault();
                                    flgNewLeave = false;
                                    LeaveHour = ShiftDuration;
                                    LeaveType = oLeaveRequest.MstLeaveType.Description;
                                    LeaveCount = oLeaveRequest.TotalCount.GetValueOrDefault();
                                    LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                                }
                            }

                            //if workhour and shift hour is ok
                            ////but not on shift time.
                            if (iShiftDuration <= iWorkHour && iShiftDuration > 0)
                            {
                                if ((flgLateIn && flgEarlyOut) || (!flgLateIn && flgEarlyOut) || (flgLateIn && !flgEarlyOut))
                                {
                                    string TimeDiff = TimeConvert(TimeConvert(LateIn) + TimeConvert(EarlyOut));
                                    var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                                      where a.LeaveFrom <= i && a.LeaveTo >= i
                                                      && a.EmpID == oEmp.ID
                                                      select a).Count();
                                    string DeductionRule = CalculateDeductionRule(TimeConvert(TimeDiff));
                                    int GracePeriodDR = ReturnGracePeriodDeductionRule(DeductionRule);

                                    if (DeductionRule == "DR_01")
                                    {
                                        ShortLeave++;
                                        GracePeriod++;
                                    }
                                    if (LeaveCheck == 0)
                                    {
                                        if (DeductionRule == "DR_01")
                                        {
                                            if (GracePeriod > GracePeriodDR)
                                            {

                                                var oLT = (from a in oLeaveBal
                                                           where a.Balance > 0
                                                           orderby a.Priority ascending
                                                           select a).FirstOrDefault();
                                                if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); continue; }
                                                flgNewLeave = true;
                                                LeaveHour = TimeDiff;
                                                LeaveType = oLT.LeaveType;
                                                LeaveTypeID = oLT.ID;
                                                LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                                oLT.Balance -= 1;
                                            }
                                            else
                                            {
                                                logger.Info($"deduction rule 1 grace period {GracePeriod} date {i.ToString("dd MMM yyyy")} employee {oEmp.EmpID} no deduction.");
                                            }
                                        }
                                        else if (DeductionRule == "DR_02")
                                        {
                                            var oLT = (from a in oLeaveBal
                                                       where a.Balance > 0
                                                       orderby a.Priority ascending
                                                       select a).FirstOrDefault();
                                            if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); continue; }
                                            flgNewLeave = true;
                                            LeaveHour = TimeDiff;
                                            LeaveType = oLT.LeaveType;
                                            LeaveTypeID = oLT.ID;
                                            LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                            oLT.Balance -= 1;
                                        }
                                        else if (DeductionRule == "DR_03")
                                        {
                                            var oLT = (from a in oLeaveBal
                                                       where a.Balance > 0
                                                       orderby a.Priority ascending
                                                       select a).FirstOrDefault();
                                            if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); continue; }
                                            flgNewLeave = true;
                                            LeaveHour = TimeDiff;
                                            LeaveType = oLT.LeaveType;
                                            LeaveTypeID = oLT.ID;
                                            LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                            oLT.Balance -= 1;
                                        }
                                    }
                                    else
                                    {
                                        var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                             where a.LeaveFrom <= i && a.LeaveTo >= i
                                                             && a.EmpID == oEmp.ID
                                                             select a).FirstOrDefault();
                                        flgNewLeave = false;
                                        LeaveHour = TimeDiff;
                                        LeaveType = oLeaveRequest.MstLeaveType.Description;
                                        LeaveCount = oLeaveRequest.TotalCount.GetValueOrDefault();
                                        LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                                    }
                                }
                            }

                            //if workhour is less then shift hour
                            if (iShiftDuration > iWorkHour && iTimeIn > 0 && iTimeOut > 0)
                            {
                                string TimeDiff = TimeConvert(iShiftDuration - iWorkHour);
                                var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                                  where a.LeaveFrom <= i && a.LeaveTo >= i
                                                  && a.EmpID == oEmp.ID
                                                  select a).Count();
                                string DeductionRule = CalculateDeductionRule(TimeConvert(TimeDiff));
                                int GracePeriodDR = ReturnGracePeriodDeductionRule(DeductionRule);

                                if (DeductionRule == "DR_01")
                                {
                                    ShortLeave++;
                                    GracePeriod++;
                                }
                                if (LeaveCheck == 0)
                                {
                                    if (DeductionRule == "DR_01")
                                    {
                                        if (GracePeriod > GracePeriodDR)
                                        {
                                            logger.Info($"deduction rule 1 grace period {GracePeriod} date {i.ToString("dd MMM yyyy")} employee {oEmp.EmpID} deducted.");
                                            var oLT = (from a in oLeaveBal
                                                       where a.Balance > 0
                                                       orderby a.Priority ascending
                                                       select a).FirstOrDefault();
                                            if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); continue; }
                                            flgNewLeave = true;
                                            LeaveHour = TimeDiff;
                                            LeaveType = oLT.LeaveType;
                                            LeaveTypeID = oLT.ID;
                                            LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                            oLT.Balance -= 1;
                                        }
                                        else
                                        {
                                            logger.Info($"deduction rule 1 grace period {GracePeriod} date {i.ToString("dd MMM yyyy")} employee {oEmp.EmpID} no deduction.");
                                        }
                                    }
                                    else if (DeductionRule == "DR_02")
                                    {
                                        var oLT = (from a in oLeaveBal
                                                   where a.Balance > 0
                                                   orderby a.Priority ascending
                                                   select a).FirstOrDefault();
                                        if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); continue; }
                                        flgNewLeave = true;
                                        LeaveHour = TimeDiff;
                                        LeaveType = oLT.LeaveType;
                                        LeaveTypeID = oLT.ID;
                                        LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                        oLT.Balance -= 1;
                                    }
                                    else if (DeductionRule == "DR_03")
                                    {
                                        var oLT = (from a in oLeaveBal
                                                   where a.Balance > 0
                                                   orderby a.Priority ascending
                                                   select a).FirstOrDefault();
                                        if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); continue; }
                                        flgNewLeave = true;
                                        LeaveHour = TimeDiff;
                                        LeaveType = oLT.LeaveType;
                                        LeaveTypeID = oLT.ID;
                                        LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                        oLT.Balance -= 1;
                                    }
                                }
                                else
                                {
                                    var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                         where a.LeaveFrom <= i && a.LeaveTo >= i
                                                         && a.EmpID == oEmp.ID
                                                         select a).FirstOrDefault();
                                    flgNewLeave = false;
                                    LeaveHour = TimeDiff;
                                    LeaveType = oLeaveRequest.MstLeaveType.Description;
                                    LeaveCount = oLeaveRequest.TotalCount.GetValueOrDefault();
                                    LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                                }
                            }

                            //Missing In/Out Penalty
                            if ((iTimeIn == 0 || iTimeOut == 0) && iShiftDuration > 0 && !flgFullDayLeave)
                            {
                                string DeductionRule = string.Empty;

                                if (oAttendanceRule.FlgMissingTimePenalty.GetValueOrDefault())
                                {
                                    DeductionRule = "DR_04";
                                    PenaltyCount += 1;
                                }
                                if (PenaltyCount >= oAttendanceRule.MissingTimePanaltyCounter.GetValueOrDefault())
                                {
                                    PenaltyCount = 0;
                                    MstLeaveType oLeaveMaster = null;
                                    if (DeductionRule == "DR_04")
                                    {
                                        oLeaveMaster = (from a in odb.MstLeaveType where a.Code == oAttendanceRule.MPTLeaveType select a).FirstOrDefault();
                                    }
                                    else
                                    {
                                        oLeaveMaster = (from a in odb.MstLeaveType where a.Code == oAttendanceRule.LeaveTypeWOP select a).FirstOrDefault();
                                    }
                                    if (oLeaveMaster != null)
                                    {
                                        flgNewLeave = true;
                                        LeaveHour = ShiftDuration;
                                        LeaveType = oLeaveMaster.Description;
                                        LeaveTypeID = oLeaveMaster.ID;
                                        LeaveCount = 0.5M;
                                    }

                                }
                            }

                            #endregion

                            #region Overtime

                            if (iWorkHour > iShiftDuration)
                            {
                                //checking timeout is also more than 
                                //expected
                                if (flgLateOut)
                                {
                                    if (iEndBuffer < TimeConvert(LateOut))
                                    {
                                        flgOT = true;
                                        OTHours = LateOut;
                                        OTType = oAttendance.MstShifts.MstOverTime == null ? "" : oAttendance.MstShifts.MstOverTime.Description;
                                        OTId = oAttendance.MstShifts.MstOverTime == null ? 0 : oAttendance.MstShifts.MstOverTime.ID;
                                    }
                                }
                            }

                            #endregion

                            #region Set Datatable

                            DataRow dr = dtProcessed.NewRow();
                            dr["Serial"] = Serial;
                            dr["EmpCode"] = oEmp.EmpID;
                            dr["EmpName"] = $"{oEmp.FirstName} {oEmp.MiddleName} {oEmp.LastName}";
                            dr["ID"] = oAttendance.Id;
                            dr["Date"] = oAttendance.Date;
                            dr["Day"] = oAttendance.Date.Value.DayOfWeek.ToString();
                            dr["Shift"] = oAttendance.MstShifts.Description;
                            dr["ShiftIn"] = ShiftIn;
                            dr["ShiftOut"] = ShiftOut;
                            dr["ShiftDuration"] = ShiftDuration;
                            dr["In"] = TimeIn;
                            dr["Out"] = TimeOut;
                            dr["WorkHour"] = WorkHour;
                            dr["EarlyIn"] = EarlyIn;
                            dr["LateIn"] = LateIn;
                            dr["EarlyOut"] = EarlyOut;
                            dr["LateOut"] = LateOut;
                            dr["LeaveHour"] = LeaveHour;
                            dr["LeaveType"] = LeaveType;
                            dr["LeaveNew"] = flgNewLeave;
                            dr["LeaveCount"] = LeaveCount;
                            dr["LTID"] = LeaveTypeID;
                            dr["OTHour"] = OTHours;
                            dr["OTType"] = OTType;
                            dr["OTID"] = OTId;
                            dr["Status"] = "Draft";
                            Serial++;
                            dtProcessed.Rows.Add(dr);
                            #endregion

                            #endregion
                        }
                        //Attendace saved didn't re-calculate
                        else
                        {
                            logger.Info($"saved data was loaded.");
                            #region saved

                            string ShiftIn, ShiftOut, ShiftDuration;
                            int iShiftIn, iShiftOut, iShiftDuration;
                            bool flgConfirmedOut, flgExpectedOut;
                            string TimeIn, TimeOut, WorkHour;
                            int iTimeIn, iTimeOut, iWorkHour;
                            bool flgEarlyIn = false, flgEarlyOut = false, flgLateIn = false, flgLateOut = false;
                            string EarlyIn = string.Empty, EarlyOut = string.Empty, LateIn = string.Empty, LateOut = string.Empty;
                            bool flgNewLeave = false;
                            string LeaveHour = string.Empty, LeaveType = string.Empty;
                            int LeaveTypeID = 0;
                            decimal LeaveCount = 0;

                            int OTId = 0;
                            string OTHours = string.Empty, OTType = string.Empty;
                            bool flgOT = false;

                            #region Shift Data

                            ShiftIn = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().StartTime;
                            ShiftOut = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().EndTime;
                            ShiftDuration = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().Duration;
                            flgConfirmedOut = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().FlgOutOverlap.GetValueOrDefault();
                            flgExpectedOut = oAttendance.MstShifts.MstShiftDetails.Where(a => a.Day == oAttendance.DateDay).FirstOrDefault().FlgExpectedOut.GetValueOrDefault();
                            iShiftIn = TimeConvert(ShiftIn);
                            iShiftOut = TimeConvert(ShiftOut);
                            iShiftDuration = TimeConvert(ShiftDuration);

                            #endregion

                            #region Time In/Out
                            TimeOut = oAttendance.TimeOut;
                            if (string.IsNullOrEmpty(oAttendance.TimeIn))
                            {
                                TimeIn = string.Empty;
                            }
                            else
                            {
                                TimeIn = oAttendance.TimeIn;
                            }
                            if (string.IsNullOrEmpty(oAttendance.TimeOut))
                            {
                                TimeOut = string.Empty;
                            }
                            else
                            {
                                TimeOut = oAttendance.TimeOut;
                            }
                            #endregion

                            #region Early/Late In/Out

                            EarlyIn = oAttendance.EarlyInMin;
                            LateIn = oAttendance.LateInMin;
                            EarlyOut = oAttendance.EarlyOutMin;
                            LateOut = oAttendance.LateOutMin;

                            #endregion

                            #region Leaves
                            if (oAttendance.FlgIsNewLeave.GetValueOrDefault())
                            {
                                flgNewLeave = true;
                                LeaveType = oAttendance.MstLeaveType.Description;
                                LeaveCount = oAttendance.LeaveCount.GetValueOrDefault();
                                LeaveTypeID = oAttendance.LeaveType.GetValueOrDefault();
                                LeaveHour = oAttendance.LeaveHour;
                            }
                            #endregion

                            #region Overtime
                            if (!string.IsNullOrEmpty(oAttendance.OTHour))
                            {
                                OTHours = oAttendance.OTHour;
                                OTType = oAttendance.MstOverTime.Description;
                                OTId = oAttendance.OTType.GetValueOrDefault();
                            }

                            #endregion

                            #region Set Datatable

                            DataRow dr = dtProcessed.NewRow();
                            dr["Serial"] = Serial;
                            dr["EmpCode"] = oEmp.EmpID;
                            dr["EmpName"] = $"{oEmp.FirstName} {oEmp.MiddleName} {oEmp.LastName}";
                            dr["ID"] = oAttendance.Id;
                            dr["Date"] = oAttendance.Date;
                            dr["Day"] = oAttendance.Date.Value.DayOfWeek.ToString();
                            dr["Shift"] = oAttendance.MstShifts.Description;
                            dr["ShiftIn"] = ShiftIn;
                            dr["ShiftOut"] = ShiftOut;
                            dr["ShiftDuration"] = ShiftDuration;
                            dr["In"] = TimeIn;
                            dr["Out"] = TimeOut;
                            dr["WorkHour"] = oAttendance.WorkHour;
                            dr["EarlyIn"] = EarlyIn;
                            dr["LateIn"] = LateIn;
                            dr["EarlyOut"] = EarlyOut;
                            dr["LateOut"] = LateOut;
                            dr["LeaveHour"] = LeaveHour;
                            dr["LeaveType"] = LeaveType;
                            dr["LeaveNew"] = flgNewLeave;
                            dr["LeaveCount"] = LeaveCount;
                            dr["LTID"] = LeaveTypeID;
                            dr["OTHour"] = OTHours;
                            dr["OTType"] = OTType;
                            dr["OTID"] = OTId;
                            if (oAttendance.FlgPost.GetValueOrDefault())
                            {
                                dr["Status"] = "Posted";
                            }
                            else
                            {
                                dr["Status"] = "Processed";
                            }
                            Serial++;
                            dtProcessed.Rows.Add(dr);
                            #endregion

                            #endregion
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        void SandwichLeaves(MstEmployee oEmp)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    DataTable dtEmployee = dtProcessed.Clone();
                    foreach (DataRow row in dtProcessed.Select($"EmpCode = {oEmp.EmpID}"))
                    {
                        dtEmployee.ImportRow(row);
                        row.Delete();
                    }
                    dtProcessed.AcceptChanges();

                    for (int i = 0; i < dtEmployee.Rows.Count; i++)
                    {
                        string EmpCode = string.Empty;
                        bool flgNewLeave = false;
                        string shifthour = string.Empty;
                        decimal LeaveCount = 0;

                        EmpCode = Convert.ToString(dtEmployee.Rows[i]["EmpCode"]);
                        flgNewLeave = Convert.ToBoolean(dtEmployee.Rows[i]["LeaveNew"]);
                        shifthour = Convert.ToString(dtEmployee.Rows[i]["ShiftDuration"]);
                        LeaveCount = Convert.ToDecimal(dtEmployee.Rows[i]["LeaveCount"]);
                        if (flgNewLeave && LeaveCount == 1)
                        {
                            //get next days value
                            int nextday = i + 1;
                            if (nextday < dtEmployee.Rows.Count)
                            {
                                string nextdayshifthour = string.Empty;
                                int ndshifthour = 0;
                                nextdayshifthour = Convert.ToString(dtEmployee.Rows[nextday]["ShiftDuration"]);
                                //check if its holiday or not.
                                ndshifthour = TimeConvert(nextdayshifthour);
                                if (ndshifthour == 0)
                                {
                                    //check for next day if its possible.
                                    int daythree = nextday + 1;
                                    if (daythree < dtEmployee.Rows.Count)
                                    {
                                        //next day is holiday get day after tommorrow.
                                        bool datflgNewLeave = false;
                                        datflgNewLeave = Convert.ToBoolean(dtEmployee.Rows[i]["LeaveNew"]);
                                        if (datflgNewLeave)
                                        {
                                            var oLeaveType = (from a in odb.MstLeaveType
                                                              where a.Code.ToLower() == "absent"
                                                              select a).FirstOrDefault();
                                            if (oLeaveType != null)
                                            {
                                                dtEmployee.Rows[nextday]["LeaveHour"] = shifthour;
                                                dtEmployee.Rows[nextday]["LeaveType"] = oLeaveType.Description;
                                                dtEmployee.Rows[nextday]["LeaveCount"] = 1;
                                                dtEmployee.Rows[nextday]["LTID"] = oLeaveType.ID;
                                                dtEmployee.Rows[nextday]["LeaveNew"] = true;
                                            }
                                            //now day three is also on leave, 
                                            //then mark day two for sandwich leave.
                                            //flgNewLeave = true;
                                            //LeaveHour = ShiftDuration;
                                            //LeaveType = oLT.LeaveType;
                                            //LeaveTypeID = oLT.ID;
                                            //LeaveCount = 1;
                                            //dr["LeaveHour"] = LeaveHour;
                                            //dr["LeaveType"] = LeaveType;
                                            //dr["LeaveNew"] = flgNewLeave;
                                            //dr["LeaveCount"] = LeaveCount;
                                            //dr["LTID"] = LeaveTypeID;
                                        }
                                    }
                                }
                            }

                        }
                    }

                    foreach (DataRow dr in dtEmployee.Rows)
                    {
                        dtProcessed.ImportRow(dr);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        int TimeConvert(string value)
        {
            try
            {
                string[] parts = value.Split(':');
                if (parts.Count() == 2)
                {
                    int hours = int.Parse(parts[0]) * 60;
                    int minutes = int.Parse(parts[1]);
                    return hours + minutes;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return 0;
            }
        }
        string TimeConvert(int value)
        {
            try
            {
                int hours = Convert.ToInt32(value / 60);
                int minutes = value % 60;
                string retvalue = $"{string.Format("{00:D2}", hours)}:{string.Format("{00:D2}", minutes)}";
                return retvalue;
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return "";
            }
        }
        bool TimeValidate(string value)
        {
            try
            {
                string[] TempValue = value.Split(':');
                if (TempValue.Count() == 2)
                {
                    int hour = Convert.ToInt32(TempValue[0]);
                    int min = Convert.ToInt32(TempValue[1]);
                    if (hour + min > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return false;
            }
        }
        string CalculateWorkHour(string TimeIn, string TimeOut)
        {
            try
            {
                int Time1 = TimeConvert(TimeIn);
                int Time2 = TimeConvert(TimeOut);
                int Diff = Time2 - Time1;
                return TimeConvert(Diff);
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return "";
            }
        }
        decimal CalculateDeductionCount(int pTimeDifference)
        {
            try
            {
                using (dbHRMS odb = new dbHRMS(ConnectionString))
                {
                    var oDedRule = (from a in odb.MstDeductionRules
                                    select a).ToList();
                    if (oDedRule.Count == 0)
                    {
                        logger.Info("no deduction found in company config. CalculateDeductionCount()");
                        return 0M;
                    }
                    foreach (var rule in oDedRule)
                    {
                        int lowerBound = TimeConvert(rule.RangeFrom);
                        int upperBound = TimeConvert(rule.RangeTo);
                        if (lowerBound < pTimeDifference && upperBound >= pTimeDifference)
                        {
                            return rule.LeaveCount.GetValueOrDefault();
                        }
                    }
                    return 0M;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return 0M;
            }
        }
        string CalculateDeductionRule(int pTimeDifference)
        {
            try
            {
                using (dbHRMS odb = new dbHRMS(ConnectionString))
                {
                    var oDedRule = (from a in odb.MstDeductionRules
                                    select a).ToList();
                    if (oDedRule.Count == 0)
                    {
                        logger.Info("no deduction found in company config. CalculateDeductionCount()");
                        return "";
                    }
                    foreach (var rule in oDedRule)
                    {
                        int lowerBound = TimeConvert(rule.RangeFrom);
                        int upperBound = TimeConvert(rule.RangeTo);
                        if (lowerBound < pTimeDifference && upperBound >= pTimeDifference)
                        {
                            return rule.Code;
                        }
                    }
                    return "";
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return "";
            }
        }
        int ReturnGracePeriodDeductionRule(string DeductionRule)
        {
            try
            {
                using (dbHRMS odb = new dbHRMS(ConnectionString))
                {
                    var oDedRule = (from a in odb.MstDeductionRules
                                    where a.Code == DeductionRule
                                    select a).FirstOrDefault();
                    if (oDedRule.GracePeriod == 0)
                    {
                        logger.Info($"no grace period on {DeductionRule}");
                        return 0;
                    }
                    else
                    {
                        return oDedRule.GracePeriod.Value;
                    }

                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return 0;
            }
        }
        void SaveAttendance()
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    lblStart.Text = "Saving attendance started, Please wait.";
                    if (dtProcessed.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtProcessed.Rows)
                        {
                            //getting all info from dr.
                            int AttendanceId = 0;
                            string TimeIn = string.Empty, TimeOut = string.Empty, WorkHour = string.Empty;
                            string EarlyIn = string.Empty, LateIn = string.Empty, EarlyOut = string.Empty, LateOut = string.Empty;
                            string LeaveType = string.Empty, OTType = string.Empty, LeaveHour = string.Empty, OTHour = string.Empty;
                            decimal LeaveCount = 0;
                            int LeaveID = 0, OTID = 0;
                            bool flgNewLeave = false, flgOT = false;
                            string AttStatus = string.Empty;

                            AttendanceId = Convert.ToInt32(dr["ID"]);
                            TimeIn = Convert.ToString(dr["In"]);
                            TimeOut = Convert.ToString(dr["Out"]);
                            WorkHour = Convert.ToString(dr["WorkHour"]);

                            EarlyIn = Convert.ToString(dr["EarlyIn"]);
                            LateIn = Convert.ToString(dr["LateIn"]);
                            EarlyOut = Convert.ToString(dr["EarlyOut"]);
                            LateOut = Convert.ToString(dr["LateOut"]);

                            flgNewLeave = Convert.ToBoolean(dr["LeaveNew"]);
                            LeaveType = Convert.ToString(dr["LeaveType"]);
                            LeaveHour = Convert.ToString(dr["LeaveHour"]);
                            LeaveCount = Convert.ToDecimal(dr["LeaveCount"]);
                            LeaveID = Convert.ToInt32(dr["LTID"]);

                            MstLeaveType oLeave;
                            if (LeaveID > 0)
                            {
                                oLeave = (from a in odb.MstLeaveType where a.ID == LeaveID select a).FirstOrDefault();
                            }
                            else
                            {
                                oLeave = null;
                            }

                            OTType = Convert.ToString(dr["OTType"]);
                            OTHour = Convert.ToString(dr["OTHour"]);
                            OTID = Convert.ToInt32(dr["OTID"]);

                            MstOverTime oOT;
                            if (OTID > 0)
                            {
                                oOT = (from a in odb.MstOverTime where a.ID == OTID select a).FirstOrDefault();
                            }
                            else
                            {
                                oOT = null;
                            }


                            AttStatus = Convert.ToString(dr["Status"]);

                            if (AttStatus == "Draft")
                            {
                                var oRecord = (from a in odb.TrnsAttendanceRegister
                                               where a.Id == AttendanceId
                                               select a).FirstOrDefault();
                                if (oRecord is null)
                                {
                                    logger.Info($"attendance record not found for employee {dr["EmpCode"].ToString()}");
                                    continue;
                                }
                                oRecord.TimeIn = TimeIn;
                                oRecord.TimeOut = TimeOut;
                                oRecord.WorkHour = WorkHour;
                                oRecord.EarlyInMin = EarlyIn;
                                oRecord.LateInMin = LateIn;
                                oRecord.EarlyOutMin = EarlyOut;
                                oRecord.LateOutMin = LateOut;
                                if (!(oLeave is null))
                                {
                                    oRecord.FlgIsNewLeave = flgNewLeave;
                                    oRecord.LeaveHour = LeaveHour;
                                    oRecord.LeaveCount = LeaveCount;
                                    oRecord.MstLeaveType = oLeave;
                                }
                                if (!(oOT is null))
                                {
                                    oRecord.OTHour = OTHour;
                                    oRecord.MstOverTime = oOT;
                                }

                                oRecord.Processed = true;
                                oRecord.FlgSave = true;
                                oRecord.FlgPost = false;
                                oRecord.FlgPosted = false;
                            }

                        }
                        odb.SubmitChanges();
                        lblStart.Text = "Saving attendance completed.";
                    }
                }
                grdEmployee.Visible = true;
                grdProcess.Visible = false;
                btnProcess.Text = "Process";
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                lblStart.Text = "Saving attendance completed. with error.";
            }
        }
        void PostAttendance()
        {
            try
            {
                lblStart.Text = "Posting attendance started, Please wait.";
                using (dbHRMS odb = new dbHRMS(ConnectionString))
                {
                    if (dtEmployees.Rows.Count > 0)
                    {
                        int LeaveDocCount = 0;
                        var SelectedEmployees = (from a in dtEmployees.AsEnumerable()
                                                 where a.Field<bool>("Select") == true
                                                 select a).ToList();

                        LeaveDocCount = (from a in odb.TrnsLeavesRequest
                                         orderby a.ID descending
                                         select a.DocNum).FirstOrDefault() ?? 0;

                        foreach (var emp in SelectedEmployees)
                        {
                            var oEmp = (from a in odb.MstEmployee
                                        where a.EmpID == emp.Field<string>("EmpCode")
                                        select a).FirstOrDefault();
                            if (oEmp is null)
                            {
                                logger.Info($"employee {emp.Field<string>("EmpCode")} not found.");
                                continue;
                            }
                            var oPayroll = (from a in odb.CfgPayrollDefination
                                            where a.ID == oEmp.PayrollID
                                            select a).FirstOrDefault();
                            if (oPayroll is null)
                            {
                                logger.Info($"payroll not found. for employee {oEmp.EmpID}");
                                continue;
                            }
                            var oPeriod = (from a in odb.CfgPeriodDates
                                           where a.PayrollId == oPayroll.ID
                                           && a.PeriodName == cmbPeriod.SelectedItem.ToString()
                                           select a).FirstOrDefault();

                            for (DateTime CurrentDay = oPeriod.StartDate.GetValueOrDefault(); CurrentDay <= oPeriod.EndDate; CurrentDay = CurrentDay.AddDays(1))
                            {
                                var oAttendance = (from a in odb.TrnsAttendanceRegister where a.EmpID == oEmp.ID && a.Date == CurrentDay select a).FirstOrDefault();
                                if (oAttendance is null)
                                {
                                    logger.Info($"Attendance register record not found for employee {oEmp.EmpID} at date {CurrentDay}");
                                    continue;
                                }

                                #region Leave Section
                                if (oAttendance.FlgIsNewLeave.GetValueOrDefault())
                                {
                                    var oRecord = (from a in odb.TrnsLeavesRequest
                                                   where a.EmpID == oEmp.ID
                                                   && a.LeaveType == oAttendance.LeaveType
                                                   && a.LeaveFrom <= CurrentDay
                                                   && a.LeaveTo >= CurrentDay
                                                   select a).FirstOrDefault();
                                    if (oRecord is null)
                                    {
                                        TrnsLeavesRequest oLeaveRec = new TrnsLeavesRequest();
                                        oLeaveRec.DocNum = ++LeaveDocCount;
                                        oLeaveRec.Series = -1;
                                        oLeaveRec.MstEmployee = oEmp;
                                        oLeaveRec.EmpName = oEmp.FirstName;
                                        oLeaveRec.DocDate = CurrentDay;
                                        oLeaveRec.MstLeaveType = oAttendance.MstLeaveType;
                                        oLeaveRec.LeaveDescription = oAttendance.MstLeaveType.Description;
                                        if (oAttendance.LeaveCount == 1)
                                        {
                                            oLeaveRec.UnitsID = "Day";
                                            oLeaveRec.UnitsLOVType = "LeaveUnits";
                                            oLeaveRec.Units = 0;
                                        }
                                        else
                                        {
                                            oLeaveRec.UnitsID = "HalfDay";
                                            oLeaveRec.UnitsLOVType = "LeaveUnits";
                                            oLeaveRec.Units = 0;
                                        }
                                        oLeaveRec.LeaveFrom = CurrentDay;
                                        oLeaveRec.LeaveTo = CurrentDay;
                                        oLeaveRec.TotalCount = oAttendance.LeaveCount;
                                        oLeaveRec.CalCode = oPeriod.CalCode;
                                        oLeaveRec.AttendanceID = oAttendance.Id;
                                        oLeaveRec.FlgPaid = false;
                                        oLeaveRec.CreatedBy = "DSKApp";
                                        oLeaveRec.CreateDate = DateTime.Now;
                                        oLeaveRec.UpdatedBy = "DSKApp";
                                        oLeaveRec.UpdateDate = DateTime.Now;
                                        odb.TrnsLeavesRequest.InsertOnSubmit(oLeaveRec);
                                    }


                                }
                                #endregion

                                #region OT Section
                                if (oAttendance.OTType > 0)
                                {
                                    TrnsEmployeeOvertime oOTRecord;
                                    TrnsEmployeeOvertimeDetail oOTRecordDetail;
                                    oOTRecordDetail = (from a in odb.TrnsEmployeeOvertimeDetail
                                                       where a.TrnsEmployeeOvertime.EmployeeId == oEmp.ID
                                                       && a.TrnsEmployeeOvertime.Period == oPeriod.ID
                                                       && a.OTDate == CurrentDay
                                                       select a).FirstOrDefault();
                                    if (oOTRecordDetail is null)
                                    {
                                        oOTRecord = new TrnsEmployeeOvertime();
                                        oOTRecordDetail = new TrnsEmployeeOvertimeDetail();

                                        oOTRecord.MstEmployee = oEmp;
                                        oOTRecord.CfgPeriodDates = oPeriod;
                                        oOTRecord.AttendanceID = oAttendance.Id;
                                        oOTRecord.UserId = "DSKApp";
                                        oOTRecord.CreateDate = DateTime.Now;
                                        oOTRecord.UpdatedBy = "DSKApp";
                                        oOTRecord.UpdateDate = DateTime.Now;
                                        oOTRecordDetail.MstOverTime = oAttendance.MstOverTime;
                                        oOTRecordDetail.ValueType = oAttendance.MstOverTime.ValueType;
                                        oOTRecordDetail.OTValue = oAttendance.MstOverTime.Value;
                                        oOTRecordDetail.OTDate = CurrentDay;
                                        oOTRecordDetail.FromTime = "00:00";
                                        oOTRecordDetail.ToTime = oAttendance.OTHour;
                                        oOTRecordDetail.OTHours = Convert.ToDecimal(TimeConvert(oAttendance.OTHour));
                                        oOTRecordDetail.Amount = OTAmountCalculate(oEmp, oAttendance.MstOverTime, oPeriod, Convert.ToDecimal(TimeConvert(oAttendance.OTHour)));
                                        oOTRecordDetail.FlgActive = true;
                                        oOTRecordDetail.UserId = "DSKApp";
                                        oOTRecordDetail.CreateDate = DateTime.Now;
                                        oOTRecordDetail.UpdatedBy = "DSKApp";
                                        oOTRecordDetail.UpdateDate = DateTime.Now;
                                        oOTRecord.TrnsEmployeeOvertimeDetail.Add(oOTRecordDetail);
                                        odb.TrnsEmployeeOvertime.InsertOnSubmit(oOTRecord);
                                    }
                                    else
                                    {
                                        logger.Info($"Overtime already entered for date {CurrentDay} of employee {oEmp.EmpID}");
                                    }
                                }
                                #endregion

                                #region Main Section

                                oAttendance.FlgPost = true;
                                oAttendance.FlgPosted = true;
                                oAttendance.UpdatedBy = "DSKApp";
                                oAttendance.UpdateDate = DateTime.Now;

                                #endregion  
                            }

                            odb.SubmitChanges();


                        }
                    }
                }
                lblStart.Text = "Posting attendance Completed.";
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                lblStart.Text = "Posting attendance Completed, with error";
            }
        }
        decimal OTAmountCalculate(MstEmployee oEmp, MstOverTime OvertimeMaster, CfgPeriodDates oPeriod, decimal OTHours)
        {
            try
            {
                int OTdays = 0;
                decimal WorkHour = 0;
                decimal CalculateOn = 0, PerDaySalary = 0, PerHourSalary = 0;
                if (string.IsNullOrEmpty(OvertimeMaster.Days))
                {
                    OTdays = Convert.ToInt32(oEmp.CfgPayrollDefination.WorkDays);
                    if (OTdays == 0)
                    {
                        var value = (oPeriod.EndDate - oPeriod.StartDate).Value.TotalDays + 1;
                        OTdays = Convert.ToInt32(value);
                    }
                }
                else
                {
                    OTdays = Convert.ToInt32(OvertimeMaster.Days);
                }
                if (string.IsNullOrEmpty(OvertimeMaster.Hours))
                {
                    WorkHour = Convert.ToDecimal(oEmp.CfgPayrollDefination.WorkHours);
                }
                else
                {
                    WorkHour = Convert.ToDecimal(OvertimeMaster.Hours);
                }
                if (OvertimeMaster.ValueType == "POB")
                {
                    CalculateOn = (oEmp.BasicSalary.GetValueOrDefault() / 100) * OvertimeMaster.Value.GetValueOrDefault();
                }
                else if (OvertimeMaster.ValueType == "POG")
                {
                    CalculateOn = (oEmp.GrossSalary.GetValueOrDefault() / 100) * OvertimeMaster.Value.GetValueOrDefault();
                }
                else
                {
                    CalculateOn = OvertimeMaster.Value.GetValueOrDefault();
                }

                //Per Day Salary
                PerDaySalary = CalculateOn / OTdays;
                PerHourSalary = PerDaySalary / WorkHour;

                return PerHourSalary * OTHours;
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return 0;
            }
        }
        void VoidAttendance()
        {
            try
            {
                lblStart.Text = "Void attendance started, Please wait.";
                using (dbHRMS odb = new dbHRMS(ConnectionString))
                {
                    if (dtEmployees.Rows.Count > 0)
                    {
                        var SelectedEmployees = (from a in dtEmployees.AsEnumerable()
                                                 where a.Field<bool>("Select") == true
                                                 select a).ToList();

                        foreach (var emp in SelectedEmployees)
                        {
                            var oEmp = (from a in odb.MstEmployee
                                        where a.EmpID == emp.Field<string>("EmpCode")
                                        select a).FirstOrDefault();
                            if (oEmp is null)
                            {
                                logger.Info($"employee {emp.Field<string>("EmpCode")} not found.");
                                continue;
                            }
                            var oPayroll = (from a in odb.CfgPayrollDefination
                                            where a.ID == oEmp.PayrollID
                                            select a).FirstOrDefault();
                            if (oPayroll is null)
                            {
                                logger.Info($"payroll not found. for employee {oEmp.EmpID}");
                                continue;
                            }
                            var oPeriod = (from a in odb.CfgPeriodDates
                                           where a.PayrollId == oPayroll.ID
                                           && a.PeriodName == cmbPeriod.SelectedItem.ToString()
                                           select a).FirstOrDefault();

                            for (DateTime CurrentDay = dtFrom.Value; CurrentDay <= dtTo.Value; CurrentDay = CurrentDay.AddDays(1))
                            {
                                var oAttendance = (from a in odb.TrnsAttendanceRegister where a.EmpID == oEmp.ID && a.Date == CurrentDay select a).FirstOrDefault();
                                if (oAttendance is null)
                                {
                                    logger.Info($"Attendance register record not found for employee {oEmp.EmpID} at date {CurrentDay}");
                                    continue;
                                }



                                #region Leave Section
                                if (oAttendance.FlgIsNewLeave.GetValueOrDefault())
                                {
                                    var oRecord = (from a in odb.TrnsLeavesRequest
                                                   where a.EmpID == oEmp.ID
                                                   && a.AttendanceID == oAttendance.Id
                                                   select a).FirstOrDefault();
                                    if (!(oRecord is null))
                                    {
                                        odb.TrnsLeavesRequest.DeleteOnSubmit(oRecord);
                                    }


                                }
                                #endregion

                                #region OT Section
                                if (oAttendance.OTType > 0)
                                {
                                    var oRecord = (from a in odb.TrnsEmployeeOvertime
                                                   where a.EmployeeId == oEmp.ID
                                                   && a.Period == oPeriod.ID
                                                   && a.AttendanceID == oAttendance.Id
                                                   select a).FirstOrDefault();
                                    if (!(oRecord is null))
                                    {
                                        odb.TrnsEmployeeOvertime.DeleteOnSubmit(oRecord);
                                    }
                                }
                                #endregion

                                #region Main Record
                                oAttendance.TimeIn = string.Empty;
                                oAttendance.TimeOut = string.Empty;
                                oAttendance.WorkHour = string.Empty;

                                oAttendance.EarlyInMin = string.Empty;
                                oAttendance.LateInMin = string.Empty;
                                oAttendance.EarlyOutMin = string.Empty;
                                oAttendance.LateOutMin = string.Empty;

                                oAttendance.FlgIsNewLeave = false;
                                oAttendance.LeaveHour = string.Empty;
                                oAttendance.LeaveCount = 0;
                                oAttendance.MstLeaveType = null;

                                oAttendance.OTHour = string.Empty;
                                oAttendance.MstOverTime = null;

                                oAttendance.Processed = false;
                                oAttendance.FlgPosted = false;
                                oAttendance.FlgSave = false;
                                oAttendance.FlgPost = false;

                                oAttendance.UpdateDate = DateTime.Now;
                                oAttendance.UpdatedBy = "DSKApp";
                                #endregion
                            }

                            odb.SubmitChanges();


                        }
                    }
                }
                lblStart.Text = "Void attendance Completed.";
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                lblStart.Text = "Void attendance Completed, with error";
            }
        }
        void LineProcess(int ProcessLine)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    string TimeIn, TimeOut, WorkHour;
                    int iTimeIn, iTimeOut, iWorkHour;
                    string ShiftIn, ShiftOut, ShiftHour, StartBuffer, EndBuffer;
                    int iShiftIn, iShiftOut, iShiftHour, iStartBuffer, iEndBuffer;
                    string EarlyIn = string.Empty, LateIn = string.Empty, EarlyOut = string.Empty, LateOut = string.Empty;
                    int iEarlyIn, iLateIn, iEarlyOut, iLateOut;
                    bool flgEarlyIn = false, flgLateIn = false, flgEarlyOut = false, flgLateOut = false;
                    string EmpCode, ShiftDesc, DateDay, StrDate;
                    DateTime CurrentDate = DateTime.Now;
                    bool flgNewLeave = false;
                    bool flgFullDayLeave = false;
                    string LeaveHour = string.Empty, LeaveType = string.Empty;
                    int LeaveTypeID = 0;
                    decimal LeaveCount = 0;
                    int OTId = 0;
                    string OTHours = string.Empty, OTType = string.Empty;
                    bool flgOT = false;
                    int ShortLeave = 0;
                    int GracePeriod = 0;
                    int PenaltyCount = 0;

                    EmpCode = Convert.ToString(dtProcessed.Rows[ProcessLine]["EmpCode"]);
                    ShiftDesc = Convert.ToString(dtProcessed.Rows[ProcessLine]["Shift"]);
                    DateDay = Convert.ToString(dtProcessed.Rows[ProcessLine]["Day"]);
                    //StrDate = Convert.ToString(dtProcessed.Rows[ProcessLine]["Date"]);
                    //CurrentDate = DateTime.ParseExact(StrDate, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None);
                    CurrentDate = Convert.ToDateTime(dtProcessed.Rows[ProcessLine]["Date"]);
                    var oEmp = (from a in odb.MstEmployee where a.EmpID == EmpCode select a).FirstOrDefault();
                    if (oEmp is null) { logger.Info($"employee not found,ProcessLine employee {EmpCode}"); return; }
                    var oShift = (from a in odb.MstShifts where a.Description == ShiftDesc select a).FirstOrDefault();
                    if (oShift is null) { logger.Info($"shift not found, ProcessLine employee {EmpCode}"); return; }

                    var oCalendar = (from a in odb.MstCalendar
                                     where a.FlgActive == true
                                     select a).FirstOrDefault();

                    if (oCalendar is null)
                    {
                        logger.Info("calendar not active or define.");
                        return;
                    }

                    #region GetEmployeeLeaves
                    var oEmpLeave = (from a in odb.MstEmployeeLeaves
                                     where a.EmpID == oEmp.ID
                                     && a.LeaveCalCode == oCalendar.Code
                                     && a.LeavesEntitled > 0
                                     select a).ToList();
                    List<LeaveStructure> oLeaveBal = new List<LeaveStructure>();
                    if (oEmpLeave is null)
                    {
                        logger.Info($"employee leave not assign, for employee {oEmp.EmpID}");
                    }
                    else
                    {
                        foreach (var leave in oEmpLeave)
                        {
                            LeaveStructure oDoc = new LeaveStructure();
                            oDoc.ID = leave.LeaveType ?? 0;
                            oDoc.LeaveType = leave.MstLeaveType.Description;
                            oDoc.Balance = leave.LeavesEntitled + leave.LeavesCarryForward;
                            switch (oDoc.ID)
                            {
                                case 10:
                                    oDoc.Priority = 10; break;
                                case 3:
                                    oDoc.Priority = 1; break;
                                case 1:
                                    oDoc.Priority = 2; break;
                                case 2:
                                    oDoc.Priority = 3; break;
                                case 7:
                                    oDoc.Priority = 4; break;
                                case 6:
                                    oDoc.Priority = 5; break;
                                case 5:
                                    oDoc.Priority = 6; break;
                                case 4:
                                    oDoc.Priority = 7; break;

                            }
                            oLeaveBal.Add(oDoc);
                        }
                    }
                    var oAttendanceRule = (from a in odb.MstAttendanceRule
                                           select a).FirstOrDefault();
                    #endregion

                    #region Time Value

                    TimeIn = Convert.ToString(dtProcessed.Rows[ProcessLine]["In"]);
                    TimeOut = Convert.ToString(dtProcessed.Rows[ProcessLine]["Out"]);

                    iTimeIn = TimeConvert(TimeIn);
                    iTimeOut = TimeConvert(TimeOut);
                    if (iTimeIn == 0)
                    {
                        TimeIn = string.Empty;
                    }
                    if (iTimeOut == 0)
                    {
                        TimeOut = string.Empty;
                    }
                    if (iTimeIn > 0 && iTimeOut > 0)
                    {
                        iWorkHour = iTimeOut - iTimeIn;
                        WorkHour = TimeConvert(iWorkHour);
                    }
                    else
                    {
                        iWorkHour = 0;
                        WorkHour = string.Empty;
                    }

                    #endregion

                    #region Shift
                    ShiftIn = Convert.ToString(dtProcessed.Rows[ProcessLine]["ShiftIn"]);
                    ShiftOut = Convert.ToString(dtProcessed.Rows[ProcessLine]["ShiftOut"]);
                    ShiftHour = Convert.ToString(dtProcessed.Rows[ProcessLine]["ShiftDuration"]);
                    iShiftIn = TimeConvert(ShiftIn);
                    iShiftOut = TimeConvert(ShiftOut);
                    iShiftHour = TimeConvert(ShiftHour);
                    StartBuffer = oShift.MstShiftDetails.Where(a => a.Day == DateDay).FirstOrDefault().BufferStartTime ?? "";
                    EndBuffer = oShift.MstShiftDetails.Where(a => a.Day == DateDay).FirstOrDefault().BufferEndTime ?? "";
                    iStartBuffer = TimeConvert(StartBuffer);
                    iEndBuffer = TimeConvert(EndBuffer);
                    #endregion

                    #region Early/Late Ins

                    if (iTimeIn > 0)
                    {
                        if (iTimeIn - iShiftIn > 0)
                        {
                            LateIn = TimeConvert(iTimeIn - iShiftIn);
                            iLateIn = TimeConvert(LateIn);
                            flgLateIn = true;
                        }
                        if (iShiftIn - iTimeIn > 0)
                        {
                            EarlyIn = TimeConvert(iShiftIn - iTimeIn);
                            iEarlyIn = TimeConvert(EarlyIn);
                            flgEarlyIn = true;
                        }
                    }
                    if (iTimeOut > 0)
                    {
                        if (iTimeOut - iShiftOut > 0)
                        {
                            LateOut = TimeConvert(iTimeOut - iShiftOut);
                            iLateOut = TimeConvert(LateOut);
                            flgLateOut = true;
                        }
                        if (iShiftOut - iTimeOut > 0)
                        {
                            EarlyOut = TimeConvert(iShiftOut - iTimeOut);
                            iEarlyOut = TimeConvert(EarlyOut);
                            flgEarlyOut = true;
                        }
                    }

                    #endregion

                    #region Penalty & Leaves

                    //full day leave
                    if (iTimeIn == 0 && iTimeOut == 0 && iShiftHour > 0)
                    {
                        var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                          where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                          && a.EmpID == oEmp.ID
                                          select a).Count();
                        if (LeaveCheck == 0)
                        {
                            var oLT = (from a in oLeaveBal
                                       where a.Balance > 0
                                       && a.LeaveType.ToLower() == "absent"
                                       select a).FirstOrDefault();
                            flgNewLeave = true;
                            LeaveHour = ShiftHour;
                            LeaveType = oLT.LeaveType;
                            LeaveTypeID = oLT.ID;
                            LeaveCount = 1;
                            oLT.Balance -= 1;
                        }
                        else
                        {
                            var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                 where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                                 && a.EmpID == oEmp.ID
                                                 select a).FirstOrDefault();
                            flgNewLeave = false;
                            LeaveHour = ShiftHour;
                            LeaveType = oLeaveRequest.MstLeaveType.Description;
                            LeaveCount = 1;
                            LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                        }
                        flgFullDayLeave = true;
                    }

                    //Missing timein or timeout
                    if ((iTimeIn > 0 && iTimeOut == 0 && iShiftHour > 0) || (iTimeIn == 0 && iTimeOut > 0 && iShiftHour > 0))
                    {
                        var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                          where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                          && a.EmpID == oEmp.ID
                                          select a).Count();
                        if (LeaveCheck == 0)
                        {
                            var oLT = (from a in oLeaveBal
                                       where a.Balance > 0
                                       && a.LeaveType.ToLower() == "absent"
                                       select a).FirstOrDefault();
                            flgNewLeave = true;
                            LeaveHour = ShiftHour;
                            LeaveType = oLT.LeaveType;
                            LeaveTypeID = oLT.ID;
                            LeaveCount = 1;
                            oLT.Balance -= 1;
                        }
                        else
                        {
                            var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                 where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                                 && a.EmpID == oEmp.ID
                                                 select a).FirstOrDefault();
                            flgNewLeave = false;
                            LeaveHour = ShiftHour;
                            LeaveType = oLeaveRequest.MstLeaveType.Description;
                            LeaveCount = oLeaveRequest.TotalCount.GetValueOrDefault();
                            LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                        }
                    }

                    //if workhour and shift hour is ok
                    ////but not on shift time.
                    if (iShiftHour <= iWorkHour && iShiftHour > 0)
                    {
                        if ((flgLateIn && flgEarlyOut) || (!flgLateIn && flgEarlyOut) || (flgLateIn && !flgEarlyOut))
                        {
                            string TimeDiff = TimeConvert(TimeConvert(LateIn) + TimeConvert(EarlyOut));
                            var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                              where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                              && a.EmpID == oEmp.ID
                                              select a).Count();
                            string DeductionRule = CalculateDeductionRule(TimeConvert(TimeDiff));
                            int GracePeriodDR = ReturnGracePeriodDeductionRule(DeductionRule);

                            if (DeductionRule == "DR_01")
                            {
                                ShortLeave++;
                                GracePeriod++;
                            }
                            if (LeaveCheck == 0)
                            {
                                if (DeductionRule == "DR_01")
                                {
                                    if (GracePeriod > GracePeriodDR)
                                    {

                                        var oLT = (from a in oLeaveBal
                                                   where a.Balance > 0
                                                   orderby a.Priority ascending
                                                   select a).FirstOrDefault();
                                        if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); }
                                        flgNewLeave = true;
                                        LeaveHour = TimeDiff;
                                        LeaveType = oLT.LeaveType;
                                        LeaveTypeID = oLT.ID;
                                        LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                        oLT.Balance -= 1;
                                    }
                                    else
                                    {
                                        logger.Info($"deduction rule 1 grace period {GracePeriod} date {CurrentDate.ToString("dd MMM yyyy")} employee {oEmp.EmpID} no deduction.");
                                    }
                                }
                                else if (DeductionRule == "DR_02")
                                {
                                    var oLT = (from a in oLeaveBal
                                               where a.Balance > 0
                                               orderby a.Priority ascending
                                               select a).FirstOrDefault();
                                    if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); }
                                    flgNewLeave = true;
                                    LeaveHour = TimeDiff;
                                    LeaveType = oLT.LeaveType;
                                    LeaveTypeID = oLT.ID;
                                    LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                    oLT.Balance -= 1;
                                }
                                else if (DeductionRule == "DR_03")
                                {
                                    var oLT = (from a in oLeaveBal
                                               where a.Balance > 0
                                               orderby a.Priority ascending
                                               select a).FirstOrDefault();
                                    if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); }
                                    flgNewLeave = true;
                                    LeaveHour = TimeDiff;
                                    LeaveType = oLT.LeaveType;
                                    LeaveTypeID = oLT.ID;
                                    LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                    oLT.Balance -= 1;
                                }
                            }
                            else
                            {
                                var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                     where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                                     && a.EmpID == oEmp.ID
                                                     select a).FirstOrDefault();
                                flgNewLeave = false;
                                LeaveHour = TimeDiff;
                                LeaveType = oLeaveRequest.MstLeaveType.Description;
                                LeaveCount = oLeaveRequest.TotalCount.GetValueOrDefault();
                                LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                            }
                        }
                    }

                    //if workhour is less then shift hour
                    if (iShiftHour > iWorkHour && iTimeIn > 0 && iTimeOut > 0)
                    {
                        string TimeDiff = TimeConvert(iShiftHour - iWorkHour);
                        var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                          where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                          && a.EmpID == oEmp.ID
                                          select a).Count();
                        string DeductionRule = CalculateDeductionRule(TimeConvert(TimeDiff));
                        int GracePeriodDR = ReturnGracePeriodDeductionRule(DeductionRule);

                        if (DeductionRule == "DR_01")
                        {
                            ShortLeave++;
                            GracePeriod++;
                        }
                        if (LeaveCheck == 0)
                        {
                            if (DeductionRule == "DR_01")
                            {
                                if (GracePeriod > GracePeriodDR)
                                {
                                    logger.Info($"deduction rule 1 grace period {GracePeriod} date {CurrentDate.ToString("dd MMM yyyy")} employee {oEmp.EmpID} deducted.");
                                    var oLT = (from a in oLeaveBal
                                               where a.Balance > 0
                                               orderby a.Priority ascending
                                               select a).FirstOrDefault();
                                    if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); }
                                    flgNewLeave = true;
                                    LeaveHour = TimeDiff;
                                    LeaveType = oLT.LeaveType;
                                    LeaveTypeID = oLT.ID;
                                    LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                    oLT.Balance -= 1;
                                }
                                else
                                {
                                    logger.Info($"deduction rule 1 grace period {GracePeriod} date {CurrentDate.ToString("dd MMM yyyy")} employee {oEmp.EmpID} no deduction.");
                                }
                            }
                            else if (DeductionRule == "DR_02")
                            {
                                var oLT = (from a in oLeaveBal
                                           where a.Balance > 0
                                           orderby a.Priority ascending
                                           select a).FirstOrDefault();
                                if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); }
                                flgNewLeave = true;
                                LeaveHour = TimeDiff;
                                LeaveType = oLT.LeaveType;
                                LeaveTypeID = oLT.ID;
                                LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                oLT.Balance -= 1;
                            }
                            else if (DeductionRule == "DR_03")
                            {
                                var oLT = (from a in oLeaveBal
                                           where a.Balance > 0
                                           orderby a.Priority ascending
                                           select a).FirstOrDefault();
                                if (oLT == null) { RadMessageBox.Show("Check leave type & balance."); }
                                flgNewLeave = true;
                                LeaveHour = TimeDiff;
                                LeaveType = oLT.LeaveType;
                                LeaveTypeID = oLT.ID;
                                LeaveCount = CalculateDeductionCount(TimeConvert(TimeDiff));
                                oLT.Balance -= 1;
                            }
                        }
                        else
                        {
                            var oLeaveRequest = (from a in odb.TrnsLeavesRequest
                                                 where a.LeaveFrom <= CurrentDate && a.LeaveTo >= CurrentDate
                                                 && a.EmpID == oEmp.ID
                                                 select a).FirstOrDefault();
                            flgNewLeave = false;
                            LeaveHour = TimeDiff;
                            LeaveType = oLeaveRequest.MstLeaveType.Description;
                            LeaveCount = oLeaveRequest.TotalCount.GetValueOrDefault();
                            LeaveTypeID = oLeaveRequest.LeaveType ?? 0;
                        }
                    }

                    //Missing In/Out Penalty
                    if ((iTimeIn == 0 || iTimeOut == 0) && iShiftHour > 0 && !flgFullDayLeave)
                    {
                        string DeductionRule = string.Empty;

                        if (oAttendanceRule.FlgMissingTimePenalty.GetValueOrDefault())
                        {
                            DeductionRule = "DR_04";
                            PenaltyCount += 1;
                        }
                        if (PenaltyCount >= oAttendanceRule.MissingTimePanaltyCounter.GetValueOrDefault())
                        {
                            PenaltyCount = 0;
                            MstLeaveType oLeaveMaster = null;
                            if (DeductionRule == "DR_04")
                            {
                                oLeaveMaster = (from a in odb.MstLeaveType where a.Code == oAttendanceRule.MPTLeaveType select a).FirstOrDefault();
                            }
                            else
                            {
                                oLeaveMaster = (from a in odb.MstLeaveType where a.Code == oAttendanceRule.LeaveTypeWOP select a).FirstOrDefault();
                            }
                            if (oLeaveMaster != null)
                            {
                                flgNewLeave = true;
                                LeaveHour = ShiftHour;
                                LeaveType = oLeaveMaster.Description;
                                LeaveTypeID = oLeaveMaster.ID;
                                LeaveCount = 0.5M;
                            }

                        }
                    }


                    #endregion

                    #region Overtime

                    if (iWorkHour > iShiftHour)
                    {
                        //checking timeout is also more than 
                        //expected
                        if (flgLateOut)
                        {
                            if (iEndBuffer < TimeConvert(LateOut))
                            {
                                flgOT = true;
                                OTHours = LateOut;
                                OTType = oShift.MstOverTime == null ? "" : oShift.MstOverTime.Description;
                                OTId = oShift.MstOverTime == null ? 0 : oShift.MstOverTime.ID;
                            }
                        }
                    }

                    #endregion

                    #region Set Datatable

                    DataRow dr = dtProcessed.NewRow();
                    //dr["Serial"] = Serial;
                    //dr["EmpCode"] = oEmp.EmpID;
                    //dr["EmpName"] = $"{oEmp.FirstName} {oEmp.MiddleName} {oEmp.LastName}";
                    //dr["ID"] = oAttendance.Id;
                    //dr["Date"] = oAttendance.Date;
                    //dr["Day"] = oAttendance.Date.Value.DayOfWeek.ToString();
                    //dr["Shift"] = oAttendance.MstShifts.Description;
                    //dr["ShiftIn"] = ShiftIn;
                    //dr["ShiftOut"] = ShiftOut;
                    //dr["ShiftDuration"] = ShiftDuration;
                    dtProcessed.Rows[ProcessLine]["In"] = TimeIn;
                    dtProcessed.Rows[ProcessLine]["Out"] = TimeOut;
                    dtProcessed.Rows[ProcessLine]["WorkHour"] = WorkHour;
                    dtProcessed.Rows[ProcessLine]["EarlyIn"] = EarlyIn;
                    dtProcessed.Rows[ProcessLine]["LateIn"] = LateIn;
                    dtProcessed.Rows[ProcessLine]["EarlyOut"] = EarlyOut;
                    dtProcessed.Rows[ProcessLine]["LateOut"] = LateOut;
                    dtProcessed.Rows[ProcessLine]["LeaveHour"] = LeaveHour;
                    dtProcessed.Rows[ProcessLine]["LeaveType"] = LeaveType;
                    dtProcessed.Rows[ProcessLine]["LeaveNew"] = flgNewLeave;
                    dtProcessed.Rows[ProcessLine]["LeaveCount"] = LeaveCount;
                    dtProcessed.Rows[ProcessLine]["LTID"] = LeaveTypeID;
                    dtProcessed.Rows[ProcessLine]["OTHour"] = OTHours;
                    dtProcessed.Rows[ProcessLine]["OTType"] = OTType;
                    dtProcessed.Rows[ProcessLine]["OTID"] = OTId;
                    //dr["Status"] = "Draft";
                    //Serial++;
                    //dtProcessed.Rows.Add(dr);
                    #endregion
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        void ImportTempData()
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {

                    if (dtEmployees.Rows.Count > 0)
                    {
                        var SelectedEmployees = (from a in dtEmployees.AsEnumerable()
                                                 where a.Field<bool>("Select") == true
                                                 select a).ToList();
                        foreach (var emp in SelectedEmployees)
                        {

                            var oEmp = (from a in odb.MstEmployee
                                        where a.EmpID == emp.Field<string>("EmpCode")
                                        select a).FirstOrDefault();

                            if (oEmp is null) { logger.Info("employee not found."); continue; }

                            var oPayroll = (from a in odb.CfgPayrollDefination
                                            where a.ID == oEmp.PayrollID
                                            select a).FirstOrDefault();

                            if (oPayroll is null) { logger.Info($"payroll not found. for employee {oEmp.EmpID}"); continue; }

                            var oPeriod = (from a in odb.CfgPeriodDates
                                           where a.PayrollId == oPayroll.ID
                                           && a.PeriodName == cmbPeriod.SelectedItem.ToString()
                                           select a).FirstOrDefault();

                            if (oPeriod is null) { logger.Info($"period not found. for employee {oEmp.EmpID}"); continue; }

                            DateTime loopStart = oPeriod.StartDate.GetValueOrDefault();
                            DateTime loopEnd = oPeriod.EndDate.GetValueOrDefault();
                            for (DateTime i = loopStart; i <= loopEnd; i = i.AddDays(1))
                            {

                                var oCheck = (from a in odb.TrnsTempAttendance
                                              where a.PunchedDate == i
                                              && a.EmpID == oEmp.EmpID
                                              select a).Count();
                                if (oCheck == 0)
                                {
                                    string strQuery =
                                        $"select a.CardNo AS EmployeeCode, CAST( a.PunchDatetime AS DATE) AS PunchedDate, CAST(CAST(a.PunchDatetime AS TIME) AS NVARCHAR(5)) AS PunchedTime, Case When Isnull(a.inout, 'In') = 'In' Then '1' Else '2' End AS PunchedType from Tran_MachineRawPunch as a where a.CardNo = {oEmp.EmpID} and CAST(a.PunchDatetime as date) = '{i.ToString("yyyy-MM-dd")}'";
                                    using (SqlConnection connection = new SqlConnection(AttConnectionString))
                                    {
                                        connection.Open();
                                        SqlCommand command = connection.CreateCommand();
                                        command.CommandText = strQuery;
                                        using (var reader = command.ExecuteReader())
                                        {
                                            while (reader.Read())
                                            {
                                                TrnsTempAttendance oRec = new TrnsTempAttendance();
                                                oRec.EmpID = oEmp.EmpID;
                                                oRec.PunchedDate = i;
                                                oRec.PunchedTime = Convert.ToString(reader["PunchedTime"]);
                                                oRec.In_Out = Convert.ToString(reader["PunchedType"]);
                                                oRec.UserID = "Auto";
                                                oRec.CreatedDate = DateTime.Now;
                                                oRec.FlgProcessed = false;
                                                //extra
                                                oRec.CostCenter = "";
                                                oRec.PolledDate = i;
                                                oRec.PunchedDateTime = i;
                                                odb.TrnsTempAttendance.InsertOnSubmit(oRec);
                                            }
                                        }
                                        connection.Close();
                                    }
                                }
                                else
                                {
                                    string strDelQuery = $"DELETE FROM dbo.TrnsTempAttendance WHERE EmpID = '{oEmp.EmpID}' AND CAST(PunchedDateTime AS DATE) = '{i.ToString("yyyy-MM-dd")}'";
                                    using (SqlConnection connection = new SqlConnection(ConnectionString))
                                    {
                                        connection.Open();
                                        SqlCommand command = connection.CreateCommand();
                                        command.CommandText = strDelQuery;
                                        command.ExecuteNonQuery();
                                        connection.Close();
                                    }
                                    string strQuery =
                                        $"select a.CardNo AS EmployeeCode, CAST( a.PunchDatetime AS DATE) AS PunchedDate, CAST(CAST(a.PunchDatetime AS TIME) AS NVARCHAR(5)) AS PunchedTime, Case When Isnull(a.inout, 'In') = 'In' Then '1' Else '2' End AS PunchedType from Tran_MachineRawPunch as a where a.CardNo = {oEmp.EmpID} and CAST(a.PunchDatetime as date) = '{i.ToString("yyyy-MM-dd")}'";

                                    using (SqlConnection connection = new SqlConnection(AttConnectionString))
                                    {
                                        connection.Open();
                                        SqlCommand command = connection.CreateCommand();
                                        command.CommandText = strQuery;
                                        using (var reader = command.ExecuteReader())
                                        {
                                            while (reader.Read())
                                            {
                                                TrnsTempAttendance oRec = new TrnsTempAttendance();
                                                oRec.EmpID = oEmp.EmpID;
                                                oRec.PunchedDate = i;
                                                oRec.PunchedTime = Convert.ToString(reader["PunchedTime"]);
                                                oRec.In_Out = Convert.ToString(reader["PunchedType"]);
                                                oRec.UserID = "Auto";
                                                oRec.CreatedDate = DateTime.Now;
                                                oRec.FlgProcessed = false;
                                                //extra
                                                oRec.CostCenter = "";
                                                oRec.PolledDate = i;
                                                oRec.PunchedDateTime = i;
                                                odb.TrnsTempAttendance.InsertOnSubmit(oRec);
                                            }
                                        }
                                        connection.Close();
                                    }
                                }
                            }
                        }

                        odb.SubmitChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        void GridToggle()
        {
            try
            {
                if (GridOneToggle)
                {
                    foreach (DataRow row in dtEmployees.Rows)
                    {
                        row[0] = !GridOneToggle;
                    }
                    GridOneToggle = false;
                }
                else
                {
                    foreach (DataRow row in dtEmployees.Rows)
                    {
                        row[0] = !GridOneToggle;
                    }
                    GridOneToggle = true;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        void ImportTempAttendance()
        {
            try
            {

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        void SendSalarySlip(int SlipId, string PeriodName, DateTime PeriodStartDate, string EmpName, string EmpEmail)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    TblRpts oReport = (from a in odb.TblRpts where a.RptCode == "slip" select a).FirstOrDefault();
                    if (oReport is null)
                    {
                        logger.Log(LogLevel.Warn, "report didn't found slip.");
                        return;
                    }
                    byte[] reportinbytes = oReport.RptFileStr.ToArray();
                    string filenamereport = Application.StartupPath + "\\SlipRpt.rpt";
                    using (FileStream fs = new FileStream(filenamereport, FileMode.Create))
                    {
                        int filelenght = reportinbytes.Length;
                        fs.Write(reportinbytes, 0, filelenght);
                        fs.Flush();
                        fs.Close();
                    }
                    string filenameexport = Application.StartupPath + $"\\Exports\\{SlipId}.pdf";
                    using (ReportDocument oReportDoc = new ReportDocument())
                    {


                        oReportDoc.Load(filenamereport);
                        //Set databasevalues.
                        SetReport(oReportDoc);
                        ParameterFieldDefinitions fielddefination = oReportDoc.DataDefinition.ParameterFields;
                        ParameterDiscreteValue prmValue = new ParameterDiscreteValue();
                        ParameterFieldDefinition prmDef = fielddefination["Critaria"];
                        ParameterValues prmCollection = new ParameterValues();
                        prmCollection = prmDef.CurrentValues;
                        prmValue.Value = string.Format(" Where TrnsSalaryProcessRegister.Id = {0}", SlipId);
                        prmCollection.Add(prmValue);
                        prmDef.ApplyCurrentValues(prmCollection);

                        ExportOptions CrExportOptions;
                        DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                        PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                        CrDiskFileDestinationOptions.DiskFileName = filenameexport;
                        CrExportOptions = oReportDoc.ExportOptions;
                        CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                        CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                        CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                        CrExportOptions.FormatOptions = CrFormatTypeOptions;
                        oReportDoc.Export();
                    }

                    var email = odb.MstEmailConfig.FirstOrDefault();
                    string MonthName = Convert.ToDateTime(PeriodStartDate).ToString("MMMM", CultureInfo.InvariantCulture);
                    string YearName = Convert.ToDateTime(PeriodStartDate).ToString("yyyy", CultureInfo.InvariantCulture);

                    MailMessage mail = new MailMessage(email.FromEmail, EmpEmail,
                    string.Format("Salary Slip for month of {0}", MonthName),
                    string.Format(@"<b>Dear {0}</b>,
                                                <br> 
                                                <br> 
                                                Please find attached your pay-slip, for the month of {1} {2}.<br><br>
                                                In case of any ambiguity, please consult HR Department for resolution of your queries.<br><br>
                                                <b>This is a SAP generated pay-slip and does not require any signature.</b><br>
                                                <br>
                                                <b>Best Regards,</b><br>
                                                <br>
                                                <b>HR Department.</b>", EmpName, MonthName, YearName));
                    mail.IsBodyHtml = true;
                    try
                    {
                        Attachment attachment = new Attachment(filenameexport);
                        attachment.Name = PeriodName + " " + EmpName + ".pdf";  // set name here
                        mail.Attachments.Add(attachment);
                        SmtpClient client = new SmtpClient(email.SMTPServer);
                        client.Port = Convert.ToInt32(email.SMTPort);
                        client.Credentials = new System.Net.NetworkCredential(email.FromEmail, email.Password);
                        if (Convert.ToBoolean(email.SSL) == true)
                        {
                            client.EnableSsl = true;
                        }
                        else
                        {
                            client.EnableSsl = false;
                        }
                        client.Send(mail);
                        mail.Dispose();
                        File.Delete(filenameexport);
                    }
                    catch (Exception ex)
                    {
                        mail.Dispose();
                        logger.Error(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        void SendAttendanceSlip(int PeriodId, string PeriodName, string EmpId, string EmpName, string EmpEmail, DateTime PeriodStartDate)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    string filename = Application.StartupPath + $"\\Report\\AttendanceReport.rpt";
                    string filenameexport = Application.StartupPath + $"\\Exports\\{EmpId}.pdf";
                    using (ReportDocument oReportDoc = new ReportDocument())
                    {
                        oReportDoc.Load(filename);
                        SetReport(oReportDoc);
                        ParameterFieldDefinitions fielddefination = oReportDoc.DataDefinition.ParameterFields;
                        ParameterDiscreteValue prmValue = new ParameterDiscreteValue();
                        ParameterFieldDefinition prmDef = fielddefination["Critaria"];
                        ParameterValues prmCollection = new ParameterValues();
                        prmCollection = prmDef.CurrentValues;
                        prmValue.Value = string.Format(" WHERE  A1.PeriodID= {0} and A2.EmpID= '{1}' ", PeriodId, EmpId);
                        prmCollection.Add(prmValue);
                        prmDef.ApplyCurrentValues(prmCollection);

                        ExportOptions CrExportOptions;
                        DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                        PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                        CrDiskFileDestinationOptions.DiskFileName = filenameexport;
                        CrExportOptions = oReportDoc.ExportOptions;
                        CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                        CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                        CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                        CrExportOptions.FormatOptions = CrFormatTypeOptions;
                        oReportDoc.Export();
                    }

                    var email = odb.MstEmailConfig.FirstOrDefault();
                    string MonthName = Convert.ToDateTime(PeriodStartDate).ToString("MMMM", CultureInfo.InvariantCulture);
                    string YearName = Convert.ToDateTime(PeriodStartDate).ToString("yyyy", CultureInfo.InvariantCulture);

                    MailMessage mail = new MailMessage(email.FromEmail, EmpEmail,
                    string.Format("Attendance Slip for month of {0}", MonthName),
                    string.Format(@"<b>Dear {0}</b>,
                                                <br> 
                                                <br> 
                                                Please find attached your Attendance-slip, for the month of {1} {2}.<br><br>
                                                In case of any ambiguity, please consult HR Department for resolution of your queries.<br><br>
                                                <b>This is a SAP generated Attendance-slip and does not require any signature.</b><br>
                                                <br>
                                                <b>Best Regards,</b><br>
                                                <br>
                                                <b>HR Department.</b>", EmpName, MonthName, YearName));
                    mail.IsBodyHtml = true;
                    try
                    {
                        Attachment attachment = new Attachment(filenameexport);
                        attachment.Name = PeriodName + " " + EmpName + ".pdf";  // set name here
                        mail.Attachments.Add(attachment);
                        SmtpClient client = new SmtpClient(email.SMTPServer);
                        client.Port = Convert.ToInt32(email.SMTPort);
                        client.Credentials = new System.Net.NetworkCredential(email.FromEmail, email.Password);
                        if (Convert.ToBoolean(email.SSL) == true)
                        {
                            client.EnableSsl = true;
                        }
                        else
                        {
                            client.EnableSsl = false;
                        }
                        client.Send(mail);
                        mail.Dispose();
                        File.Delete(filename);
                    }
                    catch (Exception ex)
                    {
                        mail.Dispose();
                        logger.Error(ex);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }

        }
        void SetReport(ReportDocument rep)
        {
            foreach (Table Table in rep.Database.Tables)
            {
                #region Dynamic Report Credentials
                ConnectionInfo coninfo = new ConnectionInfo();
                coninfo.ServerName = Settings.Default.ServerName;
                coninfo.DatabaseName = Settings.Default.Database;
                coninfo.UserID = Settings.Default.DBUser;
                coninfo.Password = Settings.Default.DBPassword;
                TableLogOnInfo info = new TableLogOnInfo();

                info.ConnectionInfo = coninfo;
                for (int i = 0; i < rep.Database.Tables.Count; i++)
                {
                    rep.Database.Tables[i].ApplyLogOnInfo(info);
                    rep.Refresh();
                }
                #endregion               
            }

            foreach (ReportDocument rpt in rep.Subreports)
            {
                foreach (Table Table in rpt.Database.Tables)
                {
                    #region Dynamic Report Credentials
                    ConnectionInfo coninfo = new ConnectionInfo();
                    coninfo.ServerName = Settings.Default.ServerName;
                    coninfo.DatabaseName = Settings.Default.Database;
                    coninfo.UserID = Settings.Default.DBUser;
                    coninfo.Password = Settings.Default.DBPassword;
                    TableLogOnInfo info = new TableLogOnInfo();

                    info.ConnectionInfo = coninfo;
                    for (int i = 0; i < rep.Database.Tables.Count; i++)
                    {
                        rep.Database.Tables[i].ApplyLogOnInfo(info);
                        rep.Refresh();
                    }
                    #endregion

                }
            }
        }
        public bool ValidateLicense(string key)
        {
            try
            {
                bool resultValue = false;
                string result = mFm.mfmVerifyLicense(key);
                //result = "Verification Succeded";
                switch (result)
                {
                    case "Verification Succeded":
                        resultValue = true;
                        break;
                    default:
                        resultValue = false;
                        break;
                }
                if (!resultValue)
                {
                    logger.Info("License expires. renew you license.");
                }
                else
                {
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                logger.Error(ex);
                return false;
            }
        }
        void CallAttendanceReport()
        {
            try
            {
                using (var db = new dbHRMS(ConnectionString))
                {
                    var EmployeeList = (from a in dtEmployees.AsEnumerable()
                                        where a.Field<bool>("Select") == true
                                        select a.Field<string>("EmpCode")).ToList();
                    if (EmployeeList.Count > 0)
                    {
                        string payrollvalue = cmbPayroll.SelectedItem.ToString();
                        string periodvalue = cmbPeriod.SelectedItem.ToString();
                        var oPeriod = (from a in db.CfgPeriodDates
                                       where a.PeriodName == periodvalue
                                       && a.CfgPayrollDefination.PayrollName == payrollvalue
                                       select a).FirstOrDefault();

                        string[] arr = EmployeeList.ToArray();
                        string formatted = string.Join(", ", arr.Select(s => $"'{s}'"));

                        frmReportViewer oDialog = new frmReportViewer();
                        oDialog.ReportCode = 1;
                        oDialog.EmpCode = "";
                        oDialog.EmpList = formatted;
                        oDialog.PeriodCode = oPeriod.ID.ToString();
                        oDialog.ShowDialog();
                    }
                }

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        void CallTempAttendanceReport()
        {
            try
            {
                using (var db = new dbHRMS(ConnectionString))
                {
                    var EmployeeList = (from a in dtEmployees.AsEnumerable()
                                        where a.Field<bool>("Select") == true
                                        select a.Field<string>("EmpCode")).ToList();
                    if (EmployeeList.Count > 0)
                    {
                        string payrollvalue = cmbPayroll.SelectedItem.ToString();
                        string periodvalue = cmbPeriod.SelectedItem.ToString();
                        var oPeriod = (from a in db.CfgPeriodDates
                                       where a.PeriodName == periodvalue
                                       && a.CfgPayrollDefination.PayrollName == payrollvalue
                                       select a).FirstOrDefault();

                        string[] arr = EmployeeList.ToArray();
                        string formatted = string.Join(", ", arr.Select(s => $"'{s}'"));

                        frmReportViewer oDialog = new frmReportViewer();
                        oDialog.ReportCode = 2;
                        oDialog.EmpCode = "";
                        oDialog.EmpList = formatted;
                        oDialog.PeriodCode = oPeriod.ID.ToString();
                        oDialog.ShowDialog();
                    }
                }

            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        #endregion

        #region Events

        public frmMain()
        {
            InitializeComponent();
        }
        private void btnProcess_Click(object sender, EventArgs e)
        {
            try
            {
                if (btnProcess.Text == "Process")
                {
                    btnProcess.Text = "Back";
                    StartProcess();
                }
                else
                {
                    btnProcess.Text = "Process";
                    grdProcess.Visible = false;
                    grdEmployee.Visible = true;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void timer_Tick(object sender, EventArgs e)
        {
            try
            {
                //Task.Run(async () =>
                //{
                //    await StartProcess();
                //});
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(Properties.Settings.Default.ServerName))
                {
                    if (!string.IsNullOrEmpty(Properties.Settings.Default.Database))
                    {
                        if (!string.IsNullOrEmpty(Properties.Settings.Default.DBUser))
                        {
                            if (!string.IsNullOrEmpty(Properties.Settings.Default.DBPassword))
                            {
                                ConnectionString = $"Server={Properties.Settings.Default.ServerName};Database={Properties.Settings.Default.Database};User Id={Properties.Settings.Default.DBUser};Password={Properties.Settings.Default.DBPassword};";
                            }
                        }
                    }
                }
                if (!string.IsNullOrEmpty(Properties.Settings.Default.AttDbServer))
                {
                    if (!string.IsNullOrEmpty(Properties.Settings.Default.AttDbDatabase))
                    {
                        if (!string.IsNullOrEmpty(Properties.Settings.Default.AttDbUser))
                        {
                            if (!string.IsNullOrEmpty(Properties.Settings.Default.AttDbPassword))
                            {
                                AttConnectionString = $"Server={Properties.Settings.Default.AttDbServer};Database={Properties.Settings.Default.AttDbDatabase};User Id={Properties.Settings.Default.AttDbUser};Password={Properties.Settings.Default.AttDbPassword};";
                            }
                        }
                    }
                }
                CreateGrid();
                FillCombo();
                grdProcess.Visible = false;
                grdEmployee.Visible = true;
                grdProcess.Size = new System.Drawing.Size(1390, 480);
                grdProcess.Location = new System.Drawing.Point(19, 260);
                grdProcess.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right | System.Windows.Forms.AnchorStyles.Bottom;
                grdEmployee.Size = new System.Drawing.Size(1390, 480);
                grdEmployee.Location = new System.Drawing.Point(19, 260);
                grdEmployee.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right | System.Windows.Forms.AnchorStyles.Bottom;

                logger.Info("System hardware key: {0}", mFm.mfmGetSystemID());
                if (!ValidateLicense(Settings.Default.LicenseKey))
                {
                    RadMessageBox.Show("Kindly contact AnsaarSoft, @ mfmlive@gmail.com");
                    logger.Info("license expired or invalid.");
                    Application.Exit();
                }
                this.Text = "Attendance Process ver " + Application.ProductVersion;
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message, ex.StackTrace);
            }
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                FillEmployees();
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void grdEmployee_DataBindingComplete(object sender, GridViewBindingCompleteEventArgs e)
        {
            try
            {
                if (grdEmployee.Rows.Count > 0)
                {
                    grdEmployee.CurrentRow = grdEmployee.Rows[0];
                    grdEmployee.TableElement.ScrollToRow(grdEmployee.Rows[0]);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void grdProcess_DataBindingComplete(object sender, GridViewBindingCompleteEventArgs e)
        {
            try
            {
                if (grdProcess.Rows.Count > 0)
                {
                    grdProcess.CurrentRow = grdProcess.Rows[0];
                    grdProcess.TableElement.ScrollToRow(grdProcess.Rows[0]);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                SaveAttendance();
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void btnPost_Click(object sender, EventArgs e)
        {
            try
            {
                PostAttendance();
                grdEmployee.Visible = true;
                grdProcess.Visible = false;
                btnProcess.Text = "Process";
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void btnVoid_Click(object sender, EventArgs e)
        {
            try
            {
                VoidAttendance();
                grdEmployee.Visible = true;
                grdProcess.Visible = false;
                btnProcess.Text = "Process";
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void cmbPayroll_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    string SelectedPayroll = cmbPayroll.SelectedItem.ToString();
                    cmbPeriod.Items.Clear();
                    if (!string.IsNullOrEmpty(SelectedPayroll))
                    {
                        var oPeriods = (from a in odb.CfgPeriodDates
                                        where a.FlgLocked == false
                                        && a.CfgPayrollDefination.PayrollName == SelectedPayroll
                                        orderby a.StartDate
                                        select a).ToList();
                        foreach (var period in oPeriods)
                        {
                            cmbPeriod.Items.Add(period.PeriodName);
                        }
                    }
                    cmbPeriod.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                ImportTempData();
                lblStart.Text = "Data successfully imported for selected employees";
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void grdProcess_CellEndEdit(object sender, GridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex != -1)
                {
                    LineProcess(e.RowIndex);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void cmbPeriod_SelectedIndexChanged(object sender, Telerik.WinControls.UI.Data.PositionChangedEventArgs e)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    string SelectedPayroll = cmbPayroll.SelectedItem is null ? "" : cmbPayroll.SelectedItem.ToString();
                    string SelectedPeriod = cmbPeriod.SelectedItem is null ? "" : cmbPeriod.SelectedItem.ToString();

                    if (!string.IsNullOrEmpty(SelectedPayroll) && !string.IsNullOrEmpty(SelectedPeriod))
                    {
                        var oPeriods = (from a in odb.CfgPeriodDates
                                        where a.FlgLocked == false
                                        && a.CfgPayrollDefination.PayrollName == SelectedPayroll
                                        && a.PeriodName == SelectedPeriod
                                        select a).FirstOrDefault();
                        if (oPeriods is null)
                        {
                            return;
                        }
                        dtFrom.Value = oPeriods.StartDate.Value;
                        dtTo.Value = oPeriods.EndDate.Value;
                    }

                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void grdEmployee_CellClick(object sender, GridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex == -1)
                {
                    if (e.ColumnIndex == 0)
                    {
                        GridToggle();
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
            }
        }
        private void btnReport1_Click(object sender, EventArgs e)
        {
            try
            {
                //using (var odb = new dbHRMS(ConnectionString))
                //{
                //    var FirstEmployee = (from a in dtEmployees.AsEnumerable()
                //                         where a.Field<bool>("Select") == true
                //                         select a).FirstOrDefault();
                //    if (FirstEmployee is null) { RadMessageBox.Show("Select atleast one employee."); }
                //    string payrollvalue = cmbPayroll.SelectedItem.ToString();
                //    string periodvalue = cmbPeriod.SelectedItem.ToString();
                //    var oEmp = (from a in odb.MstEmployee
                //                where a.EmpID == FirstEmployee.Field<string>("EmpCode")
                //                select a).FirstOrDefault();
                //    var oPeriod = (from a in odb.CfgPeriodDates
                //                   where a.PeriodName == periodvalue
                //                   && a.CfgPayrollDefination.PayrollName == payrollvalue
                //                   select a).FirstOrDefault();

                //    frmReportViewer oDialog = new frmReportViewer();
                //    oDialog.ReportCode = 1;
                //    oDialog.EmpCode = oEmp.EmpID;
                //    oDialog.PeriodCode = oPeriod.ID.ToString();
                //    oDialog.ShowDialog();
                //}
                CallAttendanceReport();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        private void btnReport2_Click(object sender, EventArgs e)
        {
            try
            {
                //using (var odb = new dbHRMS(ConnectionString))
                //{
                //    var FirstEmployee = (from a in dtEmployees.AsEnumerable()
                //                         where a.Field<bool>("Select") == true
                //                         select a).FirstOrDefault();
                //    if (FirstEmployee is null) { RadMessageBox.Show("Select atleast one employee."); }
                //    string payrollvalue = cmbPayroll.SelectedItem.ToString();
                //    string periodvalue = cmbPeriod.SelectedItem.ToString();
                //    var oEmp = (from a in odb.MstEmployee
                //                where a.EmpID == FirstEmployee.Field<string>("EmpCode")
                //                select a).FirstOrDefault();
                //    var oPeriod = (from a in odb.CfgPeriodDates
                //                   where a.PeriodName == periodvalue
                //                   && a.CfgPayrollDefination.PayrollName == payrollvalue
                //                   select a).FirstOrDefault();

                //    frmReportViewer oDialog = new frmReportViewer();
                //    oDialog.ReportCode = 2;
                //    oDialog.EmpCode = oEmp.EmpID;
                //    oDialog.FromDate = oPeriod.StartDate.Value.ToString("yyyy-MM-dd");
                //    oDialog.ToDate = oPeriod.EndDate.Value.ToString("yyyy-MM-dd");
                //    oDialog.ShowDialog();
                //}
                CallTempAttendanceReport();
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        private void btnSendSlip_Click(object sender, EventArgs e)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    var EmployeeList = (from a in dtEmployees.AsEnumerable()
                                        where a.Field<bool>("Select") == true
                                        select a).ToList();
                    if (EmployeeList is null) { RadMessageBox.Show("Select atleast one employee."); }
                    string payrollvalue = cmbPayroll.SelectedItem.ToString();
                    string periodvalue = cmbPeriod.SelectedItem.ToString();
                    lblStart.Text = "Sending Emails Salary Slip to employees please wait.";
                    foreach (var SingleEmployee in EmployeeList)
                    {
                        var oEmp = (from a in odb.MstEmployee
                                    where a.EmpID == SingleEmployee.Field<string>("EmpCode")
                                    select a).FirstOrDefault();
                        var oPeriod = (from a in odb.CfgPeriodDates
                                       where a.PeriodName == periodvalue
                                       && a.CfgPayrollDefination.PayrollName == payrollvalue
                                       select a).FirstOrDefault();
                        var SlipId = (from a in odb.TrnsSalaryProcessRegister
                                      where a.EmpID == oEmp.ID
                                      && a.PayrollPeriodID == oPeriod.ID
                                      select a).FirstOrDefault();
                        if (string.IsNullOrWhiteSpace(oEmp.OfficeEmail))
                        {
                            logger.Log(LogLevel.Warn, $"email not define emp code {oEmp.EmpID}");
                            continue;
                        }
                        string EmpName = string.Concat(oEmp.FirstName, " ", oEmp.MiddleName, " ", oEmp.LastName);
                        SendSalarySlip(SlipId.Id, oPeriod.PeriodName, oPeriod.StartDate.GetValueOrDefault(), EmpName, oEmp.OfficeEmail);
                    }

                    lblStart.Text = "Process completed.";
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        private void btnSendAtt_Click(object sender, EventArgs e)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    var EmployeeList = (from a in dtEmployees.AsEnumerable()
                                        where a.Field<bool>("Select") == true
                                        select a).ToList();
                    if (EmployeeList is null) { RadMessageBox.Show("Select atleast one employee."); }
                    string payrollvalue = cmbPayroll.SelectedItem.ToString();
                    string periodvalue = cmbPeriod.SelectedItem.ToString();
                    lblStart.Text = "Sending Emails Attendance Slip to employees please wait.";
                    foreach (var SingleEmployee in EmployeeList)
                    {
                        var oEmp = (from a in odb.MstEmployee
                                    where a.EmpID == SingleEmployee.Field<string>("EmpCode")
                                    select a).FirstOrDefault();
                        var oPeriod = (from a in odb.CfgPeriodDates
                                       where a.PeriodName == periodvalue
                                       && a.CfgPayrollDefination.PayrollName == payrollvalue
                                       select a).FirstOrDefault();
                        var SlipId = (from a in odb.TrnsSalaryProcessRegister
                                      where a.EmpID == oEmp.ID
                                      && a.PayrollPeriodID == oPeriod.ID
                                      select a).FirstOrDefault();
                        if (string.IsNullOrWhiteSpace(oEmp.OfficeEmail))
                        {
                            logger.Log(LogLevel.Warn, $"email not define emp code {oEmp.EmpID}");
                            continue;
                        }
                        string EmpName = string.Concat(oEmp.FirstName, " ", oEmp.MiddleName, " ", oEmp.LastName);
                        //SendSalarySlip(SlipId.Id, oPeriod.PeriodName, oPeriod.StartDate.GetValueOrDefault(), EmpName, oEmp.OfficeEmail);
                        SendAttendanceSlip(oPeriod.ID, oPeriod.PeriodName, oEmp.EmpID, EmpName, oEmp.OfficeEmail, oPeriod.StartDate.GetValueOrDefault());
                    }

                    lblStart.Text = "Process completed.";
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }

        #endregion

    }
}
