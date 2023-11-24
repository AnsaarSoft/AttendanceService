using DIHRMS;
using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Telerik.WinControls.Svg;
using Telerik.WinControls.UI;

namespace AttendanceService
{
    public partial class frmMain : RadForm
    {

        #region Variables
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private string ConnectionString = "";
        private DataTable dtEmployees;
        private DataTable dtProcessed;

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
                dtProcessed.Columns.Add("Date");
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
                dtProcessed.Columns.Add("OTHour");
                dtProcessed.Columns.Add("OTType");
                dtProcessed.Columns.Add("LPCount");

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
                using (var odb = new dbHRMS(ConnectionString))
                {
                    string DepartmentValue, DesignationValue, LocationValue, BranchValue, PayrollValue;
                    DepartmentValue = cmbDepartment.SelectedItem.ToString() == "All" ? string.Empty : cmbDepartment.SelectedItem.ToString();
                    DesignationValue = cmbDesignation.SelectedItem.ToString() == "All" ? string.Empty : cmbDesignation.SelectedItem.ToString();
                    LocationValue = cmbLocation.SelectedItem.ToString() == "All" ? string.Empty : cmbLocation.SelectedItem.ToString();
                    BranchValue = cmbBranch.SelectedItem.ToString() == "All" ? string.Empty : cmbBranch.SelectedItem.ToString();
                    PayrollValue = cmbPayroll.SelectedItem.ToString();
                    IEnumerable<MstEmployee> oCollection;
                    if (string.IsNullOrEmpty(DepartmentValue) && string.IsNullOrEmpty(DesignationValue) && string.IsNullOrEmpty(LocationValue) && string.IsNullOrEmpty(BranchValue))
                    {
                        oCollection = (from a in odb.MstEmployee
                                       where a.FlgActive == true
                                       && a.CfgPayrollDefination.PayrollName == PayrollValue
                                       orderby a.ID
                                       select a).ToList();
                    }
                    else if (string.IsNullOrEmpty(DepartmentValue) && string.IsNullOrEmpty(DesignationValue) && string.IsNullOrEmpty(LocationValue) && !string.IsNullOrEmpty(BranchValue))
                    {
                        oCollection = (from a in odb.MstEmployee
                                       where a.FlgActive == true
                                       && a.MstBranches.Name == BranchValue
                                       && a.CfgPayrollDefination.PayrollName == PayrollValue
                                       orderby a.ID
                                       select a).ToList();
                    }
                    else if (string.IsNullOrEmpty(DepartmentValue) && string.IsNullOrEmpty(DesignationValue) && !string.IsNullOrEmpty(LocationValue) && !string.IsNullOrEmpty(BranchValue))
                    {
                        oCollection = (from a in odb.MstEmployee
                                       where a.FlgActive == true
                                       && a.MstLocation.Name == LocationValue
                                       && a.MstBranches.Name == BranchValue
                                       && a.CfgPayrollDefination.PayrollName == PayrollValue
                                       orderby a.ID
                                       select a).ToList();
                    }
                    else if (string.IsNullOrEmpty(DepartmentValue) && !string.IsNullOrEmpty(DesignationValue) && !string.IsNullOrEmpty(LocationValue) && !string.IsNullOrEmpty(BranchValue))
                    {
                        oCollection = (from a in odb.MstEmployee
                                       where a.FlgActive == true
                                       && a.MstDesignation.Name == DesignationValue
                                       && a.MstLocation.Name == LocationValue
                                       && a.MstBranches.Name == BranchValue
                                       && a.CfgPayrollDefination.PayrollName == PayrollValue
                                       orderby a.ID
                                       select a).ToList();
                    }
                    else if (!string.IsNullOrEmpty(DepartmentValue) && !string.IsNullOrEmpty(DesignationValue) && !string.IsNullOrEmpty(LocationValue) && !string.IsNullOrEmpty(BranchValue))
                    {
                        oCollection = (from a in odb.MstEmployee
                                       where a.FlgActive == true
                                       && a.MstDepartment.DeptName == DepartmentValue
                                       && a.MstDesignation.Name == DesignationValue
                                       && a.MstLocation.Name == LocationValue
                                       && a.MstBranches.Name == BranchValue
                                       && a.CfgPayrollDefination.PayrollName == PayrollValue
                                       orderby a.ID
                                       select a).ToList();
                    }
                    else
                    {
                        oCollection = (from a in odb.MstEmployee
                                       where a.FlgActive == true
                                       && a.CfgPayrollDefination.PayrollName == PayrollValue
                                       orderby a.ID
                                       select a).ToList();
                    }
                    if (oCollection.Count() > 0)
                    {
                        foreach (var emp in oCollection)
                        {
                            dtEmployees.Rows.Add(false, emp.EmpID, $"{emp.FirstName} {emp.MiddleName} {emp.LastName}", emp.DepartmentName, emp.DesignationName, emp.LocationName, emp.BranchName, emp.PayrollName);
                        }
                    }

                }
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
                            ProcessEmployeeMonth(oEmp, oPeriod);
                        }
                    }



                }
            }
            catch (Exception ex)
            {
                logger.Error(ex);
            }
        }
        void ProcessEmployeeMonth(MstEmployee oEmp, CfgPeriodDates oPeriod)
        {
            try
            {
                using (var odb = new dbHRMS(ConnectionString))
                {
                    DateTime loopStart = oPeriod.StartDate.GetValueOrDefault();
                    DateTime loopEnd = oPeriod.EndDate.GetValueOrDefault();
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

                            int OTId = 0;
                            string OTHours = string.Empty;
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

                            if (!flgConfirmedOut)
                            {
                                if (flgExpectedOut)
                                {
                                    TimeIn = (from a in odb.TrnsTempAttendance
                                              where (a.In_Out == "1" || a.In_Out == "01" || a.In_Out == "In")
                                              && a.PunchedDate == oAttendance.Date
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
                                               orderby a.ID ascending
                                               select a.PunchedTime).LastOrDefault();
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
                                               orderby a.ID ascending
                                               select a.PunchedTime).LastOrDefault();
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
                                    LateIn = TimeConvert(iTimeIn - iShiftIn);
                                    flgLateIn = true;
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
                            if (iTimeIn == 0 && iTimeOut == 0)
                            {
                                var LeaveCheck = (from a in odb.TrnsLeavesRequest
                                                  where a.LeaveFrom <= i && a.LeaveTo >= i
                                                  select a).Count();
                                if (LeaveCheck > 0)
                                {
                                    flgNewLeave = false;
                                    LeaveHour = ShiftDuration;
                                    LeaveType = "Absent";
                                    LeaveTypeID = 10;
                                }
                            }

                            //if workhour and shift hour is ok
                            ////but not on shift time.
                            if (iShiftDuration <= iWorkHour)
                            {
                                if (flgLateIn)
                                {
                                    flgNewLeave = false;
                                    LeaveHour = LateIn;
                                    LeaveType = "Absent";
                                    LeaveTypeID = 10;
                                }
                                if (flgEarlyOut)
                                {
                                    flgNewLeave = false;
                                    LeaveHour = EarlyOut;
                                    LeaveType = "Absent";
                                    LeaveTypeID = 10;
                                }
                            }


                            //if workhour is not ok 
                            if (iShiftDuration > iWorkHour)
                            {
                                if (flgLateIn)
                                {
                                    flgNewLeave = false;
                                    LeaveHour = LateIn;
                                    LeaveType = "Absent";
                                    LeaveTypeID = 10;
                                }
                                if (flgEarlyOut)
                                {
                                    flgNewLeave = false;
                                    LeaveHour = EarlyOut;
                                    LeaveType = "Absent";
                                    LeaveTypeID = 10;
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
                                    flgOT = true;
                                    OTHours = LateOut;
                                    OTId = oAttendance.MstShifts.MstOverTime == null ? 0 : oAttendance.MstShifts.MstOverTime.ID;
                                }
                            }

                            #endregion

                            #region Set Datatable
                            dtProcessed.Columns.Add("ID");
                            dtProcessed.Columns.Add("Date");
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
                            dtProcessed.Columns.Add("OTHour");
                            dtProcessed.Columns.Add("OTType");
                            dtProcessed.Columns.Add("LPCount");

                            DataRow dr = dtProcessed.NewRow();
                            dr["ID"] = oAttendance.Id;
                            dr["Date"] = oAttendance.Date;


                            #endregion

                            #endregion
                        }
                        //Attendace saved didn't re-calculate
                        else
                        {
                            logger.Info($"saved data was loaded.");
                            #region saved
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
                return $"{string.Format("{00}", hours)}:{string.Format("{00}", minutes)}";
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message);
                return "";
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

        #endregion

        #region Events

        public frmMain()
        {
            InitializeComponent();
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                //Task.Run(async () =>
                //{
                //    await StartProcess();
                //});
                StartProcess();
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
                CreateGrid();
                FillCombo();
            }
            catch (Exception ex)
            {
                logger.Error(ex, ex.Message, ex.StackTrace);
            }
        }
        private void cmbPayroll_SelectedIndexChanged(object sender, EventArgs e)
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

        #endregion

    }
}
