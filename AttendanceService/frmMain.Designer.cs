
namespace AttendanceService
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            Telerik.WinControls.UI.GridViewCheckBoxColumn gridViewCheckBoxColumn1 = new Telerik.WinControls.UI.GridViewCheckBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn1 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn2 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn3 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn4 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn5 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn6 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.UI.GridViewTextBoxColumn gridViewTextBoxColumn7 = new Telerik.WinControls.UI.GridViewTextBoxColumn();
            Telerik.WinControls.Data.SortDescriptor sortDescriptor1 = new Telerik.WinControls.Data.SortDescriptor();
            Telerik.WinControls.UI.TableViewDefinition tableViewDefinition1 = new Telerik.WinControls.UI.TableViewDefinition();
            this.btnStart = new System.Windows.Forms.Button();
            this.lblStart = new System.Windows.Forms.Label();
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.ni = new System.Windows.Forms.NotifyIcon(this.components);
            this.lblLastRun = new System.Windows.Forms.Label();
            this.cmbDepartment = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbDesignation = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbLocation = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbBranch = new System.Windows.Forms.ComboBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbPeriod = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbPayroll = new System.Windows.Forms.ComboBox();
            this.grdEmployee = new Telerik.WinControls.UI.RadGridView();
            ((System.ComponentModel.ISupportInitialize)(this.grdEmployee)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdEmployee.MasterTemplate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(228, 90);
            this.btnStart.Margin = new System.Windows.Forms.Padding(4);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(100, 28);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // lblStart
            // 
            this.lblStart.AutoSize = true;
            this.lblStart.Location = new System.Drawing.Point(359, 85);
            this.lblStart.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStart.Name = "lblStart";
            this.lblStart.Size = new System.Drawing.Size(18, 19);
            this.lblStart.TabIndex = 1;
            this.lblStart.Text = "...";
            // 
            // timer
            // 
            this.timer.Tick += new System.EventHandler(this.timer_Tick);
            // 
            // ni
            // 
            this.ni.Icon = ((System.Drawing.Icon)(resources.GetObject("ni.Icon")));
            this.ni.Text = "notifyIcon1";
            this.ni.Visible = true;
            // 
            // lblLastRun
            // 
            this.lblLastRun.AutoSize = true;
            this.lblLastRun.Location = new System.Drawing.Point(359, 116);
            this.lblLastRun.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblLastRun.Name = "lblLastRun";
            this.lblLastRun.Size = new System.Drawing.Size(18, 19);
            this.lblLastRun.TabIndex = 2;
            this.lblLastRun.Text = "...";
            // 
            // cmbDepartment
            // 
            this.cmbDepartment.FormattingEnabled = true;
            this.cmbDepartment.Location = new System.Drawing.Point(120, 20);
            this.cmbDepartment.Name = "cmbDepartment";
            this.cmbDepartment.Size = new System.Drawing.Size(200, 24);
            this.cmbDepartment.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 19);
            this.label1.TabIndex = 4;
            this.label1.Text = "Department";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(33, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 19);
            this.label2.TabIndex = 6;
            this.label2.Text = "Designation";
            // 
            // cmbDesignation
            // 
            this.cmbDesignation.FormattingEnabled = true;
            this.cmbDesignation.Location = new System.Drawing.Point(120, 46);
            this.cmbDesignation.Name = "cmbDesignation";
            this.cmbDesignation.Size = new System.Drawing.Size(200, 24);
            this.cmbDesignation.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(365, 28);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 19);
            this.label3.TabIndex = 8;
            this.label3.Text = "Location";
            // 
            // cmbLocation
            // 
            this.cmbLocation.FormattingEnabled = true;
            this.cmbLocation.Location = new System.Drawing.Point(452, 20);
            this.cmbLocation.Name = "cmbLocation";
            this.cmbLocation.Size = new System.Drawing.Size(200, 24);
            this.cmbLocation.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(365, 54);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 19);
            this.label4.TabIndex = 10;
            this.label4.Text = "Branch";
            // 
            // cmbBranch
            // 
            this.cmbBranch.FormattingEnabled = true;
            this.cmbBranch.Location = new System.Drawing.Point(452, 46);
            this.cmbBranch.Name = "cmbBranch";
            this.cmbBranch.Size = new System.Drawing.Size(200, 24);
            this.cmbBranch.TabIndex = 9;
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(120, 90);
            this.btnSearch.Margin = new System.Windows.Forms.Padding(4);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(100, 28);
            this.btnSearch.TabIndex = 11;
            this.btnSearch.Text = "Search";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(743, 54);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(47, 19);
            this.label6.TabIndex = 17;
            this.label6.Text = "Period";
            // 
            // cmbPeriod
            // 
            this.cmbPeriod.FormattingEnabled = true;
            this.cmbPeriod.Location = new System.Drawing.Point(830, 46);
            this.cmbPeriod.Name = "cmbPeriod";
            this.cmbPeriod.Size = new System.Drawing.Size(200, 24);
            this.cmbPeriod.TabIndex = 16;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(743, 28);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(50, 19);
            this.label7.TabIndex = 15;
            this.label7.Text = "Payroll";
            // 
            // cmbPayroll
            // 
            this.cmbPayroll.FormattingEnabled = true;
            this.cmbPayroll.Location = new System.Drawing.Point(830, 20);
            this.cmbPayroll.Name = "cmbPayroll";
            this.cmbPayroll.Size = new System.Drawing.Size(200, 24);
            this.cmbPayroll.TabIndex = 14;
            this.cmbPayroll.SelectedIndexChanged += new System.EventHandler(this.cmbPayroll_SelectedIndexChanged);
            // 
            // grdEmployee
            // 
            this.grdEmployee.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(219)))), ((int)(((byte)(255)))));
            this.grdEmployee.Cursor = System.Windows.Forms.Cursors.Default;
            this.grdEmployee.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.grdEmployee.ForeColor = System.Drawing.Color.Black;
            this.grdEmployee.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.grdEmployee.Location = new System.Drawing.Point(16, 156);
            this.grdEmployee.Margin = new System.Windows.Forms.Padding(15);
            // 
            // 
            // 
            this.grdEmployee.MasterTemplate.AllowAddNewRow = false;
            gridViewCheckBoxColumn1.AllowGroup = false;
            gridViewCheckBoxColumn1.EnableExpressionEditor = false;
            gridViewCheckBoxColumn1.FieldName = "Select";
            gridViewCheckBoxColumn1.HeaderText = "Select";
            gridViewCheckBoxColumn1.MinWidth = 95;
            gridViewCheckBoxColumn1.Name = "Select";
            gridViewCheckBoxColumn1.Width = 194;
            gridViewTextBoxColumn1.EnableExpressionEditor = false;
            gridViewTextBoxColumn1.FieldName = "EmpCode";
            gridViewTextBoxColumn1.HeaderText = "EmpCode";
            gridViewTextBoxColumn1.MinWidth = 24;
            gridViewTextBoxColumn1.Name = "EmpCode";
            gridViewTextBoxColumn1.ReadOnly = true;
            gridViewTextBoxColumn1.Width = 236;
            gridViewTextBoxColumn2.EnableExpressionEditor = false;
            gridViewTextBoxColumn2.FieldName = "EmpName";
            gridViewTextBoxColumn2.HeaderText = "EmpName";
            gridViewTextBoxColumn2.MinWidth = 24;
            gridViewTextBoxColumn2.Name = "EmpName";
            gridViewTextBoxColumn2.ReadOnly = true;
            gridViewTextBoxColumn2.Width = 568;
            gridViewTextBoxColumn3.EnableExpressionEditor = false;
            gridViewTextBoxColumn3.FieldName = "Department";
            gridViewTextBoxColumn3.HeaderText = "Department";
            gridViewTextBoxColumn3.MinWidth = 24;
            gridViewTextBoxColumn3.Name = "Department";
            gridViewTextBoxColumn3.ReadOnly = true;
            gridViewTextBoxColumn3.SortOrder = Telerik.WinControls.UI.RadSortOrder.Ascending;
            gridViewTextBoxColumn3.Width = 295;
            gridViewTextBoxColumn4.EnableExpressionEditor = false;
            gridViewTextBoxColumn4.FieldName = "Designation";
            gridViewTextBoxColumn4.HeaderText = "Designation";
            gridViewTextBoxColumn4.MinWidth = 19;
            gridViewTextBoxColumn4.Name = "Designation";
            gridViewTextBoxColumn4.ReadOnly = true;
            gridViewTextBoxColumn4.Width = 281;
            gridViewTextBoxColumn5.EnableExpressionEditor = false;
            gridViewTextBoxColumn5.FieldName = "Location";
            gridViewTextBoxColumn5.HeaderText = "Location";
            gridViewTextBoxColumn5.MinWidth = 19;
            gridViewTextBoxColumn5.Name = "Location";
            gridViewTextBoxColumn5.ReadOnly = true;
            gridViewTextBoxColumn5.Width = 252;
            gridViewTextBoxColumn6.EnableExpressionEditor = false;
            gridViewTextBoxColumn6.FieldName = "Branch";
            gridViewTextBoxColumn6.HeaderText = "Branch";
            gridViewTextBoxColumn6.MinWidth = 19;
            gridViewTextBoxColumn6.Name = "Branch";
            gridViewTextBoxColumn6.ReadOnly = true;
            gridViewTextBoxColumn6.Width = 278;
            gridViewTextBoxColumn7.EnableExpressionEditor = false;
            gridViewTextBoxColumn7.FieldName = "Payroll";
            gridViewTextBoxColumn7.HeaderText = "Payroll";
            gridViewTextBoxColumn7.MinWidth = 19;
            gridViewTextBoxColumn7.Name = "Payroll";
            gridViewTextBoxColumn7.ReadOnly = true;
            gridViewTextBoxColumn7.Width = 235;
            this.grdEmployee.MasterTemplate.Columns.AddRange(new Telerik.WinControls.UI.GridViewDataColumn[] {
            gridViewCheckBoxColumn1,
            gridViewTextBoxColumn1,
            gridViewTextBoxColumn2,
            gridViewTextBoxColumn3,
            gridViewTextBoxColumn4,
            gridViewTextBoxColumn5,
            gridViewTextBoxColumn6,
            gridViewTextBoxColumn7});
            this.grdEmployee.MasterTemplate.EnableAlternatingRowColor = true;
            this.grdEmployee.MasterTemplate.EnableGrouping = false;
            sortDescriptor1.PropertyName = "Department";
            this.grdEmployee.MasterTemplate.SortDescriptors.AddRange(new Telerik.WinControls.Data.SortDescriptor[] {
            sortDescriptor1});
            this.grdEmployee.MasterTemplate.ViewDefinition = tableViewDefinition1;
            this.grdEmployee.Name = "grdEmployee";
            this.grdEmployee.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.grdEmployee.Size = new System.Drawing.Size(1228, 534);
            this.grdEmployee.TabIndex = 18;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1263, 709);
            this.Controls.Add(this.grdEmployee);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.cmbPeriod);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cmbPayroll);
            this.Controls.Add(this.btnSearch);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbBranch);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbLocation);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbDesignation);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbDepartment);
            this.Controls.Add(this.lblLastRun);
            this.Controls.Add(this.lblStart);
            this.Controls.Add(this.btnStart);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Attendance Service";
            this.Load += new System.EventHandler(this.frmMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdEmployee.MasterTemplate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grdEmployee)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Label lblStart;
        private System.Windows.Forms.Timer timer;
        private System.Windows.Forms.NotifyIcon ni;
        private System.Windows.Forms.Label lblLastRun;
        private System.Windows.Forms.ComboBox cmbDepartment;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbDesignation;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbLocation;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbBranch;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbPeriod;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbPayroll;
        private Telerik.WinControls.UI.RadGridView grdEmployee;
    }
}

