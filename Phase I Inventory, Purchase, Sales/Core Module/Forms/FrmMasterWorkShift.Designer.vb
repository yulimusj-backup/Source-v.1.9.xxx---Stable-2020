<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkShift
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnDelete = New System.Windows.Forms.Button
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvDailyActivities = New System.Windows.Forms.DataGridView
        Me.ColDay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkdays = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColOvertime = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColDuration = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColSchIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakOut_WA = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakIn_WA = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.COlEnumDayOfWeek = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvHolidayList = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_Day = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_Duration = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColItemHoliday_SchIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakOut_Holiday = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakIn_Holiday = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_SchOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_Overtime = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColItemHoliday_Leave = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemHoliday_PrimaryLeave = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.chkSendNotifEmailOvertimeOrder = New System.Windows.Forms.CheckBox
        Me.chkUseOvertimeOrder = New System.Windows.Forms.CheckBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.chkCalculateOTTolerance = New System.Windows.Forms.CheckBox
        Me.chkCalculateEarlyOutTolerance = New System.Windows.Forms.CheckBox
        Me.chkCalculateLateTolerance = New System.Windows.Forms.CheckBox
        Me.chkCalculateEarlyOTTolerance = New System.Windows.Forms.CheckBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtMaxValueOT = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeOT = New System.Windows.Forms.ComboBox
        Me.txtRoundingOT = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterOT = New System.Windows.Forms.TextBox
        Me.txtMaxValueEarlyOut = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeEarlyOut = New System.Windows.Forms.ComboBox
        Me.txtRoundingEarlyOut = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterEarlyOut = New System.Windows.Forms.TextBox
        Me.txtMaxValueLate = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeLate = New System.Windows.Forms.ComboBox
        Me.txtRoundingLate = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterLate = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtMaxValueEarlyOT = New System.Windows.Forms.TextBox
        Me.cboRoundingTypeEarlyOT = New System.Windows.Forms.ComboBox
        Me.txtRoundingEarlyOT = New System.Windows.Forms.TextBox
        Me.txtCalculateAfterEarlyOT = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.CboIsCalculateOT = New System.Windows.Forms.ComboBox
        Me.CboIsCalculateEarlyOut = New System.Windows.Forms.ComboBox
        Me.CboIsCalculateLate = New System.Windows.Forms.ComboBox
        Me.CboIsCalculateEarlyOT = New System.Windows.Forms.ComboBox
        Me.DgvAttParams = New System.Windows.Forms.DataGridView
        Me.ColAtt_Schedule = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtt_GetValue = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAtt_LimitBefore = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtt_LimitAfter = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtt_FirstOrLast = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColAtt_Default = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColAtt_Outstation = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColAtt_DefaultCustom = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAtt_ScheduleType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TxtJobName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtJobCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPrimaryJob = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.gboDetail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvDailyActivities, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvHolidayList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvAttParams, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(446, 443)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.LinkLabel1)
        Me.gboDetail.Controls.Add(Me.TxtRemark)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TxtJobName)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtJobCode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryJob)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Location = New System.Drawing.Point(10, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(816, 437)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(6, 398)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(57, 13)
        Me.LinkLabel1.TabIndex = 35
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Copy Data"
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(60, 45)
        Me.TxtRemark.MaxLength = 150
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.Size = New System.Drawing.Size(750, 20)
        Me.TxtRemark.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Remarks"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(6, 71)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(804, 324)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvDailyActivities)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage1.Size = New System.Drawing.Size(796, 298)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Weekly Activities"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvDailyActivities
        '
        Me.dgvDailyActivities.AllowUserToAddRows = False
        Me.dgvDailyActivities.AllowUserToDeleteRows = False
        Me.dgvDailyActivities.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDailyActivities.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDay, Me.ColWorkdays, Me.ColOvertime, Me.ColDuration, Me.ColSchIn, Me.ColSchBreakOut_WA, Me.ColSchBreakIn_WA, Me.ColSchOut, Me.ColRemarks, Me.COlEnumDayOfWeek})
        Me.dgvDailyActivities.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDailyActivities.Location = New System.Drawing.Point(3, 3)
        Me.dgvDailyActivities.MultiSelect = False
        Me.dgvDailyActivities.Name = "dgvDailyActivities"
        Me.dgvDailyActivities.RowTemplate.Height = 24
        Me.dgvDailyActivities.Size = New System.Drawing.Size(790, 292)
        Me.dgvDailyActivities.TabIndex = 0
        '
        'ColDay
        '
        Me.ColDay.HeaderText = "Day"
        Me.ColDay.Name = "ColDay"
        Me.ColDay.Width = 51
        '
        'ColWorkdays
        '
        Me.ColWorkdays.HeaderText = "Workdays"
        Me.ColWorkdays.Items.AddRange(New Object() {"Workdays", "Holidays"})
        Me.ColWorkdays.Name = "ColWorkdays"
        Me.ColWorkdays.Width = 61
        '
        'ColOvertime
        '
        Me.ColOvertime.HeaderText = "Overtime"
        Me.ColOvertime.Items.AddRange(New Object() {"Weekday Overtime", "Holiday Overtime", "No Overtime"})
        Me.ColOvertime.Name = "ColOvertime"
        Me.ColOvertime.Width = 55
        '
        'ColDuration
        '
        Me.ColDuration.HeaderText = "Duration"
        Me.ColDuration.Items.AddRange(New Object() {"Full day", "Half day"})
        Me.ColDuration.Name = "ColDuration"
        Me.ColDuration.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDuration.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColDuration.Width = 72
        '
        'ColSchIn
        '
        Me.ColSchIn.HeaderText = "Schedule In"
        Me.ColSchIn.Name = "ColSchIn"
        Me.ColSchIn.Width = 89
        '
        'ColSchBreakOut_WA
        '
        Me.ColSchBreakOut_WA.HeaderText = "Schedule Break Out"
        Me.ColSchBreakOut_WA.Name = "ColSchBreakOut_WA"
        Me.ColSchBreakOut_WA.Width = 128
        '
        'ColSchBreakIn_WA
        '
        Me.ColSchBreakIn_WA.HeaderText = "Schedule Break In"
        Me.ColSchBreakIn_WA.Name = "ColSchBreakIn_WA"
        Me.ColSchBreakIn_WA.Width = 120
        '
        'ColSchOut
        '
        Me.ColSchOut.HeaderText = "Schedule Out"
        Me.ColSchOut.Name = "ColSchOut"
        Me.ColSchOut.Width = 97
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'COlEnumDayOfWeek
        '
        Me.COlEnumDayOfWeek.HeaderText = "Enum Day Of Week"
        Me.COlEnumDayOfWeek.Name = "COlEnumDayOfWeek"
        Me.COlEnumDayOfWeek.Visible = False
        Me.COlEnumDayOfWeek.Width = 127
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvHolidayList)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage2.Size = New System.Drawing.Size(786, 299)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Holiday List"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvHolidayList
        '
        Me.DgvHolidayList.AllowUserToAddRows = False
        Me.DgvHolidayList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvHolidayList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColItemHoliday_Date, Me.ColItemHoliday_Day, Me.ColItemHoliday_Description, Me.ColItemHoliday_Duration, Me.ColItemHoliday_SchIn, Me.ColSchBreakOut_Holiday, Me.ColSchBreakIn_Holiday, Me.ColItemHoliday_SchOut, Me.ColItemHoliday_Overtime, Me.ColItemHoliday_Leave, Me.ColItemHoliday_Remarks, Me.ColItemHoliday_PrimaryLeave})
        Me.DgvHolidayList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvHolidayList.Location = New System.Drawing.Point(3, 3)
        Me.DgvHolidayList.Name = "DgvHolidayList"
        Me.DgvHolidayList.RowTemplate.Height = 24
        Me.DgvHolidayList.Size = New System.Drawing.Size(780, 293)
        Me.DgvHolidayList.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColItemHoliday_Date
        '
        Me.ColItemHoliday_Date.HeaderText = "Date"
        Me.ColItemHoliday_Date.Name = "ColItemHoliday_Date"
        Me.ColItemHoliday_Date.Width = 55
        '
        'ColItemHoliday_Day
        '
        Me.ColItemHoliday_Day.HeaderText = "Day"
        Me.ColItemHoliday_Day.Name = "ColItemHoliday_Day"
        Me.ColItemHoliday_Day.Width = 51
        '
        'ColItemHoliday_Description
        '
        Me.ColItemHoliday_Description.HeaderText = "Description"
        Me.ColItemHoliday_Description.Name = "ColItemHoliday_Description"
        Me.ColItemHoliday_Description.Width = 85
        '
        'ColItemHoliday_Duration
        '
        Me.ColItemHoliday_Duration.HeaderText = "Duration"
        Me.ColItemHoliday_Duration.Items.AddRange(New Object() {"Full day", "Half day"})
        Me.ColItemHoliday_Duration.Name = "ColItemHoliday_Duration"
        Me.ColItemHoliday_Duration.Width = 53
        '
        'ColItemHoliday_SchIn
        '
        Me.ColItemHoliday_SchIn.HeaderText = "Schedule In"
        Me.ColItemHoliday_SchIn.Name = "ColItemHoliday_SchIn"
        Me.ColItemHoliday_SchIn.ReadOnly = True
        Me.ColItemHoliday_SchIn.Width = 89
        '
        'ColSchBreakOut_Holiday
        '
        Me.ColSchBreakOut_Holiday.HeaderText = "Schedule Break Out"
        Me.ColSchBreakOut_Holiday.Name = "ColSchBreakOut_Holiday"
        Me.ColSchBreakOut_Holiday.ReadOnly = True
        Me.ColSchBreakOut_Holiday.Width = 128
        '
        'ColSchBreakIn_Holiday
        '
        Me.ColSchBreakIn_Holiday.HeaderText = "Schedule Break In"
        Me.ColSchBreakIn_Holiday.Name = "ColSchBreakIn_Holiday"
        Me.ColSchBreakIn_Holiday.ReadOnly = True
        Me.ColSchBreakIn_Holiday.Width = 120
        '
        'ColItemHoliday_SchOut
        '
        Me.ColItemHoliday_SchOut.HeaderText = "Schedule Out"
        Me.ColItemHoliday_SchOut.Name = "ColItemHoliday_SchOut"
        Me.ColItemHoliday_SchOut.ReadOnly = True
        Me.ColItemHoliday_SchOut.Width = 97
        '
        'ColItemHoliday_Overtime
        '
        Me.ColItemHoliday_Overtime.HeaderText = "Overtime"
        Me.ColItemHoliday_Overtime.Items.AddRange(New Object() {"Weekday Overtime", "Holiday Overtime"})
        Me.ColItemHoliday_Overtime.Name = "ColItemHoliday_Overtime"
        Me.ColItemHoliday_Overtime.Width = 55
        '
        'ColItemHoliday_Leave
        '
        Me.ColItemHoliday_Leave.HeaderText = "Leave Code"
        Me.ColItemHoliday_Leave.Name = "ColItemHoliday_Leave"
        Me.ColItemHoliday_Leave.Visible = False
        Me.ColItemHoliday_Leave.Width = 109
        '
        'ColItemHoliday_Remarks
        '
        Me.ColItemHoliday_Remarks.HeaderText = "Remarks"
        Me.ColItemHoliday_Remarks.Name = "ColItemHoliday_Remarks"
        Me.ColItemHoliday_Remarks.Width = 74
        '
        'ColItemHoliday_PrimaryLeave
        '
        Me.ColItemHoliday_PrimaryLeave.HeaderText = "Primary Leave"
        Me.ColItemHoliday_PrimaryLeave.Name = "ColItemHoliday_PrimaryLeave"
        Me.ColItemHoliday_PrimaryLeave.Visible = False
        Me.ColItemHoliday_PrimaryLeave.Width = 124
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.chkSendNotifEmailOvertimeOrder)
        Me.TabPage3.Controls.Add(Me.chkUseOvertimeOrder)
        Me.TabPage3.Controls.Add(Me.LinkLabel2)
        Me.TabPage3.Controls.Add(Me.chkCalculateOTTolerance)
        Me.TabPage3.Controls.Add(Me.chkCalculateEarlyOutTolerance)
        Me.TabPage3.Controls.Add(Me.chkCalculateLateTolerance)
        Me.TabPage3.Controls.Add(Me.chkCalculateEarlyOTTolerance)
        Me.TabPage3.Controls.Add(Me.Label12)
        Me.TabPage3.Controls.Add(Me.txtMaxValueOT)
        Me.TabPage3.Controls.Add(Me.cboRoundingTypeOT)
        Me.TabPage3.Controls.Add(Me.txtRoundingOT)
        Me.TabPage3.Controls.Add(Me.txtCalculateAfterOT)
        Me.TabPage3.Controls.Add(Me.txtMaxValueEarlyOut)
        Me.TabPage3.Controls.Add(Me.cboRoundingTypeEarlyOut)
        Me.TabPage3.Controls.Add(Me.txtRoundingEarlyOut)
        Me.TabPage3.Controls.Add(Me.txtCalculateAfterEarlyOut)
        Me.TabPage3.Controls.Add(Me.txtMaxValueLate)
        Me.TabPage3.Controls.Add(Me.cboRoundingTypeLate)
        Me.TabPage3.Controls.Add(Me.txtRoundingLate)
        Me.TabPage3.Controls.Add(Me.txtCalculateAfterLate)
        Me.TabPage3.Controls.Add(Me.Label11)
        Me.TabPage3.Controls.Add(Me.Label10)
        Me.TabPage3.Controls.Add(Me.Label9)
        Me.TabPage3.Controls.Add(Me.txtMaxValueEarlyOT)
        Me.TabPage3.Controls.Add(Me.cboRoundingTypeEarlyOT)
        Me.TabPage3.Controls.Add(Me.txtRoundingEarlyOT)
        Me.TabPage3.Controls.Add(Me.txtCalculateAfterEarlyOT)
        Me.TabPage3.Controls.Add(Me.Label8)
        Me.TabPage3.Controls.Add(Me.Label6)
        Me.TabPage3.Controls.Add(Me.Label7)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.Label4)
        Me.TabPage3.Controls.Add(Me.CboIsCalculateOT)
        Me.TabPage3.Controls.Add(Me.CboIsCalculateEarlyOut)
        Me.TabPage3.Controls.Add(Me.CboIsCalculateLate)
        Me.TabPage3.Controls.Add(Me.CboIsCalculateEarlyOT)
        Me.TabPage3.Controls.Add(Me.DgvAttParams)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabPage3.Size = New System.Drawing.Size(786, 299)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Attendance Settings"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'chkSendNotifEmailOvertimeOrder
        '
        Me.chkSendNotifEmailOvertimeOrder.AutoSize = True
        Me.chkSendNotifEmailOvertimeOrder.Location = New System.Drawing.Point(238, 122)
        Me.chkSendNotifEmailOvertimeOrder.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkSendNotifEmailOvertimeOrder.Name = "chkSendNotifEmailOvertimeOrder"
        Me.chkSendNotifEmailOvertimeOrder.Size = New System.Drawing.Size(280, 17)
        Me.chkSendNotifEmailOvertimeOrder.TabIndex = 194
        Me.chkSendNotifEmailOvertimeOrder.Text = "Send notification email for overtime order authorization"
        Me.chkSendNotifEmailOvertimeOrder.UseVisualStyleBackColor = True
        '
        'chkUseOvertimeOrder
        '
        Me.chkUseOvertimeOrder.AutoSize = True
        Me.chkUseOvertimeOrder.Location = New System.Drawing.Point(70, 122)
        Me.chkUseOvertimeOrder.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkUseOvertimeOrder.Name = "chkUseOvertimeOrder"
        Me.chkUseOvertimeOrder.Size = New System.Drawing.Size(115, 17)
        Me.chkUseOvertimeOrder.TabIndex = 193
        Me.chkUseOvertimeOrder.Text = "Use overtime order"
        Me.chkUseOvertimeOrder.UseVisualStyleBackColor = True
        '
        'LinkLabel2
        '
        Me.LinkLabel2.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(662, 25)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(91, 13)
        Me.LinkLabel2.TabIndex = 192
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Advance Settings"
        Me.LinkLabel2.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'chkCalculateOTTolerance
        '
        Me.chkCalculateOTTolerance.AutoSize = True
        Me.chkCalculateOTTolerance.Checked = True
        Me.chkCalculateOTTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalculateOTTolerance.Location = New System.Drawing.Point(535, 99)
        Me.chkCalculateOTTolerance.Name = "chkCalculateOTTolerance"
        Me.chkCalculateOTTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkCalculateOTTolerance.TabIndex = 191
        Me.chkCalculateOTTolerance.Text = "Calculate Tolerance"
        Me.chkCalculateOTTolerance.UseVisualStyleBackColor = True
        '
        'chkCalculateEarlyOutTolerance
        '
        Me.chkCalculateEarlyOutTolerance.AutoSize = True
        Me.chkCalculateEarlyOutTolerance.Checked = True
        Me.chkCalculateEarlyOutTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalculateEarlyOutTolerance.Location = New System.Drawing.Point(535, 74)
        Me.chkCalculateEarlyOutTolerance.Name = "chkCalculateEarlyOutTolerance"
        Me.chkCalculateEarlyOutTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkCalculateEarlyOutTolerance.TabIndex = 190
        Me.chkCalculateEarlyOutTolerance.Text = "Calculate Tolerance"
        Me.chkCalculateEarlyOutTolerance.UseVisualStyleBackColor = True
        '
        'chkCalculateLateTolerance
        '
        Me.chkCalculateLateTolerance.AutoSize = True
        Me.chkCalculateLateTolerance.Checked = True
        Me.chkCalculateLateTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalculateLateTolerance.Location = New System.Drawing.Point(535, 49)
        Me.chkCalculateLateTolerance.Name = "chkCalculateLateTolerance"
        Me.chkCalculateLateTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkCalculateLateTolerance.TabIndex = 189
        Me.chkCalculateLateTolerance.Text = "Calculate Tolerance"
        Me.chkCalculateLateTolerance.UseVisualStyleBackColor = True
        '
        'chkCalculateEarlyOTTolerance
        '
        Me.chkCalculateEarlyOTTolerance.AutoSize = True
        Me.chkCalculateEarlyOTTolerance.Checked = True
        Me.chkCalculateEarlyOTTolerance.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalculateEarlyOTTolerance.Location = New System.Drawing.Point(535, 24)
        Me.chkCalculateEarlyOTTolerance.Name = "chkCalculateEarlyOTTolerance"
        Me.chkCalculateEarlyOTTolerance.Size = New System.Drawing.Size(121, 17)
        Me.chkCalculateEarlyOTTolerance.TabIndex = 188
        Me.chkCalculateEarlyOTTolerance.Text = "Calculate Tolerance"
        Me.chkCalculateEarlyOTTolerance.UseVisualStyleBackColor = True
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(78, 5)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(51, 13)
        Me.Label12.TabIndex = 187
        Me.Label12.Text = "Calculate"
        '
        'txtMaxValueOT
        '
        Me.txtMaxValueOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueOT.Location = New System.Drawing.Point(448, 97)
        Me.txtMaxValueOT.MaxLength = 150
        Me.txtMaxValueOT.Name = "txtMaxValueOT"
        Me.txtMaxValueOT.Size = New System.Drawing.Size(81, 20)
        Me.txtMaxValueOT.TabIndex = 180
        Me.txtMaxValueOT.Text = "00.00"
        Me.txtMaxValueOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeOT
        '
        Me.cboRoundingTypeOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeOT.FormattingEnabled = True
        Me.cboRoundingTypeOT.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeOT.Location = New System.Drawing.Point(238, 97)
        Me.cboRoundingTypeOT.Name = "cboRoundingTypeOT"
        Me.cboRoundingTypeOT.Size = New System.Drawing.Size(100, 21)
        Me.cboRoundingTypeOT.TabIndex = 178
        '
        'txtRoundingOT
        '
        Me.txtRoundingOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingOT.Location = New System.Drawing.Point(353, 98)
        Me.txtRoundingOT.MaxLength = 150
        Me.txtRoundingOT.Name = "txtRoundingOT"
        Me.txtRoundingOT.Size = New System.Drawing.Size(81, 20)
        Me.txtRoundingOT.TabIndex = 179
        Me.txtRoundingOT.Text = "00.00"
        Me.txtRoundingOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterOT
        '
        Me.txtCalculateAfterOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterOT.Location = New System.Drawing.Point(144, 97)
        Me.txtCalculateAfterOT.MaxLength = 150
        Me.txtCalculateAfterOT.Name = "txtCalculateAfterOT"
        Me.txtCalculateAfterOT.Size = New System.Drawing.Size(81, 20)
        Me.txtCalculateAfterOT.TabIndex = 177
        Me.txtCalculateAfterOT.Text = "00.00"
        Me.txtCalculateAfterOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxValueEarlyOut
        '
        Me.txtMaxValueEarlyOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueEarlyOut.Location = New System.Drawing.Point(448, 72)
        Me.txtMaxValueEarlyOut.MaxLength = 150
        Me.txtMaxValueEarlyOut.Name = "txtMaxValueEarlyOut"
        Me.txtMaxValueEarlyOut.Size = New System.Drawing.Size(81, 20)
        Me.txtMaxValueEarlyOut.TabIndex = 175
        Me.txtMaxValueEarlyOut.Text = "00.00"
        Me.txtMaxValueEarlyOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeEarlyOut
        '
        Me.cboRoundingTypeEarlyOut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeEarlyOut.FormattingEnabled = True
        Me.cboRoundingTypeEarlyOut.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeEarlyOut.Location = New System.Drawing.Point(238, 72)
        Me.cboRoundingTypeEarlyOut.Name = "cboRoundingTypeEarlyOut"
        Me.cboRoundingTypeEarlyOut.Size = New System.Drawing.Size(100, 21)
        Me.cboRoundingTypeEarlyOut.TabIndex = 173
        '
        'txtRoundingEarlyOut
        '
        Me.txtRoundingEarlyOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingEarlyOut.Location = New System.Drawing.Point(353, 73)
        Me.txtRoundingEarlyOut.MaxLength = 150
        Me.txtRoundingEarlyOut.Name = "txtRoundingEarlyOut"
        Me.txtRoundingEarlyOut.Size = New System.Drawing.Size(81, 20)
        Me.txtRoundingEarlyOut.TabIndex = 174
        Me.txtRoundingEarlyOut.Text = "00.00"
        Me.txtRoundingEarlyOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterEarlyOut
        '
        Me.txtCalculateAfterEarlyOut.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterEarlyOut.Location = New System.Drawing.Point(144, 72)
        Me.txtCalculateAfterEarlyOut.MaxLength = 150
        Me.txtCalculateAfterEarlyOut.Name = "txtCalculateAfterEarlyOut"
        Me.txtCalculateAfterEarlyOut.Size = New System.Drawing.Size(81, 20)
        Me.txtCalculateAfterEarlyOut.TabIndex = 172
        Me.txtCalculateAfterEarlyOut.Text = "00.00"
        Me.txtCalculateAfterEarlyOut.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtMaxValueLate
        '
        Me.txtMaxValueLate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueLate.Location = New System.Drawing.Point(448, 47)
        Me.txtMaxValueLate.MaxLength = 150
        Me.txtMaxValueLate.Name = "txtMaxValueLate"
        Me.txtMaxValueLate.Size = New System.Drawing.Size(81, 20)
        Me.txtMaxValueLate.TabIndex = 170
        Me.txtMaxValueLate.Text = "00.00"
        Me.txtMaxValueLate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeLate
        '
        Me.cboRoundingTypeLate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeLate.FormattingEnabled = True
        Me.cboRoundingTypeLate.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeLate.Location = New System.Drawing.Point(238, 47)
        Me.cboRoundingTypeLate.Name = "cboRoundingTypeLate"
        Me.cboRoundingTypeLate.Size = New System.Drawing.Size(100, 21)
        Me.cboRoundingTypeLate.TabIndex = 168
        '
        'txtRoundingLate
        '
        Me.txtRoundingLate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingLate.Location = New System.Drawing.Point(353, 48)
        Me.txtRoundingLate.MaxLength = 150
        Me.txtRoundingLate.Name = "txtRoundingLate"
        Me.txtRoundingLate.Size = New System.Drawing.Size(81, 20)
        Me.txtRoundingLate.TabIndex = 169
        Me.txtRoundingLate.Text = "00.00"
        Me.txtRoundingLate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterLate
        '
        Me.txtCalculateAfterLate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterLate.Location = New System.Drawing.Point(144, 47)
        Me.txtCalculateAfterLate.MaxLength = 150
        Me.txtCalculateAfterLate.Name = "txtCalculateAfterLate"
        Me.txtCalculateAfterLate.Size = New System.Drawing.Size(81, 20)
        Me.txtCalculateAfterLate.TabIndex = 167
        Me.txtCalculateAfterLate.Text = "00.00"
        Me.txtCalculateAfterLate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(460, 4)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 13)
        Me.Label11.TabIndex = 174
        Me.Label11.Text = "Max value"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(353, 4)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(82, 13)
        Me.Label10.TabIndex = 173
        Me.Label10.Text = "Rounding value"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(250, 4)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(76, 13)
        Me.Label9.TabIndex = 172
        Me.Label9.Text = "Rounding type"
        '
        'txtMaxValueEarlyOT
        '
        Me.txtMaxValueEarlyOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaxValueEarlyOT.Location = New System.Drawing.Point(448, 22)
        Me.txtMaxValueEarlyOT.MaxLength = 150
        Me.txtMaxValueEarlyOT.Name = "txtMaxValueEarlyOT"
        Me.txtMaxValueEarlyOT.Size = New System.Drawing.Size(81, 20)
        Me.txtMaxValueEarlyOT.TabIndex = 165
        Me.txtMaxValueEarlyOT.Text = "00.00"
        Me.txtMaxValueEarlyOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboRoundingTypeEarlyOT
        '
        Me.cboRoundingTypeEarlyOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRoundingTypeEarlyOT.FormattingEnabled = True
        Me.cboRoundingTypeEarlyOT.Items.AddRange(New Object() {"No Rounding", "Round Up", "Round down", "Middle"})
        Me.cboRoundingTypeEarlyOT.Location = New System.Drawing.Point(238, 22)
        Me.cboRoundingTypeEarlyOT.Name = "cboRoundingTypeEarlyOT"
        Me.cboRoundingTypeEarlyOT.Size = New System.Drawing.Size(100, 21)
        Me.cboRoundingTypeEarlyOT.TabIndex = 163
        '
        'txtRoundingEarlyOT
        '
        Me.txtRoundingEarlyOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRoundingEarlyOT.Location = New System.Drawing.Point(353, 23)
        Me.txtRoundingEarlyOT.MaxLength = 150
        Me.txtRoundingEarlyOT.Name = "txtRoundingEarlyOT"
        Me.txtRoundingEarlyOT.Size = New System.Drawing.Size(81, 20)
        Me.txtRoundingEarlyOT.TabIndex = 164
        Me.txtRoundingEarlyOT.Text = "00.00"
        Me.txtRoundingEarlyOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCalculateAfterEarlyOT
        '
        Me.txtCalculateAfterEarlyOT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCalculateAfterEarlyOT.Location = New System.Drawing.Point(144, 22)
        Me.txtCalculateAfterEarlyOT.MaxLength = 150
        Me.txtCalculateAfterEarlyOT.Name = "txtCalculateAfterEarlyOT"
        Me.txtCalculateAfterEarlyOT.Size = New System.Drawing.Size(81, 20)
        Me.txtCalculateAfterEarlyOT.TabIndex = 162
        Me.txtCalculateAfterEarlyOT.Text = "00.00"
        Me.txtCalculateAfterEarlyOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(149, 4)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(75, 13)
        Me.Label8.TabIndex = 56
        Me.Label8.Text = "Calculate after"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 100)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(22, 13)
        Me.Label6.TabIndex = 168
        Me.Label6.Text = "OT"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 74)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 13)
        Me.Label7.TabIndex = 167
        Me.Label7.Text = "Early out"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(28, 13)
        Me.Label5.TabIndex = 166
        Me.Label5.Text = "Late"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 26)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 13)
        Me.Label4.TabIndex = 165
        Me.Label4.Text = "Early OT"
        '
        'CboIsCalculateOT
        '
        Me.CboIsCalculateOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateOT.FormattingEnabled = True
        Me.CboIsCalculateOT.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateOT.Location = New System.Drawing.Point(70, 97)
        Me.CboIsCalculateOT.Name = "CboIsCalculateOT"
        Me.CboIsCalculateOT.Size = New System.Drawing.Size(62, 21)
        Me.CboIsCalculateOT.TabIndex = 176
        '
        'CboIsCalculateEarlyOut
        '
        Me.CboIsCalculateEarlyOut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateEarlyOut.FormattingEnabled = True
        Me.CboIsCalculateEarlyOut.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateEarlyOut.Location = New System.Drawing.Point(70, 72)
        Me.CboIsCalculateEarlyOut.Name = "CboIsCalculateEarlyOut"
        Me.CboIsCalculateEarlyOut.Size = New System.Drawing.Size(62, 21)
        Me.CboIsCalculateEarlyOut.TabIndex = 171
        '
        'CboIsCalculateLate
        '
        Me.CboIsCalculateLate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateLate.FormattingEnabled = True
        Me.CboIsCalculateLate.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateLate.Location = New System.Drawing.Point(70, 47)
        Me.CboIsCalculateLate.Name = "CboIsCalculateLate"
        Me.CboIsCalculateLate.Size = New System.Drawing.Size(62, 21)
        Me.CboIsCalculateLate.TabIndex = 166
        '
        'CboIsCalculateEarlyOT
        '
        Me.CboIsCalculateEarlyOT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIsCalculateEarlyOT.FormattingEnabled = True
        Me.CboIsCalculateEarlyOT.Items.AddRange(New Object() {"Yes", "No"})
        Me.CboIsCalculateEarlyOT.Location = New System.Drawing.Point(70, 22)
        Me.CboIsCalculateEarlyOT.Name = "CboIsCalculateEarlyOT"
        Me.CboIsCalculateEarlyOT.Size = New System.Drawing.Size(62, 21)
        Me.CboIsCalculateEarlyOT.TabIndex = 161
        '
        'DgvAttParams
        '
        Me.DgvAttParams.AllowUserToAddRows = False
        Me.DgvAttParams.AllowUserToDeleteRows = False
        Me.DgvAttParams.AllowUserToResizeRows = False
        Me.DgvAttParams.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvAttParams.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvAttParams.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvAttParams.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAtt_Schedule, Me.ColAtt_GetValue, Me.ColAtt_LimitBefore, Me.ColAtt_LimitAfter, Me.ColAtt_FirstOrLast, Me.ColAtt_Default, Me.ColAtt_Outstation, Me.ColAtt_DefaultCustom, Me.ColAtt_ScheduleType})
        Me.DgvAttParams.Location = New System.Drawing.Point(9, 145)
        Me.DgvAttParams.Name = "DgvAttParams"
        Me.DgvAttParams.RowHeadersVisible = False
        Me.DgvAttParams.RowTemplate.Height = 24
        Me.DgvAttParams.Size = New System.Drawing.Size(776, 155)
        Me.DgvAttParams.TabIndex = 1
        '
        'ColAtt_Schedule
        '
        Me.ColAtt_Schedule.HeaderText = "Schedule Name"
        Me.ColAtt_Schedule.Name = "ColAtt_Schedule"
        Me.ColAtt_Schedule.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAtt_Schedule.Width = 89
        '
        'ColAtt_GetValue
        '
        Me.ColAtt_GetValue.HeaderText = "Get Value"
        Me.ColAtt_GetValue.Name = "ColAtt_GetValue"
        Me.ColAtt_GetValue.Width = 60
        '
        'ColAtt_LimitBefore
        '
        Me.ColAtt_LimitBefore.HeaderText = "Limit Before"
        Me.ColAtt_LimitBefore.MaxInputLength = 3
        Me.ColAtt_LimitBefore.Name = "ColAtt_LimitBefore"
        Me.ColAtt_LimitBefore.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAtt_LimitBefore.Width = 68
        '
        'ColAtt_LimitAfter
        '
        Me.ColAtt_LimitAfter.HeaderText = "Limit After"
        Me.ColAtt_LimitAfter.MaxInputLength = 3
        Me.ColAtt_LimitAfter.Name = "ColAtt_LimitAfter"
        Me.ColAtt_LimitAfter.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAtt_LimitAfter.Width = 59
        '
        'ColAtt_FirstOrLast
        '
        Me.ColAtt_FirstOrLast.HeaderText = "If Multiple Clock In / Out"
        Me.ColAtt_FirstOrLast.Items.AddRange(New Object() {"First", "Last"})
        Me.ColAtt_FirstOrLast.Name = "ColAtt_FirstOrLast"
        Me.ColAtt_FirstOrLast.Width = 128
        '
        'ColAtt_Default
        '
        Me.ColAtt_Default.HeaderText = "If Clock In/Out not Found"
        Me.ColAtt_Default.Items.AddRange(New Object() {"Schedule", "Custom"})
        Me.ColAtt_Default.Name = "ColAtt_Default"
        Me.ColAtt_Default.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAtt_Default.Width = 134
        '
        'ColAtt_Outstation
        '
        Me.ColAtt_Outstation.HeaderText = "If Outstation"
        Me.ColAtt_Outstation.Items.AddRange(New Object() {"Schedule", "Custom"})
        Me.ColAtt_Outstation.Name = "ColAtt_Outstation"
        Me.ColAtt_Outstation.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAtt_Outstation.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColAtt_Outstation.Width = 89
        '
        'ColAtt_DefaultCustom
        '
        Me.ColAtt_DefaultCustom.HeaderText = "Default Custom"
        Me.ColAtt_DefaultCustom.MaxInputLength = 5
        Me.ColAtt_DefaultCustom.Name = "ColAtt_DefaultCustom"
        Me.ColAtt_DefaultCustom.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAtt_DefaultCustom.Width = 85
        '
        'ColAtt_ScheduleType
        '
        Me.ColAtt_ScheduleType.HeaderText = "Schedule Type"
        Me.ColAtt_ScheduleType.Name = "ColAtt_ScheduleType"
        Me.ColAtt_ScheduleType.ReadOnly = True
        Me.ColAtt_ScheduleType.Visible = False
        Me.ColAtt_ScheduleType.Width = 104
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(412, 404)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(324, 404)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 26)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtJobName
        '
        Me.TxtJobName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJobName.Location = New System.Drawing.Point(209, 19)
        Me.TxtJobName.MaxLength = 150
        Me.TxtJobName.Name = "TxtJobName"
        Me.TxtJobName.Size = New System.Drawing.Size(601, 20)
        Me.TxtJobName.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(143, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "Description"
        '
        'TxtJobCode
        '
        Me.TxtJobCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtJobCode.Location = New System.Drawing.Point(60, 19)
        Me.TxtJobCode.MaxLength = 150
        Me.TxtJobCode.Name = "TxtJobCode"
        Me.TxtJobCode.Size = New System.Drawing.Size(81, 20)
        Me.TxtJobCode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Code"
        '
        'TxtPrimaryJob
        '
        Me.TxtPrimaryJob.Location = New System.Drawing.Point(26, 19)
        Me.TxtPrimaryJob.Name = "TxtPrimaryJob"
        Me.TxtPrimaryJob.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryJob.TabIndex = 0
        Me.TxtPrimaryJob.TabStop = False
        Me.TxtPrimaryJob.Text = "-1"
        Me.TxtPrimaryJob.Visible = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(102, 19)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(39, 20)
        Me.TxtReadOnly.TabIndex = 55
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(188, 443)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(80, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 476)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(837, 22)
        Me.StatusStrip2.SizingGrip = False
        Me.StatusStrip2.TabIndex = 54
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(205, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(205, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(205, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(205, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(532, 443)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(360, 443)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(80, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(274, 443)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'FrmMasterWorkShift
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(837, 498)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnNew)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(853, 536)
        Me.Name = "FrmMasterWorkShift"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Workshift"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvDailyActivities, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvHolidayList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.DgvAttParams, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtJobName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtJobCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryJob As System.Windows.Forms.TextBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvDailyActivities As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DgvHolidayList As System.Windows.Forms.DataGridView
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ColDay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkdays As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColOvertime As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColDuration As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColSchIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakOut_WA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakIn_WA As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents COlEnumDayOfWeek As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvAttParams As System.Windows.Forms.DataGridView
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CboIsCalculateOT As System.Windows.Forms.ComboBox
    Friend WithEvents CboIsCalculateEarlyOut As System.Windows.Forms.ComboBox
    Friend WithEvents CboIsCalculateLate As System.Windows.Forms.ComboBox
    Friend WithEvents CboIsCalculateEarlyOT As System.Windows.Forms.ComboBox
    Friend WithEvents txtMaxValueEarlyOT As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeEarlyOT As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingEarlyOT As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterEarlyOT As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtMaxValueOT As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeOT As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingOT As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterOT As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxValueEarlyOut As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeEarlyOut As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingEarlyOut As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterEarlyOut As System.Windows.Forms.TextBox
    Friend WithEvents txtMaxValueLate As System.Windows.Forms.TextBox
    Friend WithEvents cboRoundingTypeLate As System.Windows.Forms.ComboBox
    Friend WithEvents txtRoundingLate As System.Windows.Forms.TextBox
    Friend WithEvents txtCalculateAfterLate As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents ColAtt_Schedule As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_GetValue As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAtt_LimitBefore As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_LimitAfter As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_FirstOrLast As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColAtt_Default As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColAtt_Outstation As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColAtt_DefaultCustom As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAtt_ScheduleType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkCalculateOTTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateEarlyOutTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateLateTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalculateEarlyOTTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents chkSendNotifEmailOvertimeOrder As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseOvertimeOrder As System.Windows.Forms.CheckBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_Day As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_Duration As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColItemHoliday_SchIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakOut_Holiday As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakIn_Holiday As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_SchOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_Overtime As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColItemHoliday_Leave As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemHoliday_PrimaryLeave As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
