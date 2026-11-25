<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_Actual_In_Report
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CboGroupBy = New System.Windows.Forms.ComboBox
        Me.LlbColSetting = New System.Windows.Forms.LinkLabel
        Me.ChkNoClockInOut = New System.Windows.Forms.CheckBox
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.CboIfMultiple = New System.Windows.Forms.ComboBox
        Me.LblGroupOrMultipleClock = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.DtpTimeTo = New System.Windows.Forms.DateTimePicker
        Me.DtpTimeFrom = New System.Windows.Forms.DateTimePicker
        Me.BtnClose = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.ClbLocation = New System.Windows.Forms.CheckedListBox
        Me.BtnShow = New System.Windows.Forms.Button
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.LblDate = New System.Windows.Forms.Label
        Me.DtpDateTo = New System.Windows.Forms.DateTimePicker
        Me.LblDateTo = New System.Windows.Forms.Label
        Me.CrvActulaIn = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmpCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmpName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShiftDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchBreakIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSchOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActBreakOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActBreakIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColActOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.CboGroupBy)
        Me.GroupBox1.Controls.Add(Me.LlbColSetting)
        Me.GroupBox1.Controls.Add(Me.ChkNoClockInOut)
        Me.GroupBox1.Controls.Add(Me.CboFormat)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.CboIfMultiple)
        Me.GroupBox1.Controls.Add(Me.LblGroupOrMultipleClock)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.DtpTimeTo)
        Me.GroupBox1.Controls.Add(Me.DtpTimeFrom)
        Me.GroupBox1.Controls.Add(Me.BtnClose)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ClbLocation)
        Me.GroupBox1.Controls.Add(Me.BtnShow)
        Me.GroupBox1.Controls.Add(Me.DtpDate)
        Me.GroupBox1.Controls.Add(Me.LblDate)
        Me.GroupBox1.Controls.Add(Me.DtpDateTo)
        Me.GroupBox1.Controls.Add(Me.LblDateTo)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(798, 122)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'CboGroupBy
        '
        Me.CboGroupBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGroupBy.FormattingEnabled = True
        Me.CboGroupBy.Items.AddRange(New Object() {"Employee", "Shift"})
        Me.CboGroupBy.Location = New System.Drawing.Point(121, 93)
        Me.CboGroupBy.Name = "CboGroupBy"
        Me.CboGroupBy.Size = New System.Drawing.Size(147, 21)
        Me.CboGroupBy.TabIndex = 6
        '
        'LlbColSetting
        '
        Me.LlbColSetting.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbColSetting.AutoSize = True
        Me.LlbColSetting.Location = New System.Drawing.Point(342, 95)
        Me.LlbColSetting.Name = "LlbColSetting"
        Me.LlbColSetting.Size = New System.Drawing.Size(133, 13)
        Me.LlbColSetting.TabIndex = 9
        Me.LlbColSetting.TabStop = True
        Me.LlbColSetting.Text = "Column Settings (Format 2)"
        Me.LlbColSetting.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'ChkNoClockInOut
        '
        Me.ChkNoClockInOut.AutoSize = True
        Me.ChkNoClockInOut.Location = New System.Drawing.Point(344, 70)
        Me.ChkNoClockInOut.Name = "ChkNoClockInOut"
        Me.ChkNoClockInOut.Size = New System.Drawing.Size(188, 17)
        Me.ChkNoClockInOut.TabIndex = 8
        Me.ChkNoClockInOut.Text = "Only Show No Clock In/Clock Out"
        Me.ChkNoClockInOut.UseVisualStyleBackColor = True
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"Period", "Day"})
        Me.CboFormat.Location = New System.Drawing.Point(121, 14)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(147, 21)
        Me.CboFormat.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 19)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Format"
        '
        'CboIfMultiple
        '
        Me.CboIfMultiple.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboIfMultiple.FormattingEnabled = True
        Me.CboIfMultiple.Items.AddRange(New Object() {"First", "Last"})
        Me.CboIfMultiple.Location = New System.Drawing.Point(121, 93)
        Me.CboIfMultiple.Name = "CboIfMultiple"
        Me.CboIfMultiple.Size = New System.Drawing.Size(147, 21)
        Me.CboIfMultiple.TabIndex = 5
        '
        'LblGroupOrMultipleClock
        '
        Me.LblGroupOrMultipleClock.AutoSize = True
        Me.LblGroupOrMultipleClock.Location = New System.Drawing.Point(6, 96)
        Me.LblGroupOrMultipleClock.Name = "LblGroupOrMultipleClock"
        Me.LblGroupOrMultipleClock.Size = New System.Drawing.Size(111, 13)
        Me.LblGroupOrMultipleClock.TabIndex = 10
        Me.LblGroupOrMultipleClock.Text = "If multiple clock in/out"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(190, 69)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(10, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "-"
        '
        'DtpTimeTo
        '
        Me.DtpTimeTo.CustomFormat = "HH.mm"
        Me.DtpTimeTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTimeTo.Location = New System.Drawing.Point(210, 67)
        Me.DtpTimeTo.Name = "DtpTimeTo"
        Me.DtpTimeTo.ShowUpDown = True
        Me.DtpTimeTo.Size = New System.Drawing.Size(58, 20)
        Me.DtpTimeTo.TabIndex = 4
        '
        'DtpTimeFrom
        '
        Me.DtpTimeFrom.CustomFormat = "HH.mm"
        Me.DtpTimeFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTimeFrom.Location = New System.Drawing.Point(121, 67)
        Me.DtpTimeFrom.Name = "DtpTimeFrom"
        Me.DtpTimeFrom.ShowUpDown = True
        Me.DtpTimeFrom.Size = New System.Drawing.Size(58, 20)
        Me.DtpTimeFrom.TabIndex = 3
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(605, 38)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 11
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(282, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Location"
        '
        'ClbLocation
        '
        Me.ClbLocation.FormattingEnabled = True
        Me.ClbLocation.Location = New System.Drawing.Point(344, 14)
        Me.ClbLocation.Name = "ClbLocation"
        Me.ClbLocation.Size = New System.Drawing.Size(252, 49)
        Me.ClbLocation.TabIndex = 7
        '
        'BtnShow
        '
        Me.BtnShow.Location = New System.Drawing.Point(605, 14)
        Me.BtnShow.Name = "BtnShow"
        Me.BtnShow.Size = New System.Drawing.Size(75, 23)
        Me.BtnShow.TabIndex = 10
        Me.BtnShow.Text = "Show"
        Me.BtnShow.UseVisualStyleBackColor = True
        '
        'DtpDate
        '
        Me.DtpDate.CustomFormat = "dd MMMM yyyy"
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(121, 41)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(147, 20)
        Me.DtpDate.TabIndex = 1
        '
        'LblDate
        '
        Me.LblDate.AutoSize = True
        Me.LblDate.Location = New System.Drawing.Point(6, 45)
        Me.LblDate.Name = "LblDate"
        Me.LblDate.Size = New System.Drawing.Size(30, 13)
        Me.LblDate.TabIndex = 0
        Me.LblDate.Text = "Date"
        '
        'DtpDateTo
        '
        Me.DtpDateTo.CustomFormat = "dd MMMM yyyy"
        Me.DtpDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDateTo.Location = New System.Drawing.Point(121, 67)
        Me.DtpDateTo.Name = "DtpDateTo"
        Me.DtpDateTo.Size = New System.Drawing.Size(147, 20)
        Me.DtpDateTo.TabIndex = 2
        '
        'LblDateTo
        '
        Me.LblDateTo.AutoSize = True
        Me.LblDateTo.Location = New System.Drawing.Point(6, 71)
        Me.LblDateTo.Name = "LblDateTo"
        Me.LblDateTo.Size = New System.Drawing.Size(30, 13)
        Me.LblDateTo.TabIndex = 13
        Me.LblDateTo.Text = "Date"
        '
        'CrvActulaIn
        '
        Me.CrvActulaIn.ActiveViewIndex = -1
        Me.CrvActulaIn.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrvActulaIn.DisplayGroupTree = False
        Me.CrvActulaIn.DisplayStatusBar = False
        Me.CrvActulaIn.DisplayToolbar = False
        Me.CrvActulaIn.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrvActulaIn.EnableDrillDown = False
        Me.CrvActulaIn.Location = New System.Drawing.Point(3, 3)
        Me.CrvActulaIn.Name = "CrvActulaIn"
        Me.CrvActulaIn.SelectionFormula = ""
        Me.CrvActulaIn.ShowCloseButton = False
        Me.CrvActulaIn.ShowGroupTreeButton = False
        Me.CrvActulaIn.ShowRefreshButton = False
        Me.CrvActulaIn.Size = New System.Drawing.Size(784, 398)
        Me.CrvActulaIn.TabIndex = 0
        Me.CrvActulaIn.ViewTimeSelectionFormula = ""
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(7, 131)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(798, 430)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.CrvActulaIn)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(790, 404)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Format 1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvItem)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(790, 404)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Format 2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDate, Me.ColDeptCode, Me.ColDeptDesc, Me.ColEmpCode, Me.ColEmpName, Me.ColShift, Me.ColShiftDesc, Me.ColStatus, Me.ColSchIn, Me.ColSchBreakOut, Me.ColSchBreakIn, Me.ColSchOut, Me.ColActIn, Me.ColActBreakOut, Me.ColActBreakIn, Me.ColActOut})
        Me.DgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvItem.Location = New System.Drawing.Point(3, 3)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.ReadOnly = True
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.Size = New System.Drawing.Size(784, 398)
        Me.DgvItem.TabIndex = 0
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate.Width = 36
        '
        'ColDeptCode
        '
        Me.ColDeptCode.HeaderText = "Department Code"
        Me.ColDeptCode.Name = "ColDeptCode"
        Me.ColDeptCode.ReadOnly = True
        Me.ColDeptCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDeptCode.Width = 96
        '
        'ColDeptDesc
        '
        Me.ColDeptDesc.HeaderText = "Department Description"
        Me.ColDeptDesc.Name = "ColDeptDesc"
        Me.ColDeptDesc.ReadOnly = True
        Me.ColDeptDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDeptDesc.Width = 124
        '
        'ColEmpCode
        '
        Me.ColEmpCode.HeaderText = "Employee Code"
        Me.ColEmpCode.Name = "ColEmpCode"
        Me.ColEmpCode.ReadOnly = True
        Me.ColEmpCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmpCode.Width = 87
        '
        'ColEmpName
        '
        Me.ColEmpName.HeaderText = "Employee Name"
        Me.ColEmpName.Name = "ColEmpName"
        Me.ColEmpName.ReadOnly = True
        Me.ColEmpName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmpName.Width = 90
        '
        'ColShift
        '
        Me.ColShift.HeaderText = "Shift Code"
        Me.ColShift.Name = "ColShift"
        Me.ColShift.ReadOnly = True
        Me.ColShift.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColShift.Width = 62
        '
        'ColShiftDesc
        '
        Me.ColShiftDesc.HeaderText = "Shift Description"
        Me.ColShiftDesc.Name = "ColShiftDesc"
        Me.ColShiftDesc.ReadOnly = True
        Me.ColShiftDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColShiftDesc.Width = 90
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ReadOnly = True
        Me.ColStatus.Width = 62
        '
        'ColSchIn
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchIn.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColSchIn.HeaderText = "Schedule In"
        Me.ColSchIn.Name = "ColSchIn"
        Me.ColSchIn.ReadOnly = True
        Me.ColSchIn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSchIn.Width = 70
        '
        'ColSchBreakOut
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchBreakOut.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColSchBreakOut.HeaderText = "Schedule Break Out"
        Me.ColSchBreakOut.Name = "ColSchBreakOut"
        Me.ColSchBreakOut.ReadOnly = True
        Me.ColSchBreakOut.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSchBreakOut.Width = 109
        '
        'ColSchBreakIn
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchBreakIn.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColSchBreakIn.HeaderText = "Schedule Break In"
        Me.ColSchBreakIn.Name = "ColSchBreakIn"
        Me.ColSchBreakIn.ReadOnly = True
        Me.ColSchBreakIn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSchBreakIn.Width = 101
        '
        'ColSchOut
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSchOut.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColSchOut.HeaderText = "Schedule Out"
        Me.ColSchOut.Name = "ColSchOut"
        Me.ColSchOut.ReadOnly = True
        Me.ColSchOut.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSchOut.Width = 78
        '
        'ColActIn
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActIn.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColActIn.HeaderText = "Actual In"
        Me.ColActIn.Name = "ColActIn"
        Me.ColActIn.ReadOnly = True
        Me.ColActIn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColActIn.Width = 55
        '
        'ColActBreakOut
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActBreakOut.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColActBreakOut.HeaderText = "Actual Break Out"
        Me.ColActBreakOut.Name = "ColActBreakOut"
        Me.ColActBreakOut.ReadOnly = True
        Me.ColActBreakOut.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColActBreakOut.Width = 94
        '
        'ColActBreakIn
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActBreakIn.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColActBreakIn.HeaderText = "Actual Break In"
        Me.ColActBreakIn.Name = "ColActBreakIn"
        Me.ColActBreakIn.ReadOnly = True
        Me.ColActBreakIn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColActBreakIn.Width = 86
        '
        'ColActOut
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColActOut.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColActOut.HeaderText = "Actual Out"
        Me.ColActOut.Name = "ColActOut"
        Me.ColActOut.ReadOnly = True
        Me.ColActOut.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColActOut.Width = 63
        '
        'FrmTrans_Attendance_Actual_In_Report
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(812, 567)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimumSize = New System.Drawing.Size(828, 605)
        Me.Name = "FrmTrans_Attendance_Actual_In_Report"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Actual Clock In/Out Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblDate As System.Windows.Forms.Label
    Friend WithEvents BtnShow As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ClbLocation As System.Windows.Forms.CheckedListBox
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents CrvActulaIn As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents DtpTimeTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpTimeFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents CrpTrans_Attendance_Actual_In_Report1 As HR.CrpTrans_Attendance_Actual_In_Report
    Friend WithEvents CboIfMultiple As System.Windows.Forms.ComboBox
    Friend WithEvents LblGroupOrMultipleClock As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DtpDateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblDateTo As System.Windows.Forms.Label
    Friend WithEvents LlbColSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkNoClockInOut As System.Windows.Forms.CheckBox
    Friend WithEvents CboGroupBy As System.Windows.Forms.ComboBox
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmpCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmpName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShiftDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchBreakIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSchOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActBreakOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActBreakIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColActOut As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
