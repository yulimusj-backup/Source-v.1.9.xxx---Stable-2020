<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCuti_Approval
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCuti = New System.Windows.Forms.TextBox
        Me.LblDuration = New System.Windows.Forms.LinkLabel
        Me.TxtPosition = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.DtpDate = New System.Windows.Forms.DateTimePicker
        Me.LblBalance = New System.Windows.Forms.LinkLabel
        Me.LblDaysBalance = New System.Windows.Forms.Label
        Me.TxtBalance = New System.Windows.Forms.TextBox
        Me.LblDaysDuration = New System.Windows.Forms.Label
        Me.ChkHDEnd = New System.Windows.Forms.CheckBox
        Me.ChkHDStart = New System.Windows.Forms.CheckBox
        Me.TxtDuration = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.DtpEnd = New System.Windows.Forms.DateTimePicker
        Me.dtpStart = New System.Windows.Forms.DateTimePicker
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtEmployeeName = New System.Windows.Forms.TextBox
        Me.BtnFindEmployee = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtEmployeeCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryEmployee = New System.Windows.Forms.TextBox
        Me.TxtCode = New System.Windows.Forms.TextBox
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboLeave = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DtpApprovedBy = New System.Windows.Forms.DateTimePicker
        Me.TxtApprovedBy = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtStatus = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtEnumStatus = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtReason = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.BtnReject = New System.Windows.Forms.Button
        Me.BtnApprove = New System.Windows.Forms.Button
        Me.RboWaitingForApproval = New System.Windows.Forms.RadioButton
        Me.RboApproved = New System.Windows.Forms.RadioButton
        Me.RboRejected = New System.Windows.Forms.RadioButton
        Me.BtnWait = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLeaveAplication = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEnumStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColStatusNotes = New System.Windows.Forms.DataGridViewButtonColumn
        Me.StatusNotesDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColModifiy = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDocumentNo, Me.ColDate, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColPrimaryLeaveAplication, Me.ColPrimaryEmployee, Me.ColEnumStatus, Me.ColStatus, Me.ColStatusNotes, Me.StatusNotesDataGridViewTextBoxColumn, Me.ColModifiy, Me.ColPrimary})
        Me.dgvItem.Location = New System.Drawing.Point(12, 38)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(684, 185)
        Me.dgvItem.TabIndex = 6
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 229)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TextBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryCuti)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblDuration)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPosition)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblBalance)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblDaysBalance)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtBalance)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblDaysDuration)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkHDEnd)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkHDStart)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtDuration)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpEnd)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpStart)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtEmployeeName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFindEmployee)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtEmployeeCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryEmployee)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboLeave)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Size = New System.Drawing.Size(684, 257)
        Me.SplitContainer1.SplitterDistance = 336
        Me.SplitContainer1.TabIndex = 7
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox1.Location = New System.Drawing.Point(93, 131)
        Me.TextBox1.MaxLength = 50
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(239, 20)
        Me.TextBox1.TabIndex = 54
        '
        'TxtPrimaryCuti
        '
        Me.TxtPrimaryCuti.Location = New System.Drawing.Point(93, 131)
        Me.TxtPrimaryCuti.Name = "TxtPrimaryCuti"
        Me.TxtPrimaryCuti.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryCuti.TabIndex = 53
        Me.TxtPrimaryCuti.TabStop = False
        Me.TxtPrimaryCuti.Text = "-1"
        '
        'LblDuration
        '
        Me.LblDuration.AutoSize = True
        Me.LblDuration.Location = New System.Drawing.Point(4, 217)
        Me.LblDuration.Name = "LblDuration"
        Me.LblDuration.Size = New System.Drawing.Size(47, 13)
        Me.LblDuration.TabIndex = 10
        Me.LblDuration.TabStop = True
        Me.LblDuration.Text = "Duration"
        '
        'TxtPosition
        '
        Me.TxtPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPosition.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPosition.Location = New System.Drawing.Point(93, 105)
        Me.TxtPosition.MaxLength = 50
        Me.TxtPosition.Name = "TxtPosition"
        Me.TxtPosition.ReadOnly = True
        Me.TxtPosition.Size = New System.Drawing.Size(239, 20)
        Me.TxtPosition.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 108)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 52
        Me.Label4.Text = "Position"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 35)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(30, 13)
        Me.Label11.TabIndex = 50
        Me.Label11.Text = "Date"
        '
        'DtpDate
        '
        Me.DtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpDate.CustomFormat = "dd MMM yyyy"
        Me.DtpDate.Enabled = False
        Me.DtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpDate.Location = New System.Drawing.Point(93, 29)
        Me.DtpDate.Name = "DtpDate"
        Me.DtpDate.Size = New System.Drawing.Size(239, 20)
        Me.DtpDate.TabIndex = 1
        Me.DtpDate.TabStop = False
        '
        'LblBalance
        '
        Me.LblBalance.AutoSize = True
        Me.LblBalance.Location = New System.Drawing.Point(3, 239)
        Me.LblBalance.Name = "LblBalance"
        Me.LblBalance.Size = New System.Drawing.Size(46, 13)
        Me.LblBalance.TabIndex = 12
        Me.LblBalance.TabStop = True
        Me.LblBalance.Text = "Balance"
        '
        'LblDaysBalance
        '
        Me.LblDaysBalance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblDaysBalance.AutoSize = True
        Me.LblDaysBalance.Location = New System.Drawing.Point(267, 239)
        Me.LblDaysBalance.Name = "LblDaysBalance"
        Me.LblDaysBalance.Size = New System.Drawing.Size(31, 13)
        Me.LblDaysBalance.TabIndex = 48
        Me.LblDaysBalance.Text = "Days"
        '
        'TxtBalance
        '
        Me.TxtBalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtBalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtBalance.Location = New System.Drawing.Point(93, 236)
        Me.TxtBalance.MaxLength = 50
        Me.TxtBalance.Name = "TxtBalance"
        Me.TxtBalance.ReadOnly = True
        Me.TxtBalance.Size = New System.Drawing.Size(167, 20)
        Me.TxtBalance.TabIndex = 13
        Me.TxtBalance.TabStop = False
        Me.TxtBalance.Text = "0"
        Me.TxtBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblDaysDuration
        '
        Me.LblDaysDuration.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblDaysDuration.AutoSize = True
        Me.LblDaysDuration.Location = New System.Drawing.Point(267, 213)
        Me.LblDaysDuration.Name = "LblDaysDuration"
        Me.LblDaysDuration.Size = New System.Drawing.Size(31, 13)
        Me.LblDaysDuration.TabIndex = 10
        Me.LblDaysDuration.Text = "Days"
        '
        'ChkHDEnd
        '
        Me.ChkHDEnd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkHDEnd.AutoSize = True
        Me.ChkHDEnd.Enabled = False
        Me.ChkHDEnd.Location = New System.Drawing.Point(266, 186)
        Me.ChkHDEnd.Name = "ChkHDEnd"
        Me.ChkHDEnd.Size = New System.Drawing.Size(67, 17)
        Me.ChkHDEnd.TabIndex = 9
        Me.ChkHDEnd.TabStop = False
        Me.ChkHDEnd.Text = "Half Day"
        Me.ChkHDEnd.UseVisualStyleBackColor = True
        '
        'ChkHDStart
        '
        Me.ChkHDStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkHDStart.AutoSize = True
        Me.ChkHDStart.Enabled = False
        Me.ChkHDStart.Location = New System.Drawing.Point(266, 160)
        Me.ChkHDStart.Name = "ChkHDStart"
        Me.ChkHDStart.Size = New System.Drawing.Size(67, 17)
        Me.ChkHDStart.TabIndex = 7
        Me.ChkHDStart.Text = "Half Day"
        Me.ChkHDStart.UseVisualStyleBackColor = True
        '
        'TxtDuration
        '
        Me.TxtDuration.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDuration.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDuration.Location = New System.Drawing.Point(93, 210)
        Me.TxtDuration.MaxLength = 50
        Me.TxtDuration.Name = "TxtDuration"
        Me.TxtDuration.ReadOnly = True
        Me.TxtDuration.Size = New System.Drawing.Size(167, 20)
        Me.TxtDuration.TabIndex = 11
        Me.TxtDuration.TabStop = False
        Me.TxtDuration.Text = "0"
        Me.TxtDuration.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 190)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(26, 13)
        Me.Label7.TabIndex = 44
        Me.Label7.Text = "End"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 164)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 13)
        Me.Label6.TabIndex = 43
        Me.Label6.Text = "Start"
        '
        'DtpEnd
        '
        Me.DtpEnd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpEnd.CustomFormat = "dd MMM yyyy"
        Me.DtpEnd.Enabled = False
        Me.DtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpEnd.Location = New System.Drawing.Point(93, 184)
        Me.DtpEnd.Name = "DtpEnd"
        Me.DtpEnd.Size = New System.Drawing.Size(167, 20)
        Me.DtpEnd.TabIndex = 8
        Me.DtpEnd.TabStop = False
        '
        'dtpStart
        '
        Me.dtpStart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpStart.CustomFormat = "dd MMM yyyy"
        Me.dtpStart.Enabled = False
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpStart.Location = New System.Drawing.Point(93, 158)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(167, 20)
        Me.dtpStart.TabIndex = 6
        Me.dtpStart.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 134)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 13)
        Me.Label5.TabIndex = 41
        Me.Label5.Text = "Leave Code"
        '
        'TxtEmployeeName
        '
        Me.TxtEmployeeName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtEmployeeName.Location = New System.Drawing.Point(93, 81)
        Me.TxtEmployeeName.MaxLength = 50
        Me.TxtEmployeeName.Name = "TxtEmployeeName"
        Me.TxtEmployeeName.ReadOnly = True
        Me.TxtEmployeeName.Size = New System.Drawing.Size(239, 20)
        Me.TxtEmployeeName.TabIndex = 3
        '
        'BtnFindEmployee
        '
        Me.BtnFindEmployee.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindEmployee.Enabled = False
        Me.BtnFindEmployee.Location = New System.Drawing.Point(291, 55)
        Me.BtnFindEmployee.Name = "BtnFindEmployee"
        Me.BtnFindEmployee.Size = New System.Drawing.Size(42, 20)
        Me.BtnFindEmployee.TabIndex = 2
        Me.BtnFindEmployee.Text = "Find"
        Me.BtnFindEmployee.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 13)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "Employee Name"
        '
        'TxtEmployeeCode
        '
        Me.TxtEmployeeCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmployeeCode.Location = New System.Drawing.Point(93, 55)
        Me.TxtEmployeeCode.MaxLength = 100
        Me.TxtEmployeeCode.Name = "TxtEmployeeCode"
        Me.TxtEmployeeCode.ReadOnly = True
        Me.TxtEmployeeCode.Size = New System.Drawing.Size(191, 20)
        Me.TxtEmployeeCode.TabIndex = 2
        Me.TxtEmployeeCode.TabStop = False
        '
        'TxtPrimaryEmployee
        '
        Me.TxtPrimaryEmployee.Location = New System.Drawing.Point(93, 55)
        Me.TxtPrimaryEmployee.Name = "TxtPrimaryEmployee"
        Me.TxtPrimaryEmployee.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryEmployee.TabIndex = 35
        Me.TxtPrimaryEmployee.TabStop = False
        Me.TxtPrimaryEmployee.Text = "-1"
        '
        'TxtCode
        '
        Me.TxtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCode.Location = New System.Drawing.Point(93, 3)
        Me.TxtCode.MaxLength = 50
        Me.TxtCode.Name = "TxtCode"
        Me.TxtCode.ReadOnly = True
        Me.TxtCode.Size = New System.Drawing.Size(239, 20)
        Me.TxtCode.TabIndex = 0
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(93, 3)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimary.TabIndex = 0
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 13)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "Employee Code"
        '
        'cboLeave
        '
        Me.cboLeave.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLeave.Enabled = False
        Me.cboLeave.FormattingEnabled = True
        Me.cboLeave.Location = New System.Drawing.Point(93, 131)
        Me.cboLeave.Name = "cboLeave"
        Me.cboLeave.Size = New System.Drawing.Size(240, 21)
        Me.cboLeave.TabIndex = 5
        Me.cboLeave.TabStop = False
        Me.cboLeave.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Document No"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.DtpApprovedBy)
        Me.GroupBox2.Controls.Add(Me.TxtApprovedBy)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.TxtStatus)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.TxtEnumStatus)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(338, 67)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Approval"
        '
        'DtpApprovedBy
        '
        Me.DtpApprovedBy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpApprovedBy.CustomFormat = "dd MMM yyyy"
        Me.DtpApprovedBy.Enabled = False
        Me.DtpApprovedBy.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpApprovedBy.Location = New System.Drawing.Point(226, 43)
        Me.DtpApprovedBy.Name = "DtpApprovedBy"
        Me.DtpApprovedBy.Size = New System.Drawing.Size(106, 20)
        Me.DtpApprovedBy.TabIndex = 2
        Me.DtpApprovedBy.TabStop = False
        '
        'TxtApprovedBy
        '
        Me.TxtApprovedBy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtApprovedBy.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtApprovedBy.Location = New System.Drawing.Point(96, 43)
        Me.TxtApprovedBy.MaxLength = 50
        Me.TxtApprovedBy.Name = "TxtApprovedBy"
        Me.TxtApprovedBy.ReadOnly = True
        Me.TxtApprovedBy.Size = New System.Drawing.Size(123, 20)
        Me.TxtApprovedBy.TabIndex = 1
        Me.TxtApprovedBy.TabStop = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 46)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(62, 13)
        Me.Label12.TabIndex = 56
        Me.Label12.Text = "Approve By"
        '
        'TxtStatus
        '
        Me.TxtStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtStatus.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStatus.Location = New System.Drawing.Point(96, 19)
        Me.TxtStatus.MaxLength = 50
        Me.TxtStatus.Name = "TxtStatus"
        Me.TxtStatus.ReadOnly = True
        Me.TxtStatus.Size = New System.Drawing.Size(236, 20)
        Me.TxtStatus.TabIndex = 0
        Me.TxtStatus.TabStop = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 22)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(37, 13)
        Me.Label13.TabIndex = 54
        Me.Label13.Text = "Status"
        '
        'TxtEnumStatus
        '
        Me.TxtEnumStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEnumStatus.Enabled = False
        Me.TxtEnumStatus.Location = New System.Drawing.Point(96, 19)
        Me.TxtEnumStatus.MaxLength = 100
        Me.TxtEnumStatus.Name = "TxtEnumStatus"
        Me.TxtEnumStatus.Size = New System.Drawing.Size(27, 20)
        Me.TxtEnumStatus.TabIndex = 55
        Me.TxtEnumStatus.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtReason)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 84)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(338, 167)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Reason"
        '
        'TxtReason
        '
        Me.TxtReason.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtReason.Enabled = False
        Me.TxtReason.Location = New System.Drawing.Point(3, 16)
        Me.TxtReason.MaxLength = 100
        Me.TxtReason.Multiline = True
        Me.TxtReason.Name = "TxtReason"
        Me.TxtReason.Size = New System.Drawing.Size(332, 148)
        Me.TxtReason.TabIndex = 0
        Me.TxtReason.TabStop = False
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(12, 13)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(43, 13)
        Me.Label14.TabIndex = 40
        Me.Label14.Text = "Show : "
        '
        'BtnReject
        '
        Me.BtnReject.Location = New System.Drawing.Point(426, 9)
        Me.BtnReject.Name = "BtnReject"
        Me.BtnReject.Size = New System.Drawing.Size(75, 23)
        Me.BtnReject.TabIndex = 4
        Me.BtnReject.Text = "Reject"
        Me.BtnReject.UseVisualStyleBackColor = True
        '
        'BtnApprove
        '
        Me.BtnApprove.Location = New System.Drawing.Point(345, 8)
        Me.BtnApprove.Name = "BtnApprove"
        Me.BtnApprove.Size = New System.Drawing.Size(75, 23)
        Me.BtnApprove.TabIndex = 3
        Me.BtnApprove.Text = "Approve"
        Me.BtnApprove.UseVisualStyleBackColor = True
        '
        'RboWaitingForApproval
        '
        Me.RboWaitingForApproval.AutoSize = True
        Me.RboWaitingForApproval.Checked = True
        Me.RboWaitingForApproval.Location = New System.Drawing.Point(61, 12)
        Me.RboWaitingForApproval.Name = "RboWaitingForApproval"
        Me.RboWaitingForApproval.Size = New System.Drawing.Size(124, 17)
        Me.RboWaitingForApproval.TabIndex = 0
        Me.RboWaitingForApproval.TabStop = True
        Me.RboWaitingForApproval.Text = "Waiting For Approval"
        Me.RboWaitingForApproval.UseVisualStyleBackColor = True
        '
        'RboApproved
        '
        Me.RboApproved.AutoSize = True
        Me.RboApproved.Location = New System.Drawing.Point(191, 11)
        Me.RboApproved.Name = "RboApproved"
        Me.RboApproved.Size = New System.Drawing.Size(71, 17)
        Me.RboApproved.TabIndex = 1
        Me.RboApproved.Text = "Approved"
        Me.RboApproved.UseVisualStyleBackColor = True
        '
        'RboRejected
        '
        Me.RboRejected.AutoSize = True
        Me.RboRejected.Location = New System.Drawing.Point(268, 11)
        Me.RboRejected.Name = "RboRejected"
        Me.RboRejected.Size = New System.Drawing.Size(68, 17)
        Me.RboRejected.TabIndex = 2
        Me.RboRejected.Text = "Rejected"
        Me.RboRejected.UseVisualStyleBackColor = True
        '
        'BtnWait
        '
        Me.BtnWait.Location = New System.Drawing.Point(507, 9)
        Me.BtnWait.Name = "BtnWait"
        Me.BtnWait.Size = New System.Drawing.Size(75, 23)
        Me.BtnWait.TabIndex = 5
        Me.BtnWait.Text = "Wait"
        Me.BtnWait.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(565, 85)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 41
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'ColDocumentNo
        '
        Me.ColDocumentNo.DataPropertyName = "no transaksi"
        Me.ColDocumentNo.HeaderText = "Document No"
        Me.ColDocumentNo.Name = "ColDocumentNo"
        Me.ColDocumentNo.Width = 98
        '
        'ColDate
        '
        Me.ColDate.DataPropertyName = "tanggal"
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.Width = 55
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.DataPropertyName = "Kode Karyawan"
        Me.ColEmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.Width = 106
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.DataPropertyName = "Full Name"
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.ReadOnly = True
        Me.ColEmployeeName.Width = 109
        '
        'ColPrimaryLeaveAplication
        '
        Me.ColPrimaryLeaveAplication.DataPropertyName = "primary cuti"
        Me.ColPrimaryLeaveAplication.HeaderText = "Primary Leave Aplication"
        Me.ColPrimaryLeaveAplication.Name = "ColPrimaryLeaveAplication"
        Me.ColPrimaryLeaveAplication.ReadOnly = True
        Me.ColPrimaryLeaveAplication.Visible = False
        Me.ColPrimaryLeaveAplication.Width = 148
        '
        'ColPrimaryEmployee
        '
        Me.ColPrimaryEmployee.DataPropertyName = "primary karyawan"
        Me.ColPrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColPrimaryEmployee.Name = "ColPrimaryEmployee"
        Me.ColPrimaryEmployee.ReadOnly = True
        Me.ColPrimaryEmployee.Visible = False
        Me.ColPrimaryEmployee.Width = 115
        '
        'ColEnumStatus
        '
        Me.ColEnumStatus.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ColEnumStatus.DataPropertyName = "status"
        Me.ColEnumStatus.HeaderText = "Status"
        Me.ColEnumStatus.MinimumWidth = 2
        Me.ColEnumStatus.Name = "ColEnumStatus"
        Me.ColEnumStatus.ReadOnly = True
        Me.ColEnumStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEnumStatus.Visible = False
        Me.ColEnumStatus.Width = 2
        '
        'ColStatus
        '
        Me.ColStatus.AutoComplete = False
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Items.AddRange(New Object() {"Waiting For Approval", "Approved", "Rejected"})
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ReadOnly = True
        Me.ColStatus.Width = 43
        '
        'ColStatusNotes
        '
        Me.ColStatusNotes.HeaderText = "Status Notes"
        Me.ColStatusNotes.Name = "ColStatusNotes"
        Me.ColStatusNotes.ReadOnly = True
        Me.ColStatusNotes.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColStatusNotes.Width = 74
        '
        'StatusNotesDataGridViewTextBoxColumn
        '
        Me.StatusNotesDataGridViewTextBoxColumn.DataPropertyName = "status notes"
        Me.StatusNotesDataGridViewTextBoxColumn.HeaderText = "status notes"
        Me.StatusNotesDataGridViewTextBoxColumn.Name = "StatusNotesDataGridViewTextBoxColumn"
        Me.StatusNotesDataGridViewTextBoxColumn.Visible = False
        Me.StatusNotesDataGridViewTextBoxColumn.Width = 89
        '
        'ColModifiy
        '
        Me.ColModifiy.HeaderText = "Modify"
        Me.ColModifiy.Name = "ColModifiy"
        Me.ColModifiy.ReadOnly = True
        Me.ColModifiy.Width = 44
        '
        'ColPrimary
        '
        Me.ColPrimary.DataPropertyName = "primary"
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.ReadOnly = True
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 66
        '
        'FrmTransaksiCuti_Approval
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(716, 498)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.BtnWait)
        Me.Controls.Add(Me.RboRejected)
        Me.Controls.Add(Me.RboApproved)
        Me.Controls.Add(Me.RboWaitingForApproval)
        Me.Controls.Add(Me.BtnReject)
        Me.Controls.Add(Me.BtnApprove)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(724, 499)
        Me.Name = "FrmTransaksiCuti_Approval"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Rejected"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPosition As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblBalance As System.Windows.Forms.LinkLabel
    Friend WithEvents LblDaysBalance As System.Windows.Forms.Label
    Friend WithEvents TxtBalance As System.Windows.Forms.TextBox
    Friend WithEvents LblDaysDuration As System.Windows.Forms.Label
    Friend WithEvents ChkHDEnd As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHDStart As System.Windows.Forms.CheckBox
    Friend WithEvents TxtDuration As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtEmployeeName As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindEmployee As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtEmployeeCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryEmployee As System.Windows.Forms.TextBox
    Friend WithEvents TxtCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DtpApprovedBy As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtApprovedBy As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents BtnReject As System.Windows.Forms.Button
    Friend WithEvents BtnApprove As System.Windows.Forms.Button
    Friend WithEvents RboWaitingForApproval As System.Windows.Forms.RadioButton
    Friend WithEvents RboApproved As System.Windows.Forms.RadioButton
    Friend WithEvents RboRejected As System.Windows.Forms.RadioButton
    Friend WithEvents BtnWait As System.Windows.Forms.Button
    Friend WithEvents LblDuration As System.Windows.Forms.LinkLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cboLeave As System.Windows.Forms.ComboBox
    Friend WithEvents TxtReason As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCuti As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TxtEnumStatus As System.Windows.Forms.TextBox
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLeaveAplication As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEnumStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColStatusNotes As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents StatusNotesDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColModifiy As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
