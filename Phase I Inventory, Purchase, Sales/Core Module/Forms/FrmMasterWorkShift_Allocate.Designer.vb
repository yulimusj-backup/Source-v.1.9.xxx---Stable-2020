<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterWorkShift_Allocate
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnData = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpEnd = New System.Windows.Forms.DateTimePicker
        Me.DtpStart = New System.Windows.Forms.DateTimePicker
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.Lbl_Departemen_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Departemen_Checked = New System.Windows.Forms.LinkLabel
        Me.ChkDepartemen = New System.Windows.Forms.CheckBox
        Me.ClbDepartemen = New System.Windows.Forms.CheckedListBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Lbl_Position_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Position_Checked = New System.Windows.Forms.LinkLabel
        Me.ChkPosition = New System.Windows.Forms.CheckBox
        Me.ClbPositon = New System.Windows.Forms.CheckedListBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Lbl_SF_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_SF_Checked = New System.Windows.Forms.LinkLabel
        Me.ChkFormula = New System.Windows.Forms.CheckBox
        Me.ClbSalaryFormula = New System.Windows.Forms.CheckedListBox
        Me.gboEmployee = New System.Windows.Forms.GroupBox
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkEmployee = New System.Windows.Forms.CheckBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.BtnDelete = New System.Windows.Forms.Button
        Me.BtnAllocate = New System.Windows.Forms.Button
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.DgvWorkshift = New System.Windows.Forms.DataGridView
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkshiftCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWorkShiftName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSave = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColDelete = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPrimaryWorkShift = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEntity = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEntityTipe = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnClose = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker_CloseForm = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker6 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gboEmployee.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.DgvWorkshift, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtnData)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.DtpEnd)
        Me.GroupBox1.Controls.Add(Me.DtpStart)
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.gboEmployee)
        Me.GroupBox1.Location = New System.Drawing.Point(7, 25)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(235, 530)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'BtnData
        '
        Me.BtnData.Location = New System.Drawing.Point(46, 19)
        Me.BtnData.Name = "BtnData"
        Me.BtnData.Size = New System.Drawing.Size(29, 23)
        Me.BtnData.TabIndex = 14
        Me.BtnData.TabStop = False
        Me.BtnData.Text = "Data"
        Me.BtnData.UseVisualStyleBackColor = True
        Me.BtnData.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 46)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "To"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "From"
        '
        'DtpEnd
        '
        Me.DtpEnd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpEnd.CustomFormat = "dd MMM yyyy"
        Me.DtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpEnd.Location = New System.Drawing.Point(81, 42)
        Me.DtpEnd.Name = "DtpEnd"
        Me.DtpEnd.Size = New System.Drawing.Size(147, 20)
        Me.DtpEnd.TabIndex = 1
        '
        'DtpStart
        '
        Me.DtpStart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpStart.CustomFormat = "dd MMM yyyy"
        Me.DtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpStart.Location = New System.Drawing.Point(81, 19)
        Me.DtpStart.Name = "DtpStart"
        Me.DtpStart.Size = New System.Drawing.Size(147, 20)
        Me.DtpStart.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.Lbl_Departemen_Unchecked)
        Me.GroupBox5.Controls.Add(Me.Lbl_Departemen_Checked)
        Me.GroupBox5.Controls.Add(Me.ChkDepartemen)
        Me.GroupBox5.Controls.Add(Me.ClbDepartemen)
        Me.GroupBox5.Location = New System.Drawing.Point(7, 302)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(222, 109)
        Me.GroupBox5.TabIndex = 5
        Me.GroupBox5.TabStop = False
        '
        'Lbl_Departemen_Unchecked
        '
        Me.Lbl_Departemen_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Unchecked.AutoSize = True
        Me.Lbl_Departemen_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Unchecked.Location = New System.Drawing.Point(151, 0)
        Me.Lbl_Departemen_Unchecked.Name = "Lbl_Departemen_Unchecked"
        Me.Lbl_Departemen_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Departemen_Unchecked.TabIndex = 2
        Me.Lbl_Departemen_Unchecked.TabStop = True
        Me.Lbl_Departemen_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Departemen_Checked
        '
        Me.Lbl_Departemen_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Checked.AutoSize = True
        Me.Lbl_Departemen_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Checked.Location = New System.Drawing.Point(101, 0)
        Me.Lbl_Departemen_Checked.Name = "Lbl_Departemen_Checked"
        Me.Lbl_Departemen_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Departemen_Checked.TabIndex = 1
        Me.Lbl_Departemen_Checked.TabStop = True
        Me.Lbl_Departemen_Checked.Text = "Check All"
        '
        'ChkDepartemen
        '
        Me.ChkDepartemen.AutoSize = True
        Me.ChkDepartemen.Location = New System.Drawing.Point(11, -1)
        Me.ChkDepartemen.Name = "ChkDepartemen"
        Me.ChkDepartemen.Size = New System.Drawing.Size(81, 17)
        Me.ChkDepartemen.TabIndex = 0
        Me.ChkDepartemen.Text = "Department"
        Me.ChkDepartemen.UseVisualStyleBackColor = True
        '
        'ClbDepartemen
        '
        Me.ClbDepartemen.CheckOnClick = True
        Me.ClbDepartemen.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbDepartemen.FormattingEnabled = True
        Me.ClbDepartemen.Location = New System.Drawing.Point(3, 16)
        Me.ClbDepartemen.MultiColumn = True
        Me.ClbDepartemen.Name = "ClbDepartemen"
        Me.ClbDepartemen.Size = New System.Drawing.Size(216, 79)
        Me.ClbDepartemen.TabIndex = 1
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.Lbl_Position_Unchecked)
        Me.GroupBox4.Controls.Add(Me.Lbl_Position_Checked)
        Me.GroupBox4.Controls.Add(Me.ChkPosition)
        Me.GroupBox4.Controls.Add(Me.ClbPositon)
        Me.GroupBox4.Location = New System.Drawing.Point(7, 417)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(222, 104)
        Me.GroupBox4.TabIndex = 4
        Me.GroupBox4.TabStop = False
        '
        'Lbl_Position_Unchecked
        '
        Me.Lbl_Position_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Unchecked.AutoSize = True
        Me.Lbl_Position_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Unchecked.Location = New System.Drawing.Point(151, 0)
        Me.Lbl_Position_Unchecked.Name = "Lbl_Position_Unchecked"
        Me.Lbl_Position_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Position_Unchecked.TabIndex = 2
        Me.Lbl_Position_Unchecked.TabStop = True
        Me.Lbl_Position_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Position_Checked
        '
        Me.Lbl_Position_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Checked.AutoSize = True
        Me.Lbl_Position_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Checked.Location = New System.Drawing.Point(101, 0)
        Me.Lbl_Position_Checked.Name = "Lbl_Position_Checked"
        Me.Lbl_Position_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Position_Checked.TabIndex = 1
        Me.Lbl_Position_Checked.TabStop = True
        Me.Lbl_Position_Checked.Text = "Check All"
        '
        'ChkPosition
        '
        Me.ChkPosition.AutoSize = True
        Me.ChkPosition.Location = New System.Drawing.Point(11, -1)
        Me.ChkPosition.Name = "ChkPosition"
        Me.ChkPosition.Size = New System.Drawing.Size(63, 17)
        Me.ChkPosition.TabIndex = 0
        Me.ChkPosition.Text = "Position"
        Me.ChkPosition.UseVisualStyleBackColor = True
        '
        'ClbPositon
        '
        Me.ClbPositon.CheckOnClick = True
        Me.ClbPositon.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbPositon.FormattingEnabled = True
        Me.ClbPositon.Location = New System.Drawing.Point(3, 16)
        Me.ClbPositon.Name = "ClbPositon"
        Me.ClbPositon.Size = New System.Drawing.Size(216, 79)
        Me.ClbPositon.TabIndex = 1
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.Lbl_SF_Unchecked)
        Me.GroupBox3.Controls.Add(Me.Lbl_SF_Checked)
        Me.GroupBox3.Controls.Add(Me.ChkFormula)
        Me.GroupBox3.Controls.Add(Me.ClbSalaryFormula)
        Me.GroupBox3.Location = New System.Drawing.Point(6, 190)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(222, 106)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        '
        'Lbl_SF_Unchecked
        '
        Me.Lbl_SF_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Unchecked.AutoSize = True
        Me.Lbl_SF_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Unchecked.Location = New System.Drawing.Point(158, 0)
        Me.Lbl_SF_Unchecked.Name = "Lbl_SF_Unchecked"
        Me.Lbl_SF_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_SF_Unchecked.TabIndex = 2
        Me.Lbl_SF_Unchecked.TabStop = True
        Me.Lbl_SF_Unchecked.Text = "Uncheck All"
        '
        'Lbl_SF_Checked
        '
        Me.Lbl_SF_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Checked.AutoSize = True
        Me.Lbl_SF_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Checked.Location = New System.Drawing.Point(108, 0)
        Me.Lbl_SF_Checked.Name = "Lbl_SF_Checked"
        Me.Lbl_SF_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_SF_Checked.TabIndex = 1
        Me.Lbl_SF_Checked.TabStop = True
        Me.Lbl_SF_Checked.Text = "Check All"
        '
        'ChkFormula
        '
        Me.ChkFormula.AutoSize = True
        Me.ChkFormula.Location = New System.Drawing.Point(11, -1)
        Me.ChkFormula.Name = "ChkFormula"
        Me.ChkFormula.Size = New System.Drawing.Size(95, 17)
        Me.ChkFormula.TabIndex = 0
        Me.ChkFormula.Text = "Salary Formula"
        Me.ChkFormula.UseVisualStyleBackColor = True
        '
        'ClbSalaryFormula
        '
        Me.ClbSalaryFormula.CheckOnClick = True
        Me.ClbSalaryFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbSalaryFormula.FormattingEnabled = True
        Me.ClbSalaryFormula.Location = New System.Drawing.Point(3, 16)
        Me.ClbSalaryFormula.MultiColumn = True
        Me.ClbSalaryFormula.Name = "ClbSalaryFormula"
        Me.ClbSalaryFormula.Size = New System.Drawing.Size(216, 79)
        Me.ClbSalaryFormula.TabIndex = 1
        '
        'gboEmployee
        '
        Me.gboEmployee.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboEmployee.Controls.Add(Me.dgvEmployee)
        Me.gboEmployee.Controls.Add(Me.ChkEmployee)
        Me.gboEmployee.Location = New System.Drawing.Point(6, 62)
        Me.gboEmployee.Name = "gboEmployee"
        Me.gboEmployee.Size = New System.Drawing.Size(222, 122)
        Me.gboEmployee.TabIndex = 2
        Me.gboEmployee.TabStop = False
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColName, Me.ColNIK, Me.ColPrimaryCode})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(3, 16)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(216, 103)
        Me.dgvEmployee.TabIndex = 1
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.Width = 57
        '
        'ColName
        '
        Me.ColName.HeaderText = "Name"
        Me.ColName.Name = "ColName"
        Me.ColName.Width = 60
        '
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.Width = 50
        '
        'ColPrimaryCode
        '
        Me.ColPrimaryCode.HeaderText = "Primary"
        Me.ColPrimaryCode.Name = "ColPrimaryCode"
        Me.ColPrimaryCode.Visible = False
        Me.ColPrimaryCode.Width = 66
        '
        'ChkEmployee
        '
        Me.ChkEmployee.AutoSize = True
        Me.ChkEmployee.Location = New System.Drawing.Point(11, 0)
        Me.ChkEmployee.Name = "ChkEmployee"
        Me.ChkEmployee.Size = New System.Drawing.Size(72, 17)
        Me.ChkEmployee.TabIndex = 0
        Me.ChkEmployee.Text = "Employee"
        Me.ChkEmployee.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnDelete)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnAllocate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.BtnFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        Me.SplitContainer1.Panel1MinSize = 250
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.DgvWorkshift)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnClose)
        Me.SplitContainer1.Size = New System.Drawing.Size(878, 558)
        Me.SplitContainer1.SplitterDistance = 250
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'BtnDelete
        '
        Me.BtnDelete.Location = New System.Drawing.Point(169, 3)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(75, 23)
        Me.BtnDelete.TabIndex = 2
        Me.BtnDelete.Text = "Delete"
        Me.BtnDelete.UseVisualStyleBackColor = True
        '
        'BtnAllocate
        '
        Me.BtnAllocate.Location = New System.Drawing.Point(88, 3)
        Me.BtnAllocate.Name = "BtnAllocate"
        Me.BtnAllocate.Size = New System.Drawing.Size(75, 23)
        Me.BtnAllocate.TabIndex = 1
        Me.BtnAllocate.Text = "Allocate"
        Me.BtnAllocate.UseVisualStyleBackColor = True
        '
        'BtnFilter
        '
        Me.BtnFilter.Location = New System.Drawing.Point(7, 3)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(75, 23)
        Me.BtnFilter.TabIndex = 0
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'DgvWorkshift
        '
        Me.DgvWorkshift.AllowUserToAddRows = False
        Me.DgvWorkshift.AllowUserToDeleteRows = False
        Me.DgvWorkshift.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvWorkshift.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWorkshift.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColStartDate, Me.ColEndDate, Me.ColWorkshiftCode, Me.ColWorkShiftName, Me.ColRemarks, Me.ColSave, Me.ColDelete, Me.ColPrimaryWorkShift, Me.ColPrimaryEntity, Me.ColEntityTipe, Me.ColPrimary})
        Me.DgvWorkshift.Location = New System.Drawing.Point(0, 3)
        Me.DgvWorkshift.Name = "DgvWorkshift"
        Me.DgvWorkshift.Size = New System.Drawing.Size(621, 555)
        Me.DgvWorkshift.TabIndex = 0
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.Width = 106
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.Width = 109
        '
        'ColStartDate
        '
        Me.ColStartDate.HeaderText = "Start Date"
        Me.ColStartDate.Name = "ColStartDate"
        Me.ColStartDate.Width = 80
        '
        'ColEndDate
        '
        Me.ColEndDate.HeaderText = "End Date"
        Me.ColEndDate.Name = "ColEndDate"
        Me.ColEndDate.Width = 77
        '
        'ColWorkshiftCode
        '
        Me.ColWorkshiftCode.HeaderText = "Workshift Code"
        Me.ColWorkshiftCode.Name = "ColWorkshiftCode"
        Me.ColWorkshiftCode.Width = 105
        '
        'ColWorkShiftName
        '
        Me.ColWorkShiftName.HeaderText = "Workshift Name"
        Me.ColWorkShiftName.Name = "ColWorkShiftName"
        Me.ColWorkShiftName.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColWorkShiftName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWorkShiftName.Width = 89
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColSave
        '
        Me.ColSave.HeaderText = "Save"
        Me.ColSave.Name = "ColSave"
        Me.ColSave.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColSave.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColSave.Visible = False
        Me.ColSave.Width = 57
        '
        'ColDelete
        '
        Me.ColDelete.HeaderText = "Delete"
        Me.ColDelete.Name = "ColDelete"
        Me.ColDelete.Width = 44
        '
        'ColPrimaryWorkShift
        '
        Me.ColPrimaryWorkShift.HeaderText = "Primary Work Shift"
        Me.ColPrimaryWorkShift.Name = "ColPrimaryWorkShift"
        Me.ColPrimaryWorkShift.Visible = False
        Me.ColPrimaryWorkShift.Width = 119
        '
        'ColPrimaryEntity
        '
        Me.ColPrimaryEntity.HeaderText = "Primary Entity"
        Me.ColPrimaryEntity.Name = "ColPrimaryEntity"
        Me.ColPrimaryEntity.Visible = False
        Me.ColPrimaryEntity.Width = 95
        '
        'ColEntityTipe
        '
        Me.ColEntityTipe.HeaderText = "Entity Tipe"
        Me.ColEntityTipe.Name = "ColEntityTipe"
        Me.ColEntityTipe.Visible = False
        Me.ColEntityTipe.Width = 82
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 66
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(64, 103)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 4
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker3
        '
        '
        'BackgroundWorker4
        '
        '
        'BackgroundWorker5
        '
        '
        'BackgroundWorker_CloseForm
        '
        '
        'BackgroundWorker6
        '
        '
        'FrmMasterWorkShift_Allocate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(886, 559)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(902, 597)
        Me.Name = "FrmMasterWorkShift_Allocate"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Allocate Workshift"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gboEmployee.ResumeLayout(False)
        Me.gboEmployee.PerformLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.DgvWorkshift, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents gboEmployee As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWorkshift As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkPosition As System.Windows.Forms.CheckBox
    Friend WithEvents ClbPositon As System.Windows.Forms.CheckedListBox
    Friend WithEvents ChkFormula As System.Windows.Forms.CheckBox
    Friend WithEvents ClbSalaryFormula As System.Windows.Forms.CheckedListBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDepartemen As System.Windows.Forms.CheckBox
    Friend WithEvents ClbDepartemen As System.Windows.Forms.CheckedListBox
    Friend WithEvents BtnDelete As System.Windows.Forms.Button
    Friend WithEvents BtnAllocate As System.Windows.Forms.Button
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents ChkEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Lbl_Departemen_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Departemen_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Position_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Position_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_SF_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_SF_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker_CloseForm As System.ComponentModel.BackgroundWorker
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkshiftCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWorkShiftName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSave As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColDelete As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPrimaryWorkShift As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEntity As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEntityTipe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnData As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker6 As System.ComponentModel.BackgroundWorker
End Class
