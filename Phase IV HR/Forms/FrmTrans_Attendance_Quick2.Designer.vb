<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Attendance_Quick2
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
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.RboAll = New System.Windows.Forms.RadioButton
        Me.RboSelected = New System.Windows.Forms.RadioButton
        Me.dtpPeriod = New System.Windows.Forms.DateTimePicker
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpTanggalTransaksi = New System.Windows.Forms.DateTimePicker
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.CboImportFromMachine = New System.Windows.Forms.CheckBox
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.chkJobPosition = New System.Windows.Forms.CheckBox
        Me.Lbl_JP_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_JP_checked = New System.Windows.Forms.LinkLabel
        Me.ClbPosition = New System.Windows.Forms.CheckedListBox
        Me.gboDepartment = New System.Windows.Forms.GroupBox
        Me.ChkDepartment = New System.Windows.Forms.CheckBox
        Me.Lbl_Dept_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Dept_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbDepartement = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.gboSalaryFormula = New System.Windows.Forms.GroupBox
        Me.chkSalaryFormula = New System.Windows.Forms.CheckBox
        Me.Lbl_SF_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_SF_Checked = New System.Windows.Forms.LinkLabel
        Me.clbSalaryFormula = New System.Windows.Forms.CheckedListBox
        Me.gboEmployee = New System.Windows.Forms.GroupBox
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryJob = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySF = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkEmployee = New System.Windows.Forms.CheckBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.LblProgress = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.BtnPause = New System.Windows.Forms.Button
        Me.ChkAttPrev = New System.Windows.Forms.CheckBox
        Me.gboDetail.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gboItem.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gboDepartment.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.gboSalaryFormula.SuspendLayout()
        Me.gboEmployee.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(533, 543)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(533, 544)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 19)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "Start"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(2, 0)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "What would you like to do?"
        '
        'RboAll
        '
        Me.RboAll.AutoSize = True
        Me.RboAll.Checked = True
        Me.RboAll.Location = New System.Drawing.Point(5, 15)
        Me.RboAll.Margin = New System.Windows.Forms.Padding(2)
        Me.RboAll.Name = "RboAll"
        Me.RboAll.Size = New System.Drawing.Size(156, 17)
        Me.RboAll.TabIndex = 0
        Me.RboAll.TabStop = True
        Me.RboAll.Text = "Attendance for all employee"
        Me.RboAll.UseVisualStyleBackColor = True
        '
        'RboSelected
        '
        Me.RboSelected.AutoSize = True
        Me.RboSelected.Location = New System.Drawing.Point(5, 33)
        Me.RboSelected.Margin = New System.Windows.Forms.Padding(2)
        Me.RboSelected.Name = "RboSelected"
        Me.RboSelected.Size = New System.Drawing.Size(186, 17)
        Me.RboSelected.TabIndex = 1
        Me.RboSelected.Text = "Attendance for selected employee"
        Me.RboSelected.UseVisualStyleBackColor = True
        '
        'dtpPeriod
        '
        Me.dtpPeriod.CustomFormat = "dd MMM yyyy"
        Me.dtpPeriod.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPeriod.Location = New System.Drawing.Point(104, 30)
        Me.dtpPeriod.Name = "dtpPeriod"
        Me.dtpPeriod.Size = New System.Drawing.Size(176, 20)
        Me.dtpPeriod.TabIndex = 1
        Me.dtpPeriod.Value = New Date(2016, 6, 20, 0, 0, 0, 0)
        '
        'DtpTo
        '
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(104, 56)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(176, 20)
        Me.DtpTo.TabIndex = 2
        Me.DtpTo.Value = New Date(2016, 6, 20, 0, 0, 0, 0)
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 49
        Me.Label4.Text = "To"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 30)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(37, 13)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "Period"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "Date"
        '
        'DtpTanggalTransaksi
        '
        Me.DtpTanggalTransaksi.CustomFormat = "dd MMM yyyy"
        Me.DtpTanggalTransaksi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTanggalTransaksi.Location = New System.Drawing.Point(104, 4)
        Me.DtpTanggalTransaksi.Name = "DtpTanggalTransaksi"
        Me.DtpTanggalTransaksi.Size = New System.Drawing.Size(176, 20)
        Me.DtpTanggalTransaksi.TabIndex = 0
        '
        'gboDetail
        '
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Location = New System.Drawing.Point(15, 1)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(580, 110)
        Me.gboDetail.TabIndex = 0
        Me.gboDetail.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboImportFromMachine)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboSelected)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RboAll)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel2.Controls.Add(Me.DtpTanggalTransaksi)
        Me.SplitContainer1.Panel2.Controls.Add(Me.DtpTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.dtpPeriod)
        Me.SplitContainer1.Size = New System.Drawing.Size(574, 91)
        Me.SplitContainer1.SplitterDistance = 286
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'CboImportFromMachine
        '
        Me.CboImportFromMachine.AutoSize = True
        Me.CboImportFromMachine.Checked = True
        Me.CboImportFromMachine.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CboImportFromMachine.Location = New System.Drawing.Point(4, 53)
        Me.CboImportFromMachine.Name = "CboImportFromMachine"
        Me.CboImportFromMachine.Size = New System.Drawing.Size(187, 17)
        Me.CboImportFromMachine.TabIndex = 4
        Me.CboImportFromMachine.Text = "Import log from fingerprint machine"
        Me.CboImportFromMachine.UseVisualStyleBackColor = True
        '
        'gboItem
        '
        Me.gboItem.Controls.Add(Me.SplitContainer3)
        Me.gboItem.Location = New System.Drawing.Point(15, 113)
        Me.gboItem.Margin = New System.Windows.Forms.Padding(2)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Padding = New System.Windows.Forms.Padding(2)
        Me.gboItem.Size = New System.Drawing.Size(580, 425)
        Me.gboItem.TabIndex = 1
        Me.gboItem.TabStop = False
        Me.gboItem.Text = "Data Filter"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Location = New System.Drawing.Point(5, 18)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Size = New System.Drawing.Size(569, 408)
        Me.SplitContainer3.SplitterDistance = 281
        Me.SplitContainer3.TabIndex = 3
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.GroupBox3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.gboDepartment)
        Me.SplitContainer2.Size = New System.Drawing.Size(281, 408)
        Me.SplitContainer2.SplitterDistance = 196
        Me.SplitContainer2.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.chkJobPosition)
        Me.GroupBox3.Controls.Add(Me.Lbl_JP_Unchecked)
        Me.GroupBox3.Controls.Add(Me.Lbl_JP_checked)
        Me.GroupBox3.Controls.Add(Me.ClbPosition)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(281, 196)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        '
        'chkJobPosition
        '
        Me.chkJobPosition.AutoSize = True
        Me.chkJobPosition.Location = New System.Drawing.Point(6, -1)
        Me.chkJobPosition.Name = "chkJobPosition"
        Me.chkJobPosition.Size = New System.Drawing.Size(83, 17)
        Me.chkJobPosition.TabIndex = 0
        Me.chkJobPosition.Text = "Job Position"
        Me.chkJobPosition.UseVisualStyleBackColor = True
        '
        'Lbl_JP_Unchecked
        '
        Me.Lbl_JP_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JP_Unchecked.AutoSize = True
        Me.Lbl_JP_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JP_Unchecked.Location = New System.Drawing.Point(210, 0)
        Me.Lbl_JP_Unchecked.Name = "Lbl_JP_Unchecked"
        Me.Lbl_JP_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_JP_Unchecked.TabIndex = 2
        Me.Lbl_JP_Unchecked.TabStop = True
        Me.Lbl_JP_Unchecked.Text = "Uncheck All"
        '
        'Lbl_JP_checked
        '
        Me.Lbl_JP_checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JP_checked.AutoSize = True
        Me.Lbl_JP_checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JP_checked.Location = New System.Drawing.Point(152, 0)
        Me.Lbl_JP_checked.Name = "Lbl_JP_checked"
        Me.Lbl_JP_checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_JP_checked.TabIndex = 1
        Me.Lbl_JP_checked.TabStop = True
        Me.Lbl_JP_checked.Text = "Check All"
        '
        'ClbPosition
        '
        Me.ClbPosition.CheckOnClick = True
        Me.ClbPosition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbPosition.FormattingEnabled = True
        Me.ClbPosition.Location = New System.Drawing.Point(3, 16)
        Me.ClbPosition.Name = "ClbPosition"
        Me.ClbPosition.Size = New System.Drawing.Size(275, 169)
        Me.ClbPosition.TabIndex = 3
        '
        'gboDepartment
        '
        Me.gboDepartment.Controls.Add(Me.ChkDepartment)
        Me.gboDepartment.Controls.Add(Me.Lbl_Dept_Unchecked)
        Me.gboDepartment.Controls.Add(Me.Lbl_Dept_Checked)
        Me.gboDepartment.Controls.Add(Me.ClbDepartement)
        Me.gboDepartment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboDepartment.Location = New System.Drawing.Point(0, 0)
        Me.gboDepartment.Name = "gboDepartment"
        Me.gboDepartment.Size = New System.Drawing.Size(281, 208)
        Me.gboDepartment.TabIndex = 0
        Me.gboDepartment.TabStop = False
        '
        'ChkDepartment
        '
        Me.ChkDepartment.AutoSize = True
        Me.ChkDepartment.Location = New System.Drawing.Point(6, -1)
        Me.ChkDepartment.Name = "ChkDepartment"
        Me.ChkDepartment.Size = New System.Drawing.Size(81, 17)
        Me.ChkDepartment.TabIndex = 0
        Me.ChkDepartment.Text = "Department"
        Me.ChkDepartment.UseVisualStyleBackColor = True
        '
        'Lbl_Dept_Unchecked
        '
        Me.Lbl_Dept_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Dept_Unchecked.AutoSize = True
        Me.Lbl_Dept_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Dept_Unchecked.Location = New System.Drawing.Point(210, 0)
        Me.Lbl_Dept_Unchecked.Name = "Lbl_Dept_Unchecked"
        Me.Lbl_Dept_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Dept_Unchecked.TabIndex = 2
        Me.Lbl_Dept_Unchecked.TabStop = True
        Me.Lbl_Dept_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Dept_Checked
        '
        Me.Lbl_Dept_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Dept_Checked.AutoSize = True
        Me.Lbl_Dept_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Dept_Checked.Location = New System.Drawing.Point(152, 0)
        Me.Lbl_Dept_Checked.Name = "Lbl_Dept_Checked"
        Me.Lbl_Dept_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Dept_Checked.TabIndex = 1
        Me.Lbl_Dept_Checked.TabStop = True
        Me.Lbl_Dept_Checked.Text = "Check All"
        '
        'ClbDepartement
        '
        Me.ClbDepartement.CheckOnClick = True
        Me.ClbDepartement.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ClbDepartement.FormattingEnabled = True
        Me.ClbDepartement.Location = New System.Drawing.Point(3, 16)
        Me.ClbDepartement.Name = "ClbDepartement"
        Me.ClbDepartement.Size = New System.Drawing.Size(275, 184)
        Me.ClbDepartement.TabIndex = 3
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.gboSalaryFormula)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.gboEmployee)
        Me.SplitContainer4.Size = New System.Drawing.Size(284, 408)
        Me.SplitContainer4.SplitterDistance = 196
        Me.SplitContainer4.TabIndex = 0
        '
        'gboSalaryFormula
        '
        Me.gboSalaryFormula.Controls.Add(Me.chkSalaryFormula)
        Me.gboSalaryFormula.Controls.Add(Me.Lbl_SF_Unchecked)
        Me.gboSalaryFormula.Controls.Add(Me.Lbl_SF_Checked)
        Me.gboSalaryFormula.Controls.Add(Me.clbSalaryFormula)
        Me.gboSalaryFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboSalaryFormula.Location = New System.Drawing.Point(0, 0)
        Me.gboSalaryFormula.Name = "gboSalaryFormula"
        Me.gboSalaryFormula.Size = New System.Drawing.Size(284, 196)
        Me.gboSalaryFormula.TabIndex = 0
        Me.gboSalaryFormula.TabStop = False
        '
        'chkSalaryFormula
        '
        Me.chkSalaryFormula.AutoSize = True
        Me.chkSalaryFormula.Location = New System.Drawing.Point(6, -1)
        Me.chkSalaryFormula.Name = "chkSalaryFormula"
        Me.chkSalaryFormula.Size = New System.Drawing.Size(95, 17)
        Me.chkSalaryFormula.TabIndex = 0
        Me.chkSalaryFormula.Text = "Salary Formula"
        Me.chkSalaryFormula.UseVisualStyleBackColor = True
        '
        'Lbl_SF_Unchecked
        '
        Me.Lbl_SF_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Unchecked.AutoSize = True
        Me.Lbl_SF_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Unchecked.Location = New System.Drawing.Point(214, 0)
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
        Me.Lbl_SF_Checked.Location = New System.Drawing.Point(155, 0)
        Me.Lbl_SF_Checked.Name = "Lbl_SF_Checked"
        Me.Lbl_SF_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_SF_Checked.TabIndex = 1
        Me.Lbl_SF_Checked.TabStop = True
        Me.Lbl_SF_Checked.Text = "Check All"
        '
        'clbSalaryFormula
        '
        Me.clbSalaryFormula.CheckOnClick = True
        Me.clbSalaryFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbSalaryFormula.FormattingEnabled = True
        Me.clbSalaryFormula.Location = New System.Drawing.Point(3, 16)
        Me.clbSalaryFormula.Name = "clbSalaryFormula"
        Me.clbSalaryFormula.Size = New System.Drawing.Size(278, 169)
        Me.clbSalaryFormula.TabIndex = 3
        '
        'gboEmployee
        '
        Me.gboEmployee.Controls.Add(Me.dgvEmployee)
        Me.gboEmployee.Controls.Add(Me.ChkEmployee)
        Me.gboEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboEmployee.Location = New System.Drawing.Point(0, 0)
        Me.gboEmployee.Name = "gboEmployee"
        Me.gboEmployee.Size = New System.Drawing.Size(284, 208)
        Me.gboEmployee.TabIndex = 0
        Me.gboEmployee.TabStop = False
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColName, Me.ColNIK, Me.ColPrimaryCode, Me.ColPrimaryJob, Me.ColPrimaryDept, Me.ColPrimarySF})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(3, 16)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(278, 189)
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
        Me.ColName.ReadOnly = True
        Me.ColName.Width = 60
        '
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.ReadOnly = True
        Me.ColNIK.Width = 50
        '
        'ColPrimaryCode
        '
        Me.ColPrimaryCode.HeaderText = "Primary"
        Me.ColPrimaryCode.Name = "ColPrimaryCode"
        Me.ColPrimaryCode.Visible = False
        Me.ColPrimaryCode.Width = 66
        '
        'ColPrimaryJob
        '
        Me.ColPrimaryJob.HeaderText = "Primary Job"
        Me.ColPrimaryJob.Name = "ColPrimaryJob"
        Me.ColPrimaryJob.Visible = False
        Me.ColPrimaryJob.Width = 86
        '
        'ColPrimaryDept
        '
        Me.ColPrimaryDept.HeaderText = "Primary Dept"
        Me.ColPrimaryDept.Name = "ColPrimaryDept"
        Me.ColPrimaryDept.Visible = False
        Me.ColPrimaryDept.Width = 92
        '
        'ColPrimarySF
        '
        Me.ColPrimarySF.HeaderText = "Primary SF"
        Me.ColPrimarySF.Name = "ColPrimarySF"
        Me.ColPrimarySF.Visible = False
        Me.ColPrimarySF.Width = 82
        '
        'ChkEmployee
        '
        Me.ChkEmployee.AutoSize = True
        Me.ChkEmployee.Location = New System.Drawing.Point(6, -1)
        Me.ChkEmployee.Name = "ChkEmployee"
        Me.ChkEmployee.Size = New System.Drawing.Size(72, 17)
        Me.ChkEmployee.TabIndex = 0
        Me.ChkEmployee.Text = "Employee"
        Me.ChkEmployee.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'LblProgress
        '
        Me.LblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(23, 564)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(39, 13)
        Me.LblProgress.TabIndex = 67
        Me.LblProgress.Text = "Label9"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(23, 543)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(505, 18)
        Me.ProgressBar1.TabIndex = 66
        '
        'BtnPause
        '
        Me.BtnPause.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnPause.Location = New System.Drawing.Point(533, 544)
        Me.BtnPause.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnPause.Name = "BtnPause"
        Me.BtnPause.Size = New System.Drawing.Size(56, 19)
        Me.BtnPause.TabIndex = 68
        Me.BtnPause.Text = "Pause"
        Me.BtnPause.UseVisualStyleBackColor = True
        '
        'ChkAttPrev
        '
        Me.ChkAttPrev.AutoSize = True
        Me.ChkAttPrev.Checked = True
        Me.ChkAttPrev.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkAttPrev.Location = New System.Drawing.Point(22, 89)
        Me.ChkAttPrev.Name = "ChkAttPrev"
        Me.ChkAttPrev.Size = New System.Drawing.Size(191, 17)
        Me.ChkAttPrev.TabIndex = 5
        Me.ChkAttPrev.Text = "Continue from previous attendance"
        Me.ChkAttPrev.UseVisualStyleBackColor = True
        '
        'FrmTrans_Attendance_Quick2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(599, 581)
        Me.Controls.Add(Me.ChkAttPrev)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.gboItem)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.BtnPause)
        Me.Controls.Add(Me.btnCancel)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(615, 619)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(615, 619)
        Me.Name = "FrmTrans_Attendance_Quick2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Filter Attendance"
        Me.gboDetail.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.gboItem.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gboDepartment.ResumeLayout(False)
        Me.gboDepartment.PerformLayout()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        Me.gboSalaryFormula.ResumeLayout(False)
        Me.gboSalaryFormula.PerformLayout()
        Me.gboEmployee.ResumeLayout(False)
        Me.gboEmployee.PerformLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RboAll As System.Windows.Forms.RadioButton
    Friend WithEvents RboSelected As System.Windows.Forms.RadioButton
    Friend WithEvents dtpPeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpTanggalTransaksi As System.Windows.Forms.DateTimePicker
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents chkJobPosition As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_JP_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_JP_checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ClbPosition As System.Windows.Forms.CheckedListBox
    Friend WithEvents gboDepartment As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDepartment As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_Dept_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Dept_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ClbDepartement As System.Windows.Forms.CheckedListBox
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents gboSalaryFormula As System.Windows.Forms.GroupBox
    Friend WithEvents chkSalaryFormula As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_SF_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_SF_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents clbSalaryFormula As System.Windows.Forms.CheckedListBox
    Friend WithEvents gboEmployee As System.Windows.Forms.GroupBox
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents ChkEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryJob As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySF As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BtnPause As System.Windows.Forms.Button
    Friend WithEvents CboImportFromMachine As System.Windows.Forms.CheckBox
    Friend WithEvents ChkAttPrev As System.Windows.Forms.CheckBox
End Class
