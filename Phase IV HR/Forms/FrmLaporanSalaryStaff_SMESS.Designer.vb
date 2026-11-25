<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanSalaryStaff_SMESS
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
        Me.BtnClose = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label1 = New System.Windows.Forms.Label
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.SplitContainer6 = New System.Windows.Forms.SplitContainer
        Me.Lbl_Position_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Position_Checked = New System.Windows.Forms.LinkLabel
        Me.ChkPosition = New System.Windows.Forms.CheckBox
        Me.ClbPositon = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer
        Me.Lbl_Departemen_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Departemen_Checked = New System.Windows.Forms.LinkLabel
        Me.ChkDepartemen = New System.Windows.Forms.CheckBox
        Me.ClbDepartemen = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.Lbl_SF_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_SF_Checked = New System.Windows.Forms.LinkLabel
        Me.ChkFormula = New System.Windows.Forms.CheckBox
        Me.ClbSalaryFormula = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkEmployee = New System.Windows.Forms.CheckBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.crvRekap = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crvDetails = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer6.Panel1.SuspendLayout()
        Me.SplitContainer6.SuspendLayout()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(181, 32)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 20)
        Me.BtnClose.TabIndex = 3
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer1.Size = New System.Drawing.Size(986, 559)
        Me.SplitContainer1.SplitterDistance = 266
        Me.SplitContainer1.TabIndex = 1
        Me.SplitContainer1.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpTo)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnFilter)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnClose)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpFrom)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(259, 552)
        Me.SplitContainer2.SplitterDistance = 53
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(20, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "To"
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(40, 32)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(136, 20)
        Me.DtpTo.TabIndex = 1
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(181, 9)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(75, 20)
        Me.BtnFilter.TabIndex = 2
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "From"
        '
        'DtpFrom
        '
        Me.DtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpFrom.CustomFormat = "MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(40, 9)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(136, 20)
        Me.DtpFrom.TabIndex = 0
        '
        'SplitContainer6
        '
        Me.SplitContainer6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer6.Location = New System.Drawing.Point(0, 386)
        Me.SplitContainer6.Name = "SplitContainer6"
        Me.SplitContainer6.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer6.Panel1
        '
        Me.SplitContainer6.Panel1.Controls.Add(Me.Lbl_Position_Unchecked)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Lbl_Position_Checked)
        Me.SplitContainer6.Panel1.Controls.Add(Me.ChkPosition)
        Me.SplitContainer6.Panel1.Controls.Add(Me.ClbPositon)
        Me.SplitContainer6.Panel2Collapsed = True
        Me.SplitContainer6.Size = New System.Drawing.Size(259, 108)
        Me.SplitContainer6.SplitterDistance = 56
        Me.SplitContainer6.TabIndex = 3
        '
        'Lbl_Position_Unchecked
        '
        Me.Lbl_Position_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Unchecked.AutoSize = True
        Me.Lbl_Position_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Unchecked.Location = New System.Drawing.Point(192, 6)
        Me.Lbl_Position_Unchecked.Name = "Lbl_Position_Unchecked"
        Me.Lbl_Position_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Position_Unchecked.TabIndex = 16
        Me.Lbl_Position_Unchecked.TabStop = True
        Me.Lbl_Position_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Position_Checked
        '
        Me.Lbl_Position_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Checked.AutoSize = True
        Me.Lbl_Position_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Checked.Location = New System.Drawing.Point(133, 6)
        Me.Lbl_Position_Checked.Name = "Lbl_Position_Checked"
        Me.Lbl_Position_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Position_Checked.TabIndex = 15
        Me.Lbl_Position_Checked.TabStop = True
        Me.Lbl_Position_Checked.Text = "Check All"
        '
        'ChkPosition
        '
        Me.ChkPosition.AutoSize = True
        Me.ChkPosition.Location = New System.Drawing.Point(3, 6)
        Me.ChkPosition.Name = "ChkPosition"
        Me.ChkPosition.Size = New System.Drawing.Size(63, 17)
        Me.ChkPosition.TabIndex = 14
        Me.ChkPosition.Text = "Position"
        Me.ChkPosition.UseVisualStyleBackColor = True
        '
        'ClbPositon
        '
        Me.ClbPositon.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClbPositon.CheckOnClick = True
        Me.ClbPositon.FormattingEnabled = True
        Me.ClbPositon.Location = New System.Drawing.Point(3, 22)
        Me.ClbPositon.Name = "ClbPositon"
        Me.ClbPositon.Size = New System.Drawing.Size(253, 64)
        Me.ClbPositon.TabIndex = 17
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Top
        Me.SplitContainer5.Location = New System.Drawing.Point(0, 260)
        Me.SplitContainer5.Name = "SplitContainer5"
        Me.SplitContainer5.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.Lbl_Departemen_Unchecked)
        Me.SplitContainer5.Panel1.Controls.Add(Me.Lbl_Departemen_Checked)
        Me.SplitContainer5.Panel1.Controls.Add(Me.ChkDepartemen)
        Me.SplitContainer5.Panel1.Controls.Add(Me.ClbDepartemen)
        Me.SplitContainer5.Panel2Collapsed = True
        Me.SplitContainer5.Size = New System.Drawing.Size(259, 120)
        Me.SplitContainer5.SplitterDistance = 75
        Me.SplitContainer5.TabIndex = 2
        '
        'Lbl_Departemen_Unchecked
        '
        Me.Lbl_Departemen_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Unchecked.AutoSize = True
        Me.Lbl_Departemen_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Unchecked.Location = New System.Drawing.Point(192, 12)
        Me.Lbl_Departemen_Unchecked.Name = "Lbl_Departemen_Unchecked"
        Me.Lbl_Departemen_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_Departemen_Unchecked.TabIndex = 12
        Me.Lbl_Departemen_Unchecked.TabStop = True
        Me.Lbl_Departemen_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Departemen_Checked
        '
        Me.Lbl_Departemen_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Checked.AutoSize = True
        Me.Lbl_Departemen_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Checked.Location = New System.Drawing.Point(133, 12)
        Me.Lbl_Departemen_Checked.Name = "Lbl_Departemen_Checked"
        Me.Lbl_Departemen_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Departemen_Checked.TabIndex = 11
        Me.Lbl_Departemen_Checked.TabStop = True
        Me.Lbl_Departemen_Checked.Text = "Check All"
        '
        'ChkDepartemen
        '
        Me.ChkDepartemen.AutoSize = True
        Me.ChkDepartemen.Location = New System.Drawing.Point(3, 12)
        Me.ChkDepartemen.Name = "ChkDepartemen"
        Me.ChkDepartemen.Size = New System.Drawing.Size(81, 17)
        Me.ChkDepartemen.TabIndex = 10
        Me.ChkDepartemen.Text = "Department"
        Me.ChkDepartemen.UseVisualStyleBackColor = True
        '
        'ClbDepartemen
        '
        Me.ClbDepartemen.CheckOnClick = True
        Me.ClbDepartemen.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ClbDepartemen.FormattingEnabled = True
        Me.ClbDepartemen.Location = New System.Drawing.Point(0, 26)
        Me.ClbDepartemen.MultiColumn = True
        Me.ClbDepartemen.Name = "ClbDepartemen"
        Me.ClbDepartemen.Size = New System.Drawing.Size(259, 94)
        Me.ClbDepartemen.TabIndex = 13
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Top
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 140)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.Lbl_SF_Unchecked)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Lbl_SF_Checked)
        Me.SplitContainer4.Panel1.Controls.Add(Me.ChkFormula)
        Me.SplitContainer4.Panel1.Controls.Add(Me.ClbSalaryFormula)
        Me.SplitContainer4.Panel2Collapsed = True
        Me.SplitContainer4.Size = New System.Drawing.Size(259, 120)
        Me.SplitContainer4.SplitterDistance = 71
        Me.SplitContainer4.TabIndex = 1
        '
        'Lbl_SF_Unchecked
        '
        Me.Lbl_SF_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Unchecked.AutoSize = True
        Me.Lbl_SF_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Unchecked.Location = New System.Drawing.Point(192, 12)
        Me.Lbl_SF_Unchecked.Name = "Lbl_SF_Unchecked"
        Me.Lbl_SF_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_SF_Unchecked.TabIndex = 8
        Me.Lbl_SF_Unchecked.TabStop = True
        Me.Lbl_SF_Unchecked.Text = "Uncheck All"
        '
        'Lbl_SF_Checked
        '
        Me.Lbl_SF_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Checked.AutoSize = True
        Me.Lbl_SF_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Checked.Location = New System.Drawing.Point(133, 12)
        Me.Lbl_SF_Checked.Name = "Lbl_SF_Checked"
        Me.Lbl_SF_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_SF_Checked.TabIndex = 7
        Me.Lbl_SF_Checked.TabStop = True
        Me.Lbl_SF_Checked.Text = "Check All"
        '
        'ChkFormula
        '
        Me.ChkFormula.AutoSize = True
        Me.ChkFormula.Location = New System.Drawing.Point(3, 12)
        Me.ChkFormula.Name = "ChkFormula"
        Me.ChkFormula.Size = New System.Drawing.Size(95, 17)
        Me.ChkFormula.TabIndex = 6
        Me.ChkFormula.Text = "Salary Formula"
        Me.ChkFormula.UseVisualStyleBackColor = True
        '
        'ClbSalaryFormula
        '
        Me.ClbSalaryFormula.CheckOnClick = True
        Me.ClbSalaryFormula.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ClbSalaryFormula.FormattingEnabled = True
        Me.ClbSalaryFormula.Location = New System.Drawing.Point(0, 26)
        Me.ClbSalaryFormula.MultiColumn = True
        Me.ClbSalaryFormula.Name = "ClbSalaryFormula"
        Me.ClbSalaryFormula.Size = New System.Drawing.Size(259, 94)
        Me.ClbSalaryFormula.TabIndex = 9
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Top
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.dgvEmployee)
        Me.SplitContainer3.Panel1.Controls.Add(Me.ChkEmployee)
        Me.SplitContainer3.Panel2Collapsed = True
        Me.SplitContainer3.Size = New System.Drawing.Size(259, 140)
        Me.SplitContainer3.SplitterDistance = 25
        Me.SplitContainer3.TabIndex = 0
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColName, Me.ColNIK, Me.ColPrimaryCode})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dgvEmployee.Location = New System.Drawing.Point(0, 16)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(259, 124)
        Me.dgvEmployee.TabIndex = 5
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
        Me.ChkEmployee.Location = New System.Drawing.Point(3, -2)
        Me.ChkEmployee.Name = "ChkEmployee"
        Me.ChkEmployee.Size = New System.Drawing.Size(72, 17)
        Me.ChkEmployee.TabIndex = 4
        Me.ChkEmployee.Text = "Employee"
        Me.ChkEmployee.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(716, 559)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.crvRekap)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(708, 533)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Rekap"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'crvRekap
        '
        Me.crvRekap.ActiveViewIndex = -1
        Me.crvRekap.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvRekap.DisplayGroupTree = False
        Me.crvRekap.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvRekap.Location = New System.Drawing.Point(3, 3)
        Me.crvRekap.Name = "crvRekap"
        Me.crvRekap.SelectionFormula = ""
        Me.crvRekap.ShowCloseButton = False
        Me.crvRekap.ShowGroupTreeButton = False
        Me.crvRekap.ShowRefreshButton = False
        Me.crvRekap.Size = New System.Drawing.Size(702, 527)
        Me.crvRekap.TabIndex = 18
        Me.crvRekap.TabStop = False
        Me.crvRekap.ViewTimeSelectionFormula = ""
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crvDetails)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(708, 533)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Details"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'crvDetails
        '
        Me.crvDetails.ActiveViewIndex = -1
        Me.crvDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvDetails.DisplayGroupTree = False
        Me.crvDetails.DisplayStatusBar = False
        Me.crvDetails.DisplayToolbar = False
        Me.crvDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvDetails.Location = New System.Drawing.Point(3, 3)
        Me.crvDetails.Name = "crvDetails"
        Me.crvDetails.SelectionFormula = ""
        Me.crvDetails.ShowCloseButton = False
        Me.crvDetails.ShowGroupTreeButton = False
        Me.crvDetails.ShowRefreshButton = False
        Me.crvDetails.Size = New System.Drawing.Size(702, 527)
        Me.crvDetails.TabIndex = 19
        Me.crvDetails.TabStop = False
        Me.crvDetails.ViewTimeSelectionFormula = ""
        '
        'BackgroundWorker1
        '
        '
        'FrmLaporanSalaryStaff_SMESS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(986, 559)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1002, 597)
        Me.Name = "FrmLaporanSalaryStaff_SMESS"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Staff Salary Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer6.Panel1.ResumeLayout(False)
        Me.SplitContainer6.Panel1.PerformLayout()
        Me.SplitContainer6.ResumeLayout(False)
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel1.PerformLayout()
        Me.SplitContainer5.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer6 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer5 As System.Windows.Forms.SplitContainer
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents ClbSalaryFormula As System.Windows.Forms.CheckedListBox
    Friend WithEvents Lbl_SF_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_SF_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkFormula As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_Departemen_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Departemen_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkDepartemen As System.Windows.Forms.CheckBox
    Friend WithEvents ClbDepartemen As System.Windows.Forms.CheckedListBox
    Friend WithEvents Lbl_Position_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Position_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkPosition As System.Windows.Forms.CheckBox
    Friend WithEvents ClbPositon As System.Windows.Forms.CheckedListBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents crvRekap As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents crvDetails As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
