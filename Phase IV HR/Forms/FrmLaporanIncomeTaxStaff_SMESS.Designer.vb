<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanIncomeTaxStaff_SMESS
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
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpStart = New System.Windows.Forms.DateTimePicker
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
        Me.crvTaxIncome = New CrystalDecisions.Windows.Forms.CrystalReportViewer
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
        Me.SuspendLayout()
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(240, 39)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 25)
        Me.BtnClose.TabIndex = 3
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crvTaxIncome)
        Me.SplitContainer1.Size = New System.Drawing.Size(1323, 779)
        Me.SplitContainer1.SplitterDistance = 354
        Me.SplitContainer1.SplitterWidth = 5
        Me.SplitContainer1.TabIndex = 1
        Me.SplitContainer1.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(4, 4)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnFilter)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BtnClose)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.DtpStart)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer5)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Size = New System.Drawing.Size(344, 771)
        Me.SplitContainer2.SplitterDistance = 65
        Me.SplitContainer2.SplitterWidth = 5
        Me.SplitContainer2.TabIndex = 0
        Me.SplitContainer2.TabStop = False
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(240, 11)
        Me.BtnFilter.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(100, 25)
        Me.BtnFilter.TabIndex = 2
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 16)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 17)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Period"
        '
        'DtpStart
        '
        Me.DtpStart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpStart.CustomFormat = "MMM yyyy"
        Me.DtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpStart.Location = New System.Drawing.Point(101, 11)
        Me.DtpStart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.DtpStart.Name = "DtpStart"
        Me.DtpStart.Size = New System.Drawing.Size(130, 22)
        Me.DtpStart.TabIndex = 0
        '
        'SplitContainer6
        '
        Me.SplitContainer6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer6.Location = New System.Drawing.Point(0, 517)
        Me.SplitContainer6.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.SplitContainer6.Size = New System.Drawing.Size(344, 183)
        Me.SplitContainer6.SplitterDistance = 107
        Me.SplitContainer6.SplitterWidth = 5
        Me.SplitContainer6.TabIndex = 3
        '
        'Lbl_Position_Unchecked
        '
        Me.Lbl_Position_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Unchecked.AutoSize = True
        Me.Lbl_Position_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Unchecked.Location = New System.Drawing.Point(254, 7)
        Me.Lbl_Position_Unchecked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Position_Unchecked.Name = "Lbl_Position_Unchecked"
        Me.Lbl_Position_Unchecked.Size = New System.Drawing.Size(82, 17)
        Me.Lbl_Position_Unchecked.TabIndex = 16
        Me.Lbl_Position_Unchecked.TabStop = True
        Me.Lbl_Position_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Position_Checked
        '
        Me.Lbl_Position_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Checked.AutoSize = True
        Me.Lbl_Position_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Checked.Location = New System.Drawing.Point(176, 7)
        Me.Lbl_Position_Checked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Position_Checked.Name = "Lbl_Position_Checked"
        Me.Lbl_Position_Checked.Size = New System.Drawing.Size(66, 17)
        Me.Lbl_Position_Checked.TabIndex = 15
        Me.Lbl_Position_Checked.TabStop = True
        Me.Lbl_Position_Checked.Text = "Check All"
        '
        'ChkPosition
        '
        Me.ChkPosition.AutoSize = True
        Me.ChkPosition.Location = New System.Drawing.Point(4, 7)
        Me.ChkPosition.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkPosition.Name = "ChkPosition"
        Me.ChkPosition.Size = New System.Drawing.Size(80, 21)
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
        Me.ClbPositon.Location = New System.Drawing.Point(0, 30)
        Me.ClbPositon.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ClbPositon.Name = "ClbPositon"
        Me.ClbPositon.Size = New System.Drawing.Size(343, 140)
        Me.ClbPositon.TabIndex = 17
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Top
        Me.SplitContainer5.Location = New System.Drawing.Point(0, 344)
        Me.SplitContainer5.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.SplitContainer5.Size = New System.Drawing.Size(344, 172)
        Me.SplitContainer5.SplitterDistance = 75
        Me.SplitContainer5.SplitterWidth = 5
        Me.SplitContainer5.TabIndex = 2
        '
        'Lbl_Departemen_Unchecked
        '
        Me.Lbl_Departemen_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Unchecked.AutoSize = True
        Me.Lbl_Departemen_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Unchecked.Location = New System.Drawing.Point(254, 15)
        Me.Lbl_Departemen_Unchecked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Departemen_Unchecked.Name = "Lbl_Departemen_Unchecked"
        Me.Lbl_Departemen_Unchecked.Size = New System.Drawing.Size(82, 17)
        Me.Lbl_Departemen_Unchecked.TabIndex = 12
        Me.Lbl_Departemen_Unchecked.TabStop = True
        Me.Lbl_Departemen_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Departemen_Checked
        '
        Me.Lbl_Departemen_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Checked.AutoSize = True
        Me.Lbl_Departemen_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Checked.Location = New System.Drawing.Point(176, 15)
        Me.Lbl_Departemen_Checked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Departemen_Checked.Name = "Lbl_Departemen_Checked"
        Me.Lbl_Departemen_Checked.Size = New System.Drawing.Size(66, 17)
        Me.Lbl_Departemen_Checked.TabIndex = 11
        Me.Lbl_Departemen_Checked.TabStop = True
        Me.Lbl_Departemen_Checked.Text = "Check All"
        '
        'ChkDepartemen
        '
        Me.ChkDepartemen.AutoSize = True
        Me.ChkDepartemen.Location = New System.Drawing.Point(4, 15)
        Me.ChkDepartemen.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkDepartemen.Name = "ChkDepartemen"
        Me.ChkDepartemen.Size = New System.Drawing.Size(104, 21)
        Me.ChkDepartemen.TabIndex = 10
        Me.ChkDepartemen.Text = "Department"
        Me.ChkDepartemen.UseVisualStyleBackColor = True
        '
        'ClbDepartemen
        '
        Me.ClbDepartemen.CheckOnClick = True
        Me.ClbDepartemen.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ClbDepartemen.FormattingEnabled = True
        Me.ClbDepartemen.Location = New System.Drawing.Point(0, 49)
        Me.ClbDepartemen.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ClbDepartemen.MultiColumn = True
        Me.ClbDepartemen.Name = "ClbDepartemen"
        Me.ClbDepartemen.Size = New System.Drawing.Size(344, 123)
        Me.ClbDepartemen.TabIndex = 13
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Top
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 172)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.SplitContainer4.Size = New System.Drawing.Size(344, 172)
        Me.SplitContainer4.SplitterDistance = 71
        Me.SplitContainer4.SplitterWidth = 5
        Me.SplitContainer4.TabIndex = 1
        '
        'Lbl_SF_Unchecked
        '
        Me.Lbl_SF_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Unchecked.AutoSize = True
        Me.Lbl_SF_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Unchecked.Location = New System.Drawing.Point(254, 15)
        Me.Lbl_SF_Unchecked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_SF_Unchecked.Name = "Lbl_SF_Unchecked"
        Me.Lbl_SF_Unchecked.Size = New System.Drawing.Size(82, 17)
        Me.Lbl_SF_Unchecked.TabIndex = 8
        Me.Lbl_SF_Unchecked.TabStop = True
        Me.Lbl_SF_Unchecked.Text = "Uncheck All"
        '
        'Lbl_SF_Checked
        '
        Me.Lbl_SF_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Checked.AutoSize = True
        Me.Lbl_SF_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Checked.Location = New System.Drawing.Point(176, 15)
        Me.Lbl_SF_Checked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_SF_Checked.Name = "Lbl_SF_Checked"
        Me.Lbl_SF_Checked.Size = New System.Drawing.Size(66, 17)
        Me.Lbl_SF_Checked.TabIndex = 7
        Me.Lbl_SF_Checked.TabStop = True
        Me.Lbl_SF_Checked.Text = "Check All"
        '
        'ChkFormula
        '
        Me.ChkFormula.AutoSize = True
        Me.ChkFormula.Location = New System.Drawing.Point(4, 15)
        Me.ChkFormula.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkFormula.Name = "ChkFormula"
        Me.ChkFormula.Size = New System.Drawing.Size(125, 21)
        Me.ChkFormula.TabIndex = 6
        Me.ChkFormula.Text = "Salary Formula"
        Me.ChkFormula.UseVisualStyleBackColor = True
        '
        'ClbSalaryFormula
        '
        Me.ClbSalaryFormula.CheckOnClick = True
        Me.ClbSalaryFormula.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ClbSalaryFormula.FormattingEnabled = True
        Me.ClbSalaryFormula.Location = New System.Drawing.Point(0, 49)
        Me.ClbSalaryFormula.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ClbSalaryFormula.MultiColumn = True
        Me.ClbSalaryFormula.Name = "ClbSalaryFormula"
        Me.ClbSalaryFormula.Size = New System.Drawing.Size(344, 123)
        Me.ClbSalaryFormula.TabIndex = 9
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Top
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.dgvEmployee)
        Me.SplitContainer3.Panel1.Controls.Add(Me.ChkEmployee)
        Me.SplitContainer3.Panel2Collapsed = True
        Me.SplitContainer3.Size = New System.Drawing.Size(344, 172)
        Me.SplitContainer3.SplitterDistance = 25
        Me.SplitContainer3.SplitterWidth = 5
        Me.SplitContainer3.TabIndex = 0
        '
        'dgvEmployee
        '
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColName, Me.ColNIK, Me.ColPrimaryCode})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dgvEmployee.Location = New System.Drawing.Point(0, 19)
        Me.dgvEmployee.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(344, 153)
        Me.dgvEmployee.TabIndex = 5
        '
        'ColCode
        '
        Me.ColCode.HeaderText = "Code"
        Me.ColCode.Name = "ColCode"
        Me.ColCode.Width = 66
        '
        'ColName
        '
        Me.ColName.HeaderText = "Name"
        Me.ColName.Name = "ColName"
        Me.ColName.Width = 70
        '
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.Width = 55
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
        Me.ChkEmployee.Location = New System.Drawing.Point(4, -2)
        Me.ChkEmployee.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ChkEmployee.Name = "ChkEmployee"
        Me.ChkEmployee.Size = New System.Drawing.Size(92, 21)
        Me.ChkEmployee.TabIndex = 4
        Me.ChkEmployee.Text = "Employee"
        Me.ChkEmployee.UseVisualStyleBackColor = True
        '
        'crvTaxIncome
        '
        Me.crvTaxIncome.ActiveViewIndex = -1
        Me.crvTaxIncome.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvTaxIncome.DisplayGroupTree = False
        Me.crvTaxIncome.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvTaxIncome.Location = New System.Drawing.Point(0, 0)
        Me.crvTaxIncome.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.crvTaxIncome.Name = "crvTaxIncome"
        Me.crvTaxIncome.SelectionFormula = ""
        Me.crvTaxIncome.ShowCloseButton = False
        Me.crvTaxIncome.ShowGroupTreeButton = False
        Me.crvTaxIncome.ShowRefreshButton = False
        Me.crvTaxIncome.Size = New System.Drawing.Size(964, 779)
        Me.crvTaxIncome.TabIndex = 18
        Me.crvTaxIncome.TabStop = False
        Me.crvTaxIncome.ViewTimeSelectionFormula = ""
        '
        'BackgroundWorker1
        '
        '
        'FrmLaporanIncomeTaxStaff_SMESS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(1323, 779)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1331, 816)
        Me.Name = "FrmLaporanIncomeTaxStaff_SMESS"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Laporan Salary Staff"
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
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpStart As System.Windows.Forms.DateTimePicker
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
    Friend WithEvents crvTaxIncome As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
