<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanIncomeTaxStaff
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpEnd = New System.Windows.Forms.DateTimePicker
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.BtnClose = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.DtpStart = New System.Windows.Forms.DateTimePicker
        Me.ChkEmployee = New System.Windows.Forms.CheckBox
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ChkFormula = New System.Windows.Forms.CheckBox
        Me.Lbl_SF_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_SF_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbSalaryFormula = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.ChkDepartemen = New System.Windows.Forms.CheckBox
        Me.Lbl_Departemen_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Departemen_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbDepartemen = New System.Windows.Forms.CheckedListBox
        Me.ChkPosition = New System.Windows.Forms.CheckBox
        Me.Lbl_Position_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Position_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbPositon = New System.Windows.Forms.CheckedListBox
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.SplitContainer2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer1.Size = New System.Drawing.Size(992, 633)
        Me.SplitContainer1.SplitterDistance = 492
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.ChkFormula)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Lbl_SF_Unchecked)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Lbl_SF_Checked)
        Me.SplitContainer2.Panel2.Controls.Add(Me.ClbSalaryFormula)
        Me.SplitContainer2.Size = New System.Drawing.Size(492, 633)
        Me.SplitContainer2.SplitterDistance = 350
        Me.SplitContainer2.TabIndex = 1
        Me.SplitContainer2.TabStop = False
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.IsSplitterFixed = True
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer3.Panel1.Controls.Add(Me.cboFormat)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer3.Panel1.Controls.Add(Me.dtpEnd)
        Me.SplitContainer3.Panel1.Controls.Add(Me.BtnFilter)
        Me.SplitContainer3.Panel1.Controls.Add(Me.BtnClose)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer3.Panel1.Controls.Add(Me.DtpStart)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.ChkEmployee)
        Me.SplitContainer3.Panel2.Controls.Add(Me.dgvEmployee)
        Me.SplitContainer3.Size = New System.Drawing.Size(492, 350)
        Me.SplitContainer3.SplitterDistance = 100
        Me.SplitContainer3.TabIndex = 1
        Me.SplitContainer3.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 15)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(52, 17)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Format"
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Recap", "Detail"})
        Me.cboFormat.Location = New System.Drawing.Point(72, 12)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(308, 24)
        Me.cboFormat.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 78)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(25, 17)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "To"
        '
        'dtpEnd
        '
        Me.dtpEnd.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpEnd.CustomFormat = "MMM yyyy"
        Me.dtpEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEnd.Location = New System.Drawing.Point(72, 73)
        Me.dtpEnd.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpEnd.Name = "dtpEnd"
        Me.dtpEnd.ShowUpDown = True
        Me.dtpEnd.Size = New System.Drawing.Size(308, 22)
        Me.dtpEnd.TabIndex = 2
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(388, 13)
        Me.BtnFilter.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(100, 25)
        Me.BtnFilter.TabIndex = 3
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(388, 42)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 25)
        Me.BtnClose.TabIndex = 4
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 46)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 17)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "From"
        '
        'DtpStart
        '
        Me.DtpStart.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpStart.CustomFormat = "MMM yyyy"
        Me.DtpStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpStart.Location = New System.Drawing.Point(72, 43)
        Me.DtpStart.Margin = New System.Windows.Forms.Padding(4)
        Me.DtpStart.Name = "DtpStart"
        Me.DtpStart.ShowUpDown = True
        Me.DtpStart.Size = New System.Drawing.Size(308, 22)
        Me.DtpStart.TabIndex = 1
        '
        'ChkEmployee
        '
        Me.ChkEmployee.AutoSize = True
        Me.ChkEmployee.Location = New System.Drawing.Point(4, 4)
        Me.ChkEmployee.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkEmployee.Name = "ChkEmployee"
        Me.ChkEmployee.Size = New System.Drawing.Size(135, 21)
        Me.ChkEmployee.TabIndex = 0
        Me.ChkEmployee.Text = "Select Employee"
        Me.ChkEmployee.UseVisualStyleBackColor = True
        '
        'dgvEmployee
        '
        Me.dgvEmployee.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvEmployee.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCode, Me.ColName, Me.ColNIK, Me.ColPrimaryCode})
        Me.dgvEmployee.Location = New System.Drawing.Point(4, 25)
        Me.dgvEmployee.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(484, 217)
        Me.dgvEmployee.TabIndex = 1
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
        Me.ColPrimaryCode.Width = 81
        '
        'ChkFormula
        '
        Me.ChkFormula.AutoSize = True
        Me.ChkFormula.Location = New System.Drawing.Point(4, 3)
        Me.ChkFormula.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkFormula.Name = "ChkFormula"
        Me.ChkFormula.Size = New System.Drawing.Size(168, 21)
        Me.ChkFormula.TabIndex = 0
        Me.ChkFormula.Text = "Select Salary Formula"
        Me.ChkFormula.UseVisualStyleBackColor = True
        '
        'Lbl_SF_Unchecked
        '
        Me.Lbl_SF_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Unchecked.AutoSize = True
        Me.Lbl_SF_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Unchecked.Location = New System.Drawing.Point(400, 25)
        Me.Lbl_SF_Unchecked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_SF_Unchecked.Name = "Lbl_SF_Unchecked"
        Me.Lbl_SF_Unchecked.Size = New System.Drawing.Size(82, 17)
        Me.Lbl_SF_Unchecked.TabIndex = 2
        Me.Lbl_SF_Unchecked.TabStop = True
        Me.Lbl_SF_Unchecked.Text = "Uncheck All"
        '
        'Lbl_SF_Checked
        '
        Me.Lbl_SF_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_SF_Checked.AutoSize = True
        Me.Lbl_SF_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_SF_Checked.Location = New System.Drawing.Point(322, 25)
        Me.Lbl_SF_Checked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_SF_Checked.Name = "Lbl_SF_Checked"
        Me.Lbl_SF_Checked.Size = New System.Drawing.Size(66, 17)
        Me.Lbl_SF_Checked.TabIndex = 1
        Me.Lbl_SF_Checked.TabStop = True
        Me.Lbl_SF_Checked.Text = "Check All"
        '
        'ClbSalaryFormula
        '
        Me.ClbSalaryFormula.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClbSalaryFormula.CheckOnClick = True
        Me.ClbSalaryFormula.FormattingEnabled = True
        Me.ClbSalaryFormula.Location = New System.Drawing.Point(4, 35)
        Me.ClbSalaryFormula.Margin = New System.Windows.Forms.Padding(4)
        Me.ClbSalaryFormula.MultiColumn = True
        Me.ClbSalaryFormula.Name = "ClbSalaryFormula"
        Me.ClbSalaryFormula.Size = New System.Drawing.Size(484, 225)
        Me.ClbSalaryFormula.TabIndex = 3
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.IsSplitterFixed = True
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.ChkDepartemen)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Lbl_Departemen_Unchecked)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Lbl_Departemen_Checked)
        Me.SplitContainer4.Panel1.Controls.Add(Me.ClbDepartemen)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.ChkPosition)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Lbl_Position_Unchecked)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Lbl_Position_Checked)
        Me.SplitContainer4.Panel2.Controls.Add(Me.ClbPositon)
        Me.SplitContainer4.Size = New System.Drawing.Size(496, 633)
        Me.SplitContainer4.SplitterDistance = 313
        Me.SplitContainer4.TabIndex = 0
        Me.SplitContainer4.TabStop = False
        '
        'ChkDepartemen
        '
        Me.ChkDepartemen.AutoSize = True
        Me.ChkDepartemen.Location = New System.Drawing.Point(7, 4)
        Me.ChkDepartemen.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkDepartemen.Name = "ChkDepartemen"
        Me.ChkDepartemen.Size = New System.Drawing.Size(147, 21)
        Me.ChkDepartemen.TabIndex = 0
        Me.ChkDepartemen.Text = "Select Department"
        Me.ChkDepartemen.UseVisualStyleBackColor = True
        '
        'Lbl_Departemen_Unchecked
        '
        Me.Lbl_Departemen_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Unchecked.AutoSize = True
        Me.Lbl_Departemen_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Unchecked.Location = New System.Drawing.Point(401, 16)
        Me.Lbl_Departemen_Unchecked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Departemen_Unchecked.Name = "Lbl_Departemen_Unchecked"
        Me.Lbl_Departemen_Unchecked.Size = New System.Drawing.Size(82, 17)
        Me.Lbl_Departemen_Unchecked.TabIndex = 2
        Me.Lbl_Departemen_Unchecked.TabStop = True
        Me.Lbl_Departemen_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Departemen_Checked
        '
        Me.Lbl_Departemen_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Departemen_Checked.AutoSize = True
        Me.Lbl_Departemen_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Departemen_Checked.Location = New System.Drawing.Point(323, 16)
        Me.Lbl_Departemen_Checked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Departemen_Checked.Name = "Lbl_Departemen_Checked"
        Me.Lbl_Departemen_Checked.Size = New System.Drawing.Size(66, 17)
        Me.Lbl_Departemen_Checked.TabIndex = 1
        Me.Lbl_Departemen_Checked.TabStop = True
        Me.Lbl_Departemen_Checked.Text = "Check All"
        '
        'ClbDepartemen
        '
        Me.ClbDepartemen.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClbDepartemen.CheckOnClick = True
        Me.ClbDepartemen.FormattingEnabled = True
        Me.ClbDepartemen.Location = New System.Drawing.Point(4, 25)
        Me.ClbDepartemen.Margin = New System.Windows.Forms.Padding(4)
        Me.ClbDepartemen.MultiColumn = True
        Me.ClbDepartemen.Name = "ClbDepartemen"
        Me.ClbDepartemen.Size = New System.Drawing.Size(488, 276)
        Me.ClbDepartemen.TabIndex = 3
        '
        'ChkPosition
        '
        Me.ChkPosition.AutoSize = True
        Me.ChkPosition.Location = New System.Drawing.Point(7, 8)
        Me.ChkPosition.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkPosition.Name = "ChkPosition"
        Me.ChkPosition.Size = New System.Drawing.Size(150, 21)
        Me.ChkPosition.TabIndex = 0
        Me.ChkPosition.Text = "Select Job Position"
        Me.ChkPosition.UseVisualStyleBackColor = True
        '
        'Lbl_Position_Unchecked
        '
        Me.Lbl_Position_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Unchecked.AutoSize = True
        Me.Lbl_Position_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Unchecked.Location = New System.Drawing.Point(397, 28)
        Me.Lbl_Position_Unchecked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Position_Unchecked.Name = "Lbl_Position_Unchecked"
        Me.Lbl_Position_Unchecked.Size = New System.Drawing.Size(82, 17)
        Me.Lbl_Position_Unchecked.TabIndex = 2
        Me.Lbl_Position_Unchecked.TabStop = True
        Me.Lbl_Position_Unchecked.Text = "Uncheck All"
        '
        'Lbl_Position_Checked
        '
        Me.Lbl_Position_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Position_Checked.AutoSize = True
        Me.Lbl_Position_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Position_Checked.Location = New System.Drawing.Point(319, 28)
        Me.Lbl_Position_Checked.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Position_Checked.Name = "Lbl_Position_Checked"
        Me.Lbl_Position_Checked.Size = New System.Drawing.Size(66, 17)
        Me.Lbl_Position_Checked.TabIndex = 1
        Me.Lbl_Position_Checked.TabStop = True
        Me.Lbl_Position_Checked.Text = "Check All"
        '
        'ClbPositon
        '
        Me.ClbPositon.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClbPositon.CheckOnClick = True
        Me.ClbPositon.FormattingEnabled = True
        Me.ClbPositon.Location = New System.Drawing.Point(7, 37)
        Me.ClbPositon.Margin = New System.Windows.Forms.Padding(4)
        Me.ClbPositon.Name = "ClbPositon"
        Me.ClbPositon.Size = New System.Drawing.Size(485, 276)
        Me.ClbPositon.TabIndex = 3
        '
        'BackgroundWorker1
        '
        '
        'FrmLaporanIncomeTaxStaff
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(992, 633)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimumSize = New System.Drawing.Size(1000, 670)
        Me.Name = "FrmLaporanIncomeTaxStaff"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Income Tax Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DtpStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents ChkEmployee As System.Windows.Forms.CheckBox
    Friend WithEvents ClbSalaryFormula As System.Windows.Forms.CheckedListBox
    Friend WithEvents Lbl_SF_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_SF_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents ClbDepartemen As System.Windows.Forms.CheckedListBox
    Friend WithEvents Lbl_Departemen_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Departemen_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkFormula As System.Windows.Forms.CheckBox
    Friend WithEvents ChkDepartemen As System.Windows.Forms.CheckBox
    Friend WithEvents ClbPositon As System.Windows.Forms.CheckedListBox
    Friend WithEvents Lbl_Position_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Position_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkPosition As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents ColCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
End Class
