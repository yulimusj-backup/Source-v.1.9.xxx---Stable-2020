<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanOtherPayment
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
        Me.crvLaporanKontrakKaryawan = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnShow = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.CboGrouping = New System.Windows.Forms.ComboBox
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.dgvEmployee = New System.Windows.Forms.DataGridView
        Me.ColEmployee_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployee_primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.dgvDepartment = New System.Windows.Forms.DataGridView
        Me.ColDeptCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeptPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.dgvJobPosition = New System.Windows.Forms.DataGridView
        Me.ColJobPosisitonCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPosisitonDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJobPosisitonPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvSalaryFormula = New System.Windows.Forms.DataGridView
        Me.ColSC_ComponentName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSC_SalaryFormula = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSC_PrimarySalary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSC_IdentifierComponent = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvDocNo = New System.Windows.Forms.DataGridView
        Me.ColDocNo_Document = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocNo_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label3 = New System.Windows.Forms.Label
        Me.CboReport = New System.Windows.Forms.ComboBox
        Me.BgwEmployee = New System.ComponentModel.BackgroundWorker
        Me.BgwDepartment = New System.ComponentModel.BackgroundWorker
        Me.BgwJobPosition = New System.ComponentModel.BackgroundWorker
        Me.BgwSalaryComponent = New System.ComponentModel.BackgroundWorker
        Me.BgwDocNo = New System.ComponentModel.BackgroundWorker
        Me.GroupBox2.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvSalaryFormula, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvDocNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'crvLaporanKontrakKaryawan
        '
        Me.crvLaporanKontrakKaryawan.ActiveViewIndex = -1
        Me.crvLaporanKontrakKaryawan.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvLaporanKontrakKaryawan.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvLaporanKontrakKaryawan.DisplayGroupTree = False
        Me.crvLaporanKontrakKaryawan.EnableDrillDown = False
        Me.crvLaporanKontrakKaryawan.Location = New System.Drawing.Point(3, 152)
        Me.crvLaporanKontrakKaryawan.Name = "crvLaporanKontrakKaryawan"
        Me.crvLaporanKontrakKaryawan.SelectionFormula = ""
        Me.crvLaporanKontrakKaryawan.ShowCloseButton = False
        Me.crvLaporanKontrakKaryawan.ShowGroupTreeButton = False
        Me.crvLaporanKontrakKaryawan.ShowRefreshButton = False
        Me.crvLaporanKontrakKaryawan.Size = New System.Drawing.Size(737, 361)
        Me.crvLaporanKontrakKaryawan.TabIndex = 12
        Me.crvLaporanKontrakKaryawan.TabStop = False
        Me.crvLaporanKontrakKaryawan.ViewTimeSelectionFormula = ""
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(76, 41)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(176, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "From"
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(76, 67)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.ShowUpDown = True
        Me.dtpTo.Size = New System.Drawing.Size(176, 20)
        Me.dtpTo.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 70)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "To"
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(177, 117)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 20)
        Me.btnShow.TabIndex = 4
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(177, 117)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 20)
        Me.btnClose.TabIndex = 12
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.CboGrouping)
        Me.GroupBox2.Controls.Add(Me.TabControl2)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.btnShow)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.CboReport)
        Me.GroupBox2.Location = New System.Drawing.Point(3, 1)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(741, 146)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 95)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 13)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "Grouping"
        '
        'CboGrouping
        '
        Me.CboGrouping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboGrouping.FormattingEnabled = True
        Me.CboGrouping.Items.AddRange(New Object() {"No Grouping", "Group By Employee", "Group By Component Name", "Group By Document No"})
        Me.CboGrouping.Location = New System.Drawing.Point(76, 92)
        Me.CboGrouping.Margin = New System.Windows.Forms.Padding(2)
        Me.CboGrouping.Name = "CboGrouping"
        Me.CboGrouping.Size = New System.Drawing.Size(176, 21)
        Me.CboGrouping.TabIndex = 3
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Controls.Add(Me.TabPage1)
        Me.TabControl2.Controls.Add(Me.TabPage2)
        Me.TabControl2.Location = New System.Drawing.Point(257, 17)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(480, 124)
        Me.TabControl2.TabIndex = 26
        Me.TabControl2.TabStop = False
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.dgvEmployee)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(472, 98)
        Me.TabPage5.TabIndex = 0
        Me.TabPage5.Text = "Employee"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'dgvEmployee
        '
        Me.dgvEmployee.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvEmployee.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColEmployee_Code, Me.ColEmployee_Name, Me.ColEmployee_primary})
        Me.dgvEmployee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmployee.Location = New System.Drawing.Point(2, 2)
        Me.dgvEmployee.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvEmployee.Name = "dgvEmployee"
        Me.dgvEmployee.RowTemplate.Height = 24
        Me.dgvEmployee.Size = New System.Drawing.Size(468, 94)
        Me.dgvEmployee.TabIndex = 0
        Me.dgvEmployee.TabStop = False
        '
        'ColEmployee_Code
        '
        Me.ColEmployee_Code.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColEmployee_Code.HeaderText = "Code"
        Me.ColEmployee_Code.MinimumWidth = 50
        Me.ColEmployee_Code.Name = "ColEmployee_Code"
        Me.ColEmployee_Code.Width = 57
        '
        'ColEmployee_Name
        '
        Me.ColEmployee_Name.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColEmployee_Name.HeaderText = "Name"
        Me.ColEmployee_Name.MinimumWidth = 100
        Me.ColEmployee_Name.Name = "ColEmployee_Name"
        '
        'ColEmployee_primary
        '
        Me.ColEmployee_primary.HeaderText = "Primary"
        Me.ColEmployee_primary.Name = "ColEmployee_primary"
        Me.ColEmployee_primary.Visible = False
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.dgvDepartment)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage6.Size = New System.Drawing.Size(472, 98)
        Me.TabPage6.TabIndex = 1
        Me.TabPage6.Text = "Department"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'dgvDepartment
        '
        Me.dgvDepartment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDepartment.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDeptCode, Me.ColDeptDesc, Me.ColDeptPrimary})
        Me.dgvDepartment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDepartment.Location = New System.Drawing.Point(2, 2)
        Me.dgvDepartment.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvDepartment.Name = "dgvDepartment"
        Me.dgvDepartment.RowTemplate.Height = 24
        Me.dgvDepartment.Size = New System.Drawing.Size(468, 94)
        Me.dgvDepartment.TabIndex = 1
        Me.dgvDepartment.TabStop = False
        '
        'ColDeptCode
        '
        Me.ColDeptCode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColDeptCode.HeaderText = "Code"
        Me.ColDeptCode.MinimumWidth = 50
        Me.ColDeptCode.Name = "ColDeptCode"
        Me.ColDeptCode.Width = 57
        '
        'ColDeptDesc
        '
        Me.ColDeptDesc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColDeptDesc.HeaderText = "Description"
        Me.ColDeptDesc.MinimumWidth = 50
        Me.ColDeptDesc.Name = "ColDeptDesc"
        Me.ColDeptDesc.Width = 85
        '
        'ColDeptPrimary
        '
        Me.ColDeptPrimary.HeaderText = "Primary"
        Me.ColDeptPrimary.Name = "ColDeptPrimary"
        Me.ColDeptPrimary.Visible = False
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.dgvJobPosition)
        Me.TabPage7.Location = New System.Drawing.Point(4, 22)
        Me.TabPage7.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage7.Size = New System.Drawing.Size(472, 98)
        Me.TabPage7.TabIndex = 2
        Me.TabPage7.Text = "Job Position"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'dgvJobPosition
        '
        Me.dgvJobPosition.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvJobPosition.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColJobPosisitonCode, Me.ColJobPosisitonDesc, Me.ColJobPosisitonPrimary})
        Me.dgvJobPosition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvJobPosition.Location = New System.Drawing.Point(2, 2)
        Me.dgvJobPosition.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvJobPosition.Name = "dgvJobPosition"
        Me.dgvJobPosition.RowTemplate.Height = 24
        Me.dgvJobPosition.Size = New System.Drawing.Size(468, 94)
        Me.dgvJobPosition.TabIndex = 2
        Me.dgvJobPosition.TabStop = False
        '
        'ColJobPosisitonCode
        '
        Me.ColJobPosisitonCode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColJobPosisitonCode.HeaderText = "Code"
        Me.ColJobPosisitonCode.MinimumWidth = 50
        Me.ColJobPosisitonCode.Name = "ColJobPosisitonCode"
        Me.ColJobPosisitonCode.Width = 57
        '
        'ColJobPosisitonDesc
        '
        Me.ColJobPosisitonDesc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColJobPosisitonDesc.HeaderText = "Description"
        Me.ColJobPosisitonDesc.MinimumWidth = 50
        Me.ColJobPosisitonDesc.Name = "ColJobPosisitonDesc"
        Me.ColJobPosisitonDesc.Width = 85
        '
        'ColJobPosisitonPrimary
        '
        Me.ColJobPosisitonPrimary.HeaderText = "Primary"
        Me.ColJobPosisitonPrimary.Name = "ColJobPosisitonPrimary"
        Me.ColJobPosisitonPrimary.Visible = False
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvSalaryFormula)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(472, 98)
        Me.TabPage1.TabIndex = 3
        Me.TabPage1.Text = "Salary Component"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvSalaryFormula
        '
        Me.DgvSalaryFormula.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvSalaryFormula.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvSalaryFormula.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSC_ComponentName, Me.ColSC_SalaryFormula, Me.ColSC_PrimarySalary, Me.ColSC_IdentifierComponent})
        Me.DgvSalaryFormula.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvSalaryFormula.Location = New System.Drawing.Point(3, 3)
        Me.DgvSalaryFormula.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvSalaryFormula.Name = "DgvSalaryFormula"
        Me.DgvSalaryFormula.RowTemplate.Height = 24
        Me.DgvSalaryFormula.Size = New System.Drawing.Size(466, 92)
        Me.DgvSalaryFormula.TabIndex = 3
        Me.DgvSalaryFormula.TabStop = False
        '
        'ColSC_ComponentName
        '
        Me.ColSC_ComponentName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColSC_ComponentName.HeaderText = "Component Name"
        Me.ColSC_ComponentName.MinimumWidth = 50
        Me.ColSC_ComponentName.Name = "ColSC_ComponentName"
        Me.ColSC_ComponentName.Width = 117
        '
        'ColSC_SalaryFormula
        '
        Me.ColSC_SalaryFormula.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColSC_SalaryFormula.HeaderText = "Salary Formula"
        Me.ColSC_SalaryFormula.MinimumWidth = 50
        Me.ColSC_SalaryFormula.Name = "ColSC_SalaryFormula"
        Me.ColSC_SalaryFormula.Width = 101
        '
        'ColSC_PrimarySalary
        '
        Me.ColSC_PrimarySalary.HeaderText = "Primary"
        Me.ColSC_PrimarySalary.Name = "ColSC_PrimarySalary"
        Me.ColSC_PrimarySalary.Visible = False
        '
        'ColSC_IdentifierComponent
        '
        Me.ColSC_IdentifierComponent.HeaderText = "Identifier Component"
        Me.ColSC_IdentifierComponent.Name = "ColSC_IdentifierComponent"
        Me.ColSC_IdentifierComponent.ReadOnly = True
        Me.ColSC_IdentifierComponent.Visible = False
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvDocNo)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(472, 98)
        Me.TabPage2.TabIndex = 4
        Me.TabPage2.Text = "Document No"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvDocNo
        '
        Me.DgvDocNo.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvDocNo.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvDocNo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvDocNo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDocNo_Document, Me.ColDocNo_Primary})
        Me.DgvDocNo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvDocNo.Location = New System.Drawing.Point(3, 3)
        Me.DgvDocNo.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvDocNo.Name = "DgvDocNo"
        Me.DgvDocNo.RowTemplate.Height = 24
        Me.DgvDocNo.Size = New System.Drawing.Size(466, 92)
        Me.DgvDocNo.TabIndex = 3
        Me.DgvDocNo.TabStop = False
        '
        'ColDocNo_Document
        '
        Me.ColDocNo_Document.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ColDocNo_Document.HeaderText = "Document No"
        Me.ColDocNo_Document.MinimumWidth = 50
        Me.ColDocNo_Document.Name = "ColDocNo_Document"
        Me.ColDocNo_Document.Width = 98
        '
        'ColDocNo_Primary
        '
        Me.ColDocNo_Primary.HeaderText = "Primary"
        Me.ColDocNo_Primary.Name = "ColDocNo_Primary"
        Me.ColDocNo_Primary.Visible = False
        Me.ColDocNo_Primary.Width = 66
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(39, 13)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Report"
        '
        'CboReport
        '
        Me.CboReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboReport.FormattingEnabled = True
        Me.CboReport.Items.AddRange(New Object() {"Other Payment", "Salary Override "})
        Me.CboReport.Location = New System.Drawing.Point(76, 17)
        Me.CboReport.Margin = New System.Windows.Forms.Padding(2)
        Me.CboReport.Name = "CboReport"
        Me.CboReport.Size = New System.Drawing.Size(176, 21)
        Me.CboReport.TabIndex = 0
        '
        'BgwEmployee
        '
        '
        'BgwDepartment
        '
        '
        'BgwJobPosition
        '
        '
        'BgwSalaryComponent
        '
        '
        'BgwDocNo
        '
        '
        'FrmLaporanOtherPayment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(744, 514)
        Me.Controls.Add(Me.crvLaporanKontrakKaryawan)
        Me.Controls.Add(Me.GroupBox2)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmLaporanOtherPayment"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Other Payment Report"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvSalaryFormula, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvDocNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crvLaporanKontrakKaryawan As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CboReport As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents dgvDepartment As System.Windows.Forms.DataGridView
    Friend WithEvents ColDeptCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeptPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents dgvJobPosition As System.Windows.Forms.DataGridView
    Friend WithEvents ColJobPosisitonCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPosisitonDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJobPosisitonPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwEmployee As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwDepartment As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwJobPosition As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwSalaryComponent As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CboGrouping As System.Windows.Forms.ComboBox
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DgvSalaryFormula As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents DgvDocNo As System.Windows.Forms.DataGridView
    Friend WithEvents ColDocNo_Document As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocNo_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSC_ComponentName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSC_SalaryFormula As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSC_PrimarySalary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSC_IdentifierComponent As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwDocNo As System.ComponentModel.BackgroundWorker
End Class
