<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanKontrakKaryawan
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
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.dgvContractType = New System.Windows.Forms.DataGridView
        Me.ColContractTipeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContractTipeDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColContractTipePrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboPeriod = New System.Windows.Forms.ComboBox
        Me.BgwEmployee = New System.ComponentModel.BackgroundWorker
        Me.BgwDepartment = New System.ComponentModel.BackgroundWorker
        Me.BgwJobPosition = New System.ComponentModel.BackgroundWorker
        Me.BgwContractType = New System.ComponentModel.BackgroundWorker
        Me.CboFilterEmployee = New System.Windows.Forms.ComboBox
        Me.GroupBox2.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage8.SuspendLayout()
        CType(Me.dgvContractType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.crvLaporanKontrakKaryawan.Location = New System.Drawing.Point(19, 163)
        Me.crvLaporanKontrakKaryawan.Name = "crvLaporanKontrakKaryawan"
        Me.crvLaporanKontrakKaryawan.SelectionFormula = ""
        Me.crvLaporanKontrakKaryawan.ShowCloseButton = False
        Me.crvLaporanKontrakKaryawan.ShowGroupTreeButton = False
        Me.crvLaporanKontrakKaryawan.ShowRefreshButton = False
        Me.crvLaporanKontrakKaryawan.Size = New System.Drawing.Size(725, 339)
        Me.crvLaporanKontrakKaryawan.TabIndex = 12
        Me.crvLaporanKontrakKaryawan.TabStop = False
        Me.crvLaporanKontrakKaryawan.ViewTimeSelectionFormula = ""
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(76, 42)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(119, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 45)
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
        Me.dtpTo.Size = New System.Drawing.Size(119, 20)
        Me.dtpTo.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "To"
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(76, 119)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 21)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(76, 118)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 20)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.CboFilterEmployee)
        Me.GroupBox2.Controls.Add(Me.TabControl2)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.cboPeriod)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.btnShow)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 10)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(726, 145)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage8)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Location = New System.Drawing.Point(200, 17)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(522, 124)
        Me.TabControl2.TabIndex = 26
        Me.TabControl2.TabStop = False
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.dgvContractType)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage8.Size = New System.Drawing.Size(514, 98)
        Me.TabPage8.TabIndex = 3
        Me.TabPage8.Text = "Contract Type"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'dgvContractType
        '
        Me.dgvContractType.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvContractType.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColContractTipeCode, Me.ColContractTipeDesc, Me.ColContractTipePrimary})
        Me.dgvContractType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvContractType.Location = New System.Drawing.Point(3, 3)
        Me.dgvContractType.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvContractType.Name = "dgvContractType"
        Me.dgvContractType.RowTemplate.Height = 24
        Me.dgvContractType.Size = New System.Drawing.Size(508, 92)
        Me.dgvContractType.TabIndex = 2
        Me.dgvContractType.TabStop = False
        '
        'ColContractTipeCode
        '
        Me.ColContractTipeCode.HeaderText = "Code"
        Me.ColContractTipeCode.Name = "ColContractTipeCode"
        '
        'ColContractTipeDesc
        '
        Me.ColContractTipeDesc.HeaderText = "Description"
        Me.ColContractTipeDesc.Name = "ColContractTipeDesc"
        '
        'ColContractTipePrimary
        '
        Me.ColContractTipePrimary.HeaderText = "Primary"
        Me.ColContractTipePrimary.Name = "ColContractTipePrimary"
        Me.ColContractTipePrimary.Visible = False
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.dgvEmployee)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(514, 98)
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
        Me.dgvEmployee.Size = New System.Drawing.Size(510, 94)
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
        Me.TabPage6.Size = New System.Drawing.Size(514, 98)
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
        Me.dgvDepartment.Size = New System.Drawing.Size(510, 94)
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
        Me.TabPage7.Size = New System.Drawing.Size(514, 98)
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
        Me.dgvJobPosition.Size = New System.Drawing.Size(510, 94)
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
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 20)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 13)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Period Type"
        '
        'cboPeriod
        '
        Me.cboPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPeriod.FormattingEnabled = True
        Me.cboPeriod.Items.AddRange(New Object() {"Contract Date", "Contract End Date", "Contract Start Date", "All Data"})
        Me.cboPeriod.Location = New System.Drawing.Point(76, 17)
        Me.cboPeriod.Margin = New System.Windows.Forms.Padding(2)
        Me.cboPeriod.Name = "cboPeriod"
        Me.cboPeriod.Size = New System.Drawing.Size(119, 21)
        Me.cboPeriod.TabIndex = 0
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
        'BgwContractType
        '
        '
        'CboFilterEmployee
        '
        Me.CboFilterEmployee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFilterEmployee.FormattingEnabled = True
        Me.CboFilterEmployee.Items.AddRange(New Object() {"All", "Active", "Not Active"})
        Me.CboFilterEmployee.Location = New System.Drawing.Point(76, 92)
        Me.CboFilterEmployee.Margin = New System.Windows.Forms.Padding(2)
        Me.CboFilterEmployee.Name = "CboFilterEmployee"
        Me.CboFilterEmployee.Size = New System.Drawing.Size(119, 21)
        Me.CboFilterEmployee.TabIndex = 100
        '
        'FrmLaporanKontrakKaryawan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(744, 514)
        Me.Controls.Add(Me.crvLaporanKontrakKaryawan)
        Me.Controls.Add(Me.GroupBox2)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmLaporanKontrakKaryawan"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Employee Contract Report"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage8.ResumeLayout(False)
        CType(Me.dgvContractType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents cboPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents dgvContractType As System.Windows.Forms.DataGridView
    Friend WithEvents ColContractTipeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContractTipeDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColContractTipePrimary As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents BgwContractType As System.ComponentModel.BackgroundWorker
    Friend WithEvents CboFilterEmployee As System.Windows.Forms.ComboBox
End Class
