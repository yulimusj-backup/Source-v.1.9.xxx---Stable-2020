<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanOvertimeAnalysis2
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
        Me.crvLaporanAnalysisOvertime = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnShow = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboFormat = New System.Windows.Forms.ComboBox
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.LblDataGrouping = New System.Windows.Forms.Label
        Me.cboGrouping = New System.Windows.Forms.ComboBox
        Me.cboData1 = New System.Windows.Forms.ComboBox
        Me.LblData1 = New System.Windows.Forms.Label
        Me.LblData2 = New System.Windows.Forms.Label
        Me.LblData3 = New System.Windows.Forms.Label
        Me.cboData3 = New System.Windows.Forms.ComboBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ChkShowSubTotal = New System.Windows.Forms.CheckBox
        Me.ChkShowOther = New System.Windows.Forms.CheckBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.CboFilterDate = New System.Windows.Forms.ComboBox
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
        Me.LblChart_DataType = New System.Windows.Forms.Label
        Me.TxtChart_DataSelection = New System.Windows.Forms.TextBox
        Me.LblChart_DataSelection = New System.Windows.Forms.Label
        Me.cboChart_DataSelection = New System.Windows.Forms.ComboBox
        Me.LblChartType = New System.Windows.Forms.Label
        Me.cboChartType = New System.Windows.Forms.ComboBox
        Me.cboData2 = New System.Windows.Forms.ComboBox
        Me.cboChart_DataType = New System.Windows.Forms.ComboBox
        Me.BgwDepartment = New System.ComponentModel.BackgroundWorker
        Me.BgwJobPosition = New System.ComponentModel.BackgroundWorker
        Me.BgwEmployee = New System.ComponentModel.BackgroundWorker
        Me.GroupBox2.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'crvLaporanAnalysisOvertime
        '
        Me.crvLaporanAnalysisOvertime.ActiveViewIndex = -1
        Me.crvLaporanAnalysisOvertime.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crvLaporanAnalysisOvertime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvLaporanAnalysisOvertime.DisplayGroupTree = False
        Me.crvLaporanAnalysisOvertime.EnableDrillDown = False
        Me.crvLaporanAnalysisOvertime.Location = New System.Drawing.Point(5, 146)
        Me.crvLaporanAnalysisOvertime.Name = "crvLaporanAnalysisOvertime"
        Me.crvLaporanAnalysisOvertime.SelectionFormula = ""
        Me.crvLaporanAnalysisOvertime.ShowCloseButton = False
        Me.crvLaporanAnalysisOvertime.ShowGroupTreeButton = False
        Me.crvLaporanAnalysisOvertime.ShowRefreshButton = False
        Me.crvLaporanAnalysisOvertime.Size = New System.Drawing.Size(813, 363)
        Me.crvLaporanAnalysisOvertime.TabIndex = 10
        Me.crvLaporanAnalysisOvertime.TabStop = False
        Me.crvLaporanAnalysisOvertime.ViewTimeSelectionFormula = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 94)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Report Format"
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(364, 112)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(75, 20)
        Me.btnShow.TabIndex = 11
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(365, 112)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 20)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(90, 43)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.ShowUpDown = True
        Me.dtpFrom.Size = New System.Drawing.Size(135, 20)
        Me.dtpFrom.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 70)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "To"
        '
        'cboFormat
        '
        Me.cboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFormat.FormattingEnabled = True
        Me.cboFormat.Items.AddRange(New Object() {"Data", "Chart"})
        Me.cboFormat.Location = New System.Drawing.Point(90, 91)
        Me.cboFormat.Margin = New System.Windows.Forms.Padding(2)
        Me.cboFormat.Name = "cboFormat"
        Me.cboFormat.Size = New System.Drawing.Size(135, 21)
        Me.cboFormat.TabIndex = 0
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(90, 67)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.ShowUpDown = True
        Me.dtpTo.Size = New System.Drawing.Size(135, 20)
        Me.dtpTo.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 47)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "From"
        '
        'LblDataGrouping
        '
        Me.LblDataGrouping.AutoSize = True
        Me.LblDataGrouping.Location = New System.Drawing.Point(7, 118)
        Me.LblDataGrouping.Name = "LblDataGrouping"
        Me.LblDataGrouping.Size = New System.Drawing.Size(76, 13)
        Me.LblDataGrouping.TabIndex = 16
        Me.LblDataGrouping.Text = "Data Grouping"
        '
        'cboGrouping
        '
        Me.cboGrouping.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGrouping.FormattingEnabled = True
        Me.cboGrouping.Items.AddRange(New Object() {"No Grouping", "Single Grouping", "Double Grouping"})
        Me.cboGrouping.Location = New System.Drawing.Point(90, 116)
        Me.cboGrouping.Margin = New System.Windows.Forms.Padding(2)
        Me.cboGrouping.Name = "cboGrouping"
        Me.cboGrouping.Size = New System.Drawing.Size(135, 21)
        Me.cboGrouping.TabIndex = 15
        '
        'cboData1
        '
        Me.cboData1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboData1.FormattingEnabled = True
        Me.cboData1.Items.AddRange(New Object() {"Department", "Employee", "Job Position", "Monthly Period"})
        Me.cboData1.Location = New System.Drawing.Point(319, 17)
        Me.cboData1.Margin = New System.Windows.Forms.Padding(2)
        Me.cboData1.Name = "cboData1"
        Me.cboData1.Size = New System.Drawing.Size(120, 21)
        Me.cboData1.TabIndex = 17
        '
        'LblData1
        '
        Me.LblData1.AutoSize = True
        Me.LblData1.Location = New System.Drawing.Point(235, 20)
        Me.LblData1.Name = "LblData1"
        Me.LblData1.Size = New System.Drawing.Size(30, 13)
        Me.LblData1.TabIndex = 18
        Me.LblData1.Text = "Data"
        '
        'LblData2
        '
        Me.LblData2.AutoSize = True
        Me.LblData2.Location = New System.Drawing.Point(235, 46)
        Me.LblData2.Name = "LblData2"
        Me.LblData2.Size = New System.Drawing.Size(30, 13)
        Me.LblData2.TabIndex = 20
        Me.LblData2.Text = "Data"
        '
        'LblData3
        '
        Me.LblData3.AutoSize = True
        Me.LblData3.Location = New System.Drawing.Point(235, 69)
        Me.LblData3.Name = "LblData3"
        Me.LblData3.Size = New System.Drawing.Size(30, 13)
        Me.LblData3.TabIndex = 22
        Me.LblData3.Text = "Data"
        '
        'cboData3
        '
        Me.cboData3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboData3.FormattingEnabled = True
        Me.cboData3.Items.AddRange(New Object() {"Department", "Employee", "Job Position", "Monthly Period", "n/a"})
        Me.cboData3.Location = New System.Drawing.Point(319, 65)
        Me.cboData3.Margin = New System.Windows.Forms.Padding(2)
        Me.cboData3.Name = "cboData3"
        Me.cboData3.Size = New System.Drawing.Size(119, 21)
        Me.cboData3.TabIndex = 21
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.ChkShowSubTotal)
        Me.GroupBox2.Controls.Add(Me.btnShow)
        Me.GroupBox2.Controls.Add(Me.ChkShowOther)
        Me.GroupBox2.Controls.Add(Me.btnClose)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.CboFilterDate)
        Me.GroupBox2.Controls.Add(Me.TabControl2)
        Me.GroupBox2.Controls.Add(Me.LblChart_DataType)
        Me.GroupBox2.Controls.Add(Me.TxtChart_DataSelection)
        Me.GroupBox2.Controls.Add(Me.LblChart_DataSelection)
        Me.GroupBox2.Controls.Add(Me.cboChart_DataSelection)
        Me.GroupBox2.Controls.Add(Me.LblChartType)
        Me.GroupBox2.Controls.Add(Me.cboChartType)
        Me.GroupBox2.Controls.Add(Me.LblData3)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.LblData2)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.cboData2)
        Me.GroupBox2.Controls.Add(Me.cboFormat)
        Me.GroupBox2.Controls.Add(Me.LblData1)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.cboData1)
        Me.GroupBox2.Controls.Add(Me.LblDataGrouping)
        Me.GroupBox2.Controls.Add(Me.cboGrouping)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.cboChart_DataType)
        Me.GroupBox2.Controls.Add(Me.cboData3)
        Me.GroupBox2.Location = New System.Drawing.Point(5, -1)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(813, 142)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        '
        'ChkShowSubTotal
        '
        Me.ChkShowSubTotal.AutoSize = True
        Me.ChkShowSubTotal.Location = New System.Drawing.Point(238, 95)
        Me.ChkShowSubTotal.Name = "ChkShowSubTotal"
        Me.ChkShowSubTotal.Size = New System.Drawing.Size(102, 17)
        Me.ChkShowSubTotal.TabIndex = 35
        Me.ChkShowSubTotal.Text = "Show Sub Total"
        Me.ChkShowSubTotal.UseVisualStyleBackColor = True
        '
        'ChkShowOther
        '
        Me.ChkShowOther.AutoSize = True
        Me.ChkShowOther.Location = New System.Drawing.Point(238, 118)
        Me.ChkShowOther.Name = "ChkShowOther"
        Me.ChkShowOther.Size = New System.Drawing.Size(108, 17)
        Me.ChkShowOther.TabIndex = 34
        Me.ChkShowOther.Text = "Show Other Data"
        Me.ChkShowOther.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 13)
        Me.Label4.TabIndex = 33
        Me.Label4.Text = "Base On"
        '
        'CboFilterDate
        '
        Me.CboFilterDate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFilterDate.FormattingEnabled = True
        Me.CboFilterDate.Items.AddRange(New Object() {"Actual Date", "Salary Calculation Date"})
        Me.CboFilterDate.Location = New System.Drawing.Point(90, 17)
        Me.CboFilterDate.Margin = New System.Windows.Forms.Padding(2)
        Me.CboFilterDate.Name = "CboFilterDate"
        Me.CboFilterDate.Size = New System.Drawing.Size(135, 21)
        Me.CboFilterDate.TabIndex = 32
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Location = New System.Drawing.Point(445, 17)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(364, 121)
        Me.TabControl2.TabIndex = 31
        Me.TabControl2.TabStop = False
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.dgvEmployee)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(356, 95)
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
        Me.dgvEmployee.Size = New System.Drawing.Size(352, 91)
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
        Me.TabPage6.Size = New System.Drawing.Size(356, 95)
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
        Me.dgvDepartment.Size = New System.Drawing.Size(352, 91)
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
        Me.TabPage7.Size = New System.Drawing.Size(356, 95)
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
        Me.dgvJobPosition.Size = New System.Drawing.Size(352, 91)
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
        'LblChart_DataType
        '
        Me.LblChart_DataType.AutoSize = True
        Me.LblChart_DataType.Location = New System.Drawing.Point(235, 68)
        Me.LblChart_DataType.Name = "LblChart_DataType"
        Me.LblChart_DataType.Size = New System.Drawing.Size(57, 13)
        Me.LblChart_DataType.TabIndex = 30
        Me.LblChart_DataType.Text = "Data Type"
        '
        'TxtChart_DataSelection
        '
        Me.TxtChart_DataSelection.Location = New System.Drawing.Point(398, 41)
        Me.TxtChart_DataSelection.Margin = New System.Windows.Forms.Padding(2)
        Me.TxtChart_DataSelection.Name = "TxtChart_DataSelection"
        Me.TxtChart_DataSelection.Size = New System.Drawing.Size(39, 20)
        Me.TxtChart_DataSelection.TabIndex = 28
        Me.TxtChart_DataSelection.Text = "0"
        Me.TxtChart_DataSelection.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblChart_DataSelection
        '
        Me.LblChart_DataSelection.AutoSize = True
        Me.LblChart_DataSelection.Location = New System.Drawing.Point(235, 44)
        Me.LblChart_DataSelection.Name = "LblChart_DataSelection"
        Me.LblChart_DataSelection.Size = New System.Drawing.Size(56, 13)
        Me.LblChart_DataSelection.TabIndex = 27
        Me.LblChart_DataSelection.Text = "Data View"
        '
        'cboChart_DataSelection
        '
        Me.cboChart_DataSelection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChart_DataSelection.FormattingEnabled = True
        Me.cboChart_DataSelection.Items.AddRange(New Object() {"All Data", "Top"})
        Me.cboChart_DataSelection.Location = New System.Drawing.Point(319, 41)
        Me.cboChart_DataSelection.Margin = New System.Windows.Forms.Padding(2)
        Me.cboChart_DataSelection.Name = "cboChart_DataSelection"
        Me.cboChart_DataSelection.Size = New System.Drawing.Size(76, 21)
        Me.cboChart_DataSelection.TabIndex = 26
        '
        'LblChartType
        '
        Me.LblChartType.AutoSize = True
        Me.LblChartType.Location = New System.Drawing.Point(7, 118)
        Me.LblChartType.Name = "LblChartType"
        Me.LblChartType.Size = New System.Drawing.Size(59, 13)
        Me.LblChartType.TabIndex = 25
        Me.LblChartType.Text = "Chart Type"
        '
        'cboChartType
        '
        Me.cboChartType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChartType.FormattingEnabled = True
        Me.cboChartType.Items.AddRange(New Object() {"Line", "Pie"})
        Me.cboChartType.Location = New System.Drawing.Point(90, 116)
        Me.cboChartType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboChartType.Name = "cboChartType"
        Me.cboChartType.Size = New System.Drawing.Size(135, 21)
        Me.cboChartType.TabIndex = 24
        '
        'cboData2
        '
        Me.cboData2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboData2.FormattingEnabled = True
        Me.cboData2.Items.AddRange(New Object() {"Department", "Employee", "Job Position", "Monthly Period", "n/a"})
        Me.cboData2.Location = New System.Drawing.Point(319, 41)
        Me.cboData2.Margin = New System.Windows.Forms.Padding(2)
        Me.cboData2.Name = "cboData2"
        Me.cboData2.Size = New System.Drawing.Size(120, 21)
        Me.cboData2.TabIndex = 19
        '
        'cboChart_DataType
        '
        Me.cboChart_DataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboChart_DataType.FormattingEnabled = True
        Me.cboChart_DataType.Items.AddRange(New Object() {"OT Amount", "OT Duration"})
        Me.cboChart_DataType.Location = New System.Drawing.Point(319, 65)
        Me.cboChart_DataType.Margin = New System.Windows.Forms.Padding(2)
        Me.cboChart_DataType.Name = "cboChart_DataType"
        Me.cboChart_DataType.Size = New System.Drawing.Size(120, 21)
        Me.cboChart_DataType.TabIndex = 29
        '
        'BgwDepartment
        '
        '
        'BgwJobPosition
        '
        '
        'BgwEmployee
        '
        '
        'FrmLaporanOvertimeAnalysis2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(823, 514)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.crvLaporanAnalysisOvertime)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmLaporanOvertimeAnalysis2"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Analisis"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crvLaporanAnalysisOvertime As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents LblData3 As System.Windows.Forms.Label
    Friend WithEvents cboData3 As System.Windows.Forms.ComboBox
    Friend WithEvents LblData2 As System.Windows.Forms.Label
    Friend WithEvents LblData1 As System.Windows.Forms.Label
    Friend WithEvents cboData1 As System.Windows.Forms.ComboBox
    Friend WithEvents LblDataGrouping As System.Windows.Forms.Label
    Friend WithEvents cboGrouping As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents LblChartType As System.Windows.Forms.Label
    Friend WithEvents cboChartType As System.Windows.Forms.ComboBox
    Friend WithEvents LblChart_DataSelection As System.Windows.Forms.Label
    Friend WithEvents cboChart_DataSelection As System.Windows.Forms.ComboBox
    Friend WithEvents TxtChart_DataSelection As System.Windows.Forms.TextBox
    Friend WithEvents LblChart_DataType As System.Windows.Forms.Label
    Friend WithEvents cboChart_DataType As System.Windows.Forms.ComboBox
    Friend WithEvents cboData2 As System.Windows.Forms.ComboBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents dgvEmployee As System.Windows.Forms.DataGridView
    Friend WithEvents ColEmployee_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployee_primary As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents BgwDepartment As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwJobPosition As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwEmployee As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CboFilterDate As System.Windows.Forms.ComboBox
    Friend WithEvents ChkShowSubTotal As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowOther As System.Windows.Forms.CheckBox
End Class
