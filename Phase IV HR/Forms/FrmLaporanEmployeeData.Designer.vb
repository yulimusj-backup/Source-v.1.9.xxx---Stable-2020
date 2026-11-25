<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanEmployeeData
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
        Me.Lbl_JB_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_JB_Checked = New System.Windows.Forms.LinkLabel
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkJobPosition = New System.Windows.Forms.CheckBox
        Me.clbJobPosition = New System.Windows.Forms.CheckedListBox
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.btnShow = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ChkDepartement = New System.Windows.Forms.CheckBox
        Me.Lbl_Dept_Unchecked = New System.Windows.Forms.LinkLabel
        Me.Lbl_Dept_Checked = New System.Windows.Forms.LinkLabel
        Me.ClbDepartement = New System.Windows.Forms.CheckedListBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.RboFormat1 = New System.Windows.Forms.RadioButton
        Me.RboFormat2 = New System.Windows.Forms.RadioButton
        Me.RboFamilyData = New System.Windows.Forms.RadioButton
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboPeriod = New System.Windows.Forms.ComboBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.BgwEmployee = New System.ComponentModel.BackgroundWorker
        Me.BgwDepartment = New System.ComponentModel.BackgroundWorker
        Me.BgwJobPosition = New System.ComponentModel.BackgroundWorker
        Me.CboFilterEmployee = New System.Windows.Forms.ComboBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Lbl_JB_Unchecked
        '
        Me.Lbl_JB_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JB_Unchecked.AutoSize = True
        Me.Lbl_JB_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JB_Unchecked.Location = New System.Drawing.Point(315, 0)
        Me.Lbl_JB_Unchecked.Name = "Lbl_JB_Unchecked"
        Me.Lbl_JB_Unchecked.Size = New System.Drawing.Size(65, 13)
        Me.Lbl_JB_Unchecked.TabIndex = 2
        Me.Lbl_JB_Unchecked.TabStop = True
        Me.Lbl_JB_Unchecked.Text = "Uncheck All"
        '
        'Lbl_JB_Checked
        '
        Me.Lbl_JB_Checked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_JB_Checked.AutoSize = True
        Me.Lbl_JB_Checked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_JB_Checked.Location = New System.Drawing.Point(257, 0)
        Me.Lbl_JB_Checked.Name = "Lbl_JB_Checked"
        Me.Lbl_JB_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_JB_Checked.TabIndex = 1
        Me.Lbl_JB_Checked.TabStop = True
        Me.Lbl_JB_Checked.Text = "Check All"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(765, 186)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 91
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.ChkJobPosition)
        Me.GroupBox1.Controls.Add(Me.Lbl_JB_Unchecked)
        Me.GroupBox1.Controls.Add(Me.Lbl_JB_Checked)
        Me.GroupBox1.Controls.Add(Me.clbJobPosition)
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(386, 126)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'ChkJobPosition
        '
        Me.ChkJobPosition.AutoSize = True
        Me.ChkJobPosition.Location = New System.Drawing.Point(6, -1)
        Me.ChkJobPosition.Name = "ChkJobPosition"
        Me.ChkJobPosition.Size = New System.Drawing.Size(83, 17)
        Me.ChkJobPosition.TabIndex = 0
        Me.ChkJobPosition.Text = "Job Position"
        Me.ChkJobPosition.UseVisualStyleBackColor = True
        '
        'clbJobPosition
        '
        Me.clbJobPosition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbJobPosition.CheckOnClick = True
        Me.clbJobPosition.FormattingEnabled = True
        Me.clbJobPosition.Location = New System.Drawing.Point(3, 16)
        Me.clbJobPosition.Name = "clbJobPosition"
        Me.clbJobPosition.Size = New System.Drawing.Size(380, 94)
        Me.clbJobPosition.TabIndex = 3
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Location = New System.Drawing.Point(12, 148)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(860, 408)
        Me.crv1.TabIndex = 4
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'btnShow
        '
        Me.btnShow.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnShow.Location = New System.Drawing.Point(795, 112)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(81, 23)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.ChkDepartement)
        Me.GroupBox2.Controls.Add(Me.Lbl_Dept_Unchecked)
        Me.GroupBox2.Controls.Add(Me.Lbl_Dept_Checked)
        Me.GroupBox2.Controls.Add(Me.ClbDepartement)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(385, 126)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'ChkDepartement
        '
        Me.ChkDepartement.AutoSize = True
        Me.ChkDepartement.Location = New System.Drawing.Point(6, -1)
        Me.ChkDepartement.Name = "ChkDepartement"
        Me.ChkDepartement.Size = New System.Drawing.Size(81, 17)
        Me.ChkDepartement.TabIndex = 0
        Me.ChkDepartement.Text = "Department"
        Me.ChkDepartement.UseVisualStyleBackColor = True
        '
        'Lbl_Dept_Unchecked
        '
        Me.Lbl_Dept_Unchecked.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Lbl_Dept_Unchecked.AutoSize = True
        Me.Lbl_Dept_Unchecked.BackColor = System.Drawing.SystemColors.Control
        Me.Lbl_Dept_Unchecked.Location = New System.Drawing.Point(314, 0)
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
        Me.Lbl_Dept_Checked.Location = New System.Drawing.Point(256, 0)
        Me.Lbl_Dept_Checked.Name = "Lbl_Dept_Checked"
        Me.Lbl_Dept_Checked.Size = New System.Drawing.Size(52, 13)
        Me.Lbl_Dept_Checked.TabIndex = 1
        Me.Lbl_Dept_Checked.TabStop = True
        Me.Lbl_Dept_Checked.Text = "Check All"
        '
        'ClbDepartement
        '
        Me.ClbDepartement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClbDepartement.CheckOnClick = True
        Me.ClbDepartement.FormattingEnabled = True
        Me.ClbDepartement.Location = New System.Drawing.Point(3, 16)
        Me.ClbDepartement.Name = "ClbDepartement"
        Me.ClbDepartement.Size = New System.Drawing.Size(379, 94)
        Me.ClbDepartement.TabIndex = 3
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(12, 11)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
        Me.SplitContainer1.Size = New System.Drawing.Size(778, 126)
        Me.SplitContainer1.SplitterDistance = 390
        Me.SplitContainer1.TabIndex = 96
        '
        'RboFormat1
        '
        Me.RboFormat1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RboFormat1.AutoSize = True
        Me.RboFormat1.Checked = True
        Me.RboFormat1.Location = New System.Drawing.Point(796, 12)
        Me.RboFormat1.Name = "RboFormat1"
        Me.RboFormat1.Size = New System.Drawing.Size(66, 17)
        Me.RboFormat1.TabIndex = 0
        Me.RboFormat1.TabStop = True
        Me.RboFormat1.Text = "Format 1"
        Me.RboFormat1.UseVisualStyleBackColor = True
        '
        'RboFormat2
        '
        Me.RboFormat2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RboFormat2.AutoSize = True
        Me.RboFormat2.Location = New System.Drawing.Point(796, 35)
        Me.RboFormat2.Name = "RboFormat2"
        Me.RboFormat2.Size = New System.Drawing.Size(66, 17)
        Me.RboFormat2.TabIndex = 1
        Me.RboFormat2.Text = "Format 2"
        Me.RboFormat2.UseVisualStyleBackColor = True
        '
        'RboFamilyData
        '
        Me.RboFamilyData.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RboFamilyData.AutoSize = True
        Me.RboFamilyData.Location = New System.Drawing.Point(796, 58)
        Me.RboFamilyData.Name = "RboFamilyData"
        Me.RboFamilyData.Size = New System.Drawing.Size(80, 17)
        Me.RboFamilyData.TabIndex = 2
        Me.RboFamilyData.Text = "Family Data"
        Me.RboFamilyData.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.TabControl2)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.cboPeriod)
        Me.GroupBox3.Controls.Add(Me.dtpFrom)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.dtpTo)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Location = New System.Drawing.Point(13, -2)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox3.Size = New System.Drawing.Size(778, 145)
        Me.GroupBox3.TabIndex = 98
        Me.GroupBox3.TabStop = False
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Location = New System.Drawing.Point(200, 17)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(574, 124)
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
        Me.TabPage5.Size = New System.Drawing.Size(566, 98)
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
        Me.dgvEmployee.Size = New System.Drawing.Size(562, 94)
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
        Me.TabPage6.Size = New System.Drawing.Size(566, 98)
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
        Me.dgvDepartment.Size = New System.Drawing.Size(562, 94)
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
        Me.TabPage7.Size = New System.Drawing.Size(566, 98)
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
        Me.dgvJobPosition.Size = New System.Drawing.Size(562, 94)
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
        'BgwEmployee
        '
        '
        'BgwDepartment
        '
        '
        'BgwJobPosition
        '
        '
        'CboFilterEmployee
        '
        Me.CboFilterEmployee.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboFilterEmployee.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFilterEmployee.FormattingEnabled = True
        Me.CboFilterEmployee.Items.AddRange(New Object() {"All", "Active", "Not Active"})
        Me.CboFilterEmployee.Location = New System.Drawing.Point(795, 83)
        Me.CboFilterEmployee.Margin = New System.Windows.Forms.Padding(2)
        Me.CboFilterEmployee.Name = "CboFilterEmployee"
        Me.CboFilterEmployee.Size = New System.Drawing.Size(81, 21)
        Me.CboFilterEmployee.TabIndex = 99
        '
        'FrmLaporanEmployeeData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(884, 562)
        Me.Controls.Add(Me.CboFilterEmployee)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.crv1)
        Me.Controls.Add(Me.RboFamilyData)
        Me.Controls.Add(Me.RboFormat2)
        Me.Controls.Add(Me.RboFormat1)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnShow)
        Me.MinimizeBox = False
        Me.Name = "FrmLaporanEmployeeData"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Employee Data Report"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        CType(Me.dgvEmployee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.dgvDepartment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        CType(Me.dgvJobPosition, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Lbl_JB_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_JB_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents clbJobPosition As System.Windows.Forms.CheckedListBox
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents ChkJobPosition As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkDepartement As System.Windows.Forms.CheckBox
    Friend WithEvents Lbl_Dept_Unchecked As System.Windows.Forms.LinkLabel
    Friend WithEvents Lbl_Dept_Checked As System.Windows.Forms.LinkLabel
    Friend WithEvents ClbDepartement As System.Windows.Forms.CheckedListBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents RboFormat1 As System.Windows.Forms.RadioButton
    Friend WithEvents RboFormat2 As System.Windows.Forms.RadioButton
    Friend WithEvents RboFamilyData As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BgwEmployee As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwDepartment As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwJobPosition As System.ComponentModel.BackgroundWorker
    Friend WithEvents CboFilterEmployee As System.Windows.Forms.ComboBox
End Class
