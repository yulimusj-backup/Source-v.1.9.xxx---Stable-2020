<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch3
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.rboAutomatic = New System.Windows.Forms.RadioButton
        Me.rboManual = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DgvWUAllocate_Process = New System.Windows.Forms.DataGridView
        Me.ColAll_Process_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_IdentifierProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.DgvWUAllocate_WU = New System.Windows.Forms.DataGridView
        Me.CoAll_WU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Allocate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CoAll_WU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_ProdRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Efficiency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PlanRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.cboWorkUnitConfig = New System.Windows.Forms.ComboBox
        Me.llbAdd = New System.Windows.Forms.LinkLabel
        Me.llbDelete = New System.Windows.Forms.LinkLabel
        Me.llbEdit = New System.Windows.Forms.LinkLabel
        Me.btnStop = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvWUAllocate_Process, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DgvWUAllocate_WU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 7)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select scheduling plan:"
        '
        'rboAutomatic
        '
        Me.rboAutomatic.AutoSize = True
        Me.rboAutomatic.Checked = True
        Me.rboAutomatic.Location = New System.Drawing.Point(9, 24)
        Me.rboAutomatic.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboAutomatic.Name = "rboAutomatic"
        Me.rboAutomatic.Size = New System.Drawing.Size(241, 17)
        Me.rboAutomatic.TabIndex = 0
        Me.rboAutomatic.TabStop = True
        Me.rboAutomatic.Text = "Let the software optimize production schedule"
        Me.rboAutomatic.UseVisualStyleBackColor = True
        '
        'rboManual
        '
        Me.rboManual.AutoSize = True
        Me.rboManual.Location = New System.Drawing.Point(9, 45)
        Me.rboManual.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboManual.Name = "rboManual"
        Me.rboManual.Size = New System.Drawing.Size(241, 17)
        Me.rboManual.TabIndex = 1
        Me.rboManual.Text = "Select from predefined work unit configuration"
        Me.rboManual.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 67)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(550, 365)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Work Unit Configuration"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(2, 15)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox3)
        Me.SplitContainer1.Size = New System.Drawing.Size(546, 348)
        Me.SplitContainer1.SplitterDistance = 179
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 1
        Me.SplitContainer1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DgvWUAllocate_Process)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox2.Size = New System.Drawing.Size(546, 179)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Process"
        '
        'DgvWUAllocate_Process
        '
        Me.DgvWUAllocate_Process.AllowUserToAddRows = False
        Me.DgvWUAllocate_Process.AllowUserToDeleteRows = False
        Me.DgvWUAllocate_Process.AllowUserToResizeRows = False
        Me.DgvWUAllocate_Process.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWUAllocate_Process.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvWUAllocate_Process.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_Process.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAll_Process_No, Me.ColAll_Process_Code, Me.ColAll_Process_CodeNo, Me.ColAll_Process_Desc, Me.ColAll_Process_PrimaryProcess, Me.ColAll_Process_PrimaryCodeNo, Me.ColAll_IdentifierProcess})
        Me.DgvWUAllocate_Process.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWUAllocate_Process.Location = New System.Drawing.Point(2, 16)
        Me.DgvWUAllocate_Process.MultiSelect = False
        Me.DgvWUAllocate_Process.Name = "DgvWUAllocate_Process"
        Me.DgvWUAllocate_Process.ReadOnly = True
        Me.DgvWUAllocate_Process.RowHeadersVisible = False
        Me.DgvWUAllocate_Process.RowTemplate.Height = 24
        Me.DgvWUAllocate_Process.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvWUAllocate_Process.Size = New System.Drawing.Size(542, 160)
        Me.DgvWUAllocate_Process.TabIndex = 0
        '
        'ColAll_Process_No
        '
        Me.ColAll_Process_No.HeaderText = "No"
        Me.ColAll_Process_No.Name = "ColAll_Process_No"
        Me.ColAll_Process_No.ReadOnly = True
        Me.ColAll_Process_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_No.Width = 27
        '
        'ColAll_Process_Code
        '
        Me.ColAll_Process_Code.HeaderText = "Code"
        Me.ColAll_Process_Code.Name = "ColAll_Process_Code"
        Me.ColAll_Process_Code.ReadOnly = True
        Me.ColAll_Process_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Code.Width = 38
        '
        'ColAll_Process_CodeNo
        '
        Me.ColAll_Process_CodeNo.HeaderText = "Code No"
        Me.ColAll_Process_CodeNo.Name = "ColAll_Process_CodeNo"
        Me.ColAll_Process_CodeNo.ReadOnly = True
        Me.ColAll_Process_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_CodeNo.Width = 55
        '
        'ColAll_Process_Desc
        '
        Me.ColAll_Process_Desc.HeaderText = "Description"
        Me.ColAll_Process_Desc.Name = "ColAll_Process_Desc"
        Me.ColAll_Process_Desc.ReadOnly = True
        Me.ColAll_Process_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Desc.Width = 66
        '
        'ColAll_Process_PrimaryProcess
        '
        Me.ColAll_Process_PrimaryProcess.HeaderText = "Primary Process"
        Me.ColAll_Process_PrimaryProcess.Name = "ColAll_Process_PrimaryProcess"
        Me.ColAll_Process_PrimaryProcess.ReadOnly = True
        Me.ColAll_Process_PrimaryProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_PrimaryProcess.Visible = False
        Me.ColAll_Process_PrimaryProcess.Width = 88
        '
        'ColAll_Process_PrimaryCodeNo
        '
        Me.ColAll_Process_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColAll_Process_PrimaryCodeNo.Name = "ColAll_Process_PrimaryCodeNo"
        Me.ColAll_Process_PrimaryCodeNo.ReadOnly = True
        Me.ColAll_Process_PrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_PrimaryCodeNo.Visible = False
        Me.ColAll_Process_PrimaryCodeNo.Width = 92
        '
        'ColAll_IdentifierProcess
        '
        Me.ColAll_IdentifierProcess.HeaderText = "Identifier"
        Me.ColAll_IdentifierProcess.Name = "ColAll_IdentifierProcess"
        Me.ColAll_IdentifierProcess.ReadOnly = True
        Me.ColAll_IdentifierProcess.Visible = False
        Me.ColAll_IdentifierProcess.Width = 72
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DgvWUAllocate_WU)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox3.Size = New System.Drawing.Size(546, 166)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Work Unit"
        '
        'DgvWUAllocate_WU
        '
        Me.DgvWUAllocate_WU.AllowUserToAddRows = False
        Me.DgvWUAllocate_WU.AllowUserToDeleteRows = False
        Me.DgvWUAllocate_WU.AllowUserToResizeRows = False
        Me.DgvWUAllocate_WU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWUAllocate_WU.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvWUAllocate_WU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_WU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CoAll_WU_No, Me.CoAll_WU_Allocate, Me.CoAll_WU_Code, Me.CoAll_WU_Desc, Me.CoAll_WU_ProdRate, Me.CoAll_WU_Efficiency, Me.CoAll_WU_PlanRate, Me.CoAll_WU_PrimaryWU, Me.CoAll_WU_StartDate})
        Me.DgvWUAllocate_WU.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWUAllocate_WU.Location = New System.Drawing.Point(2, 16)
        Me.DgvWUAllocate_WU.MultiSelect = False
        Me.DgvWUAllocate_WU.Name = "DgvWUAllocate_WU"
        Me.DgvWUAllocate_WU.ReadOnly = True
        Me.DgvWUAllocate_WU.RowHeadersVisible = False
        Me.DgvWUAllocate_WU.RowTemplate.Height = 24
        Me.DgvWUAllocate_WU.Size = New System.Drawing.Size(542, 147)
        Me.DgvWUAllocate_WU.TabIndex = 0
        '
        'CoAll_WU_No
        '
        Me.CoAll_WU_No.HeaderText = "Priority"
        Me.CoAll_WU_No.Name = "CoAll_WU_No"
        Me.CoAll_WU_No.ReadOnly = True
        Me.CoAll_WU_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_No.Width = 44
        '
        'CoAll_WU_Allocate
        '
        Me.CoAll_WU_Allocate.HeaderText = "Use"
        Me.CoAll_WU_Allocate.Name = "CoAll_WU_Allocate"
        Me.CoAll_WU_Allocate.ReadOnly = True
        Me.CoAll_WU_Allocate.Width = 32
        '
        'CoAll_WU_Code
        '
        Me.CoAll_WU_Code.HeaderText = "Code"
        Me.CoAll_WU_Code.Name = "CoAll_WU_Code"
        Me.CoAll_WU_Code.ReadOnly = True
        Me.CoAll_WU_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Code.Width = 38
        '
        'CoAll_WU_Desc
        '
        Me.CoAll_WU_Desc.HeaderText = "Description"
        Me.CoAll_WU_Desc.Name = "CoAll_WU_Desc"
        Me.CoAll_WU_Desc.ReadOnly = True
        Me.CoAll_WU_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Desc.Width = 66
        '
        'CoAll_WU_ProdRate
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_ProdRate.DefaultCellStyle = DataGridViewCellStyle1
        Me.CoAll_WU_ProdRate.HeaderText = "Rate (Unit/Hour)"
        Me.CoAll_WU_ProdRate.Name = "CoAll_WU_ProdRate"
        Me.CoAll_WU_ProdRate.ReadOnly = True
        Me.CoAll_WU_ProdRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_ProdRate.Width = 92
        '
        'CoAll_WU_Efficiency
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_Efficiency.DefaultCellStyle = DataGridViewCellStyle2
        Me.CoAll_WU_Efficiency.HeaderText = "Efficiency (%)"
        Me.CoAll_WU_Efficiency.Name = "CoAll_WU_Efficiency"
        Me.CoAll_WU_Efficiency.ReadOnly = True
        Me.CoAll_WU_Efficiency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Efficiency.Width = 76
        '
        'CoAll_WU_PlanRate
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_PlanRate.DefaultCellStyle = DataGridViewCellStyle3
        Me.CoAll_WU_PlanRate.HeaderText = "Plan Rate (Unit/Hour)"
        Me.CoAll_WU_PlanRate.Name = "CoAll_WU_PlanRate"
        Me.CoAll_WU_PlanRate.ReadOnly = True
        Me.CoAll_WU_PlanRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_PlanRate.Width = 116
        '
        'CoAll_WU_PrimaryWU
        '
        Me.CoAll_WU_PrimaryWU.HeaderText = "Primary WU"
        Me.CoAll_WU_PrimaryWU.Name = "CoAll_WU_PrimaryWU"
        Me.CoAll_WU_PrimaryWU.ReadOnly = True
        Me.CoAll_WU_PrimaryWU.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_PrimaryWU.Visible = False
        Me.CoAll_WU_PrimaryWU.Width = 89
        '
        'CoAll_WU_StartDate
        '
        Me.CoAll_WU_StartDate.HeaderText = "Start"
        Me.CoAll_WU_StartDate.Name = "CoAll_WU_StartDate"
        Me.CoAll_WU_StartDate.ReadOnly = True
        Me.CoAll_WU_StartDate.Width = 54
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(391, 10)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(50, 22)
        Me.btnOK.TabIndex = 7
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(445, 10)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(50, 22)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'cboWorkUnitConfig
        '
        Me.cboWorkUnitConfig.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboWorkUnitConfig.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWorkUnitConfig.Enabled = False
        Me.cboWorkUnitConfig.FormattingEnabled = True
        Me.cboWorkUnitConfig.Location = New System.Drawing.Point(252, 43)
        Me.cboWorkUnitConfig.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.cboWorkUnitConfig.Name = "cboWorkUnitConfig"
        Me.cboWorkUnitConfig.Size = New System.Drawing.Size(200, 21)
        Me.cboWorkUnitConfig.TabIndex = 2
        '
        'llbAdd
        '
        Me.llbAdd.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAdd.AutoSize = True
        Me.llbAdd.Enabled = False
        Me.llbAdd.Location = New System.Drawing.Point(456, 47)
        Me.llbAdd.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbAdd.Name = "llbAdd"
        Me.llbAdd.Size = New System.Drawing.Size(26, 13)
        Me.llbAdd.TabIndex = 3
        Me.llbAdd.TabStop = True
        Me.llbAdd.Text = "Add"
        Me.llbAdd.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbDelete
        '
        Me.llbDelete.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbDelete.AutoSize = True
        Me.llbDelete.Enabled = False
        Me.llbDelete.Location = New System.Drawing.Point(514, 47)
        Me.llbDelete.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbDelete.Name = "llbDelete"
        Me.llbDelete.Size = New System.Drawing.Size(38, 13)
        Me.llbDelete.TabIndex = 5
        Me.llbDelete.TabStop = True
        Me.llbDelete.Text = "Delete"
        Me.llbDelete.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbEdit
        '
        Me.llbEdit.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbEdit.AutoSize = True
        Me.llbEdit.Enabled = False
        Me.llbEdit.Location = New System.Drawing.Point(485, 47)
        Me.llbEdit.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbEdit.Name = "llbEdit"
        Me.llbEdit.Size = New System.Drawing.Size(25, 13)
        Me.llbEdit.TabIndex = 4
        Me.llbEdit.TabStop = True
        Me.llbEdit.Text = "Edit"
        Me.llbEdit.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'btnStop
        '
        Me.btnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStop.Location = New System.Drawing.Point(500, 10)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(50, 22)
        Me.btnStop.TabIndex = 9
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'FrmTrans_ProductionSchedule_NewSch3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(568, 445)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.llbEdit)
        Me.Controls.Add(Me.llbDelete)
        Me.Controls.Add(Me.llbAdd)
        Me.Controls.Add(Me.cboWorkUnitConfig)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.rboManual)
        Me.Controls.Add(Me.rboAutomatic)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(496, 483)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch3"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Please select scheduling plan"
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DgvWUAllocate_Process, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.DgvWUAllocate_WU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents rboAutomatic As System.Windows.Forms.RadioButton
    Friend WithEvents rboManual As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWUAllocate_Process As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWUAllocate_WU As System.Windows.Forms.DataGridView
    Friend WithEvents cboWorkUnitConfig As System.Windows.Forms.ComboBox
    Friend WithEvents llbAdd As System.Windows.Forms.LinkLabel
    Friend WithEvents llbDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents llbEdit As System.Windows.Forms.LinkLabel
    Friend WithEvents ColAll_Process_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_IdentifierProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents CoAll_WU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Allocate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CoAll_WU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_ProdRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Efficiency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PlanRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
