<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_NewSch4
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
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.txtCode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimary = New System.Windows.Forms.TextBox
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
        Me.ColAll_SetRef = New System.Windows.Forms.DataGridViewButtonColumn
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.DgvWUAllocate_WU = New System.Windows.Forms.DataGridView
        Me.CoAll_WU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Allocate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CoAll_WU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_ProdRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Efficiency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PlanRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(423, 412)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(50, 22)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(369, 412)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(50, 22)
        Me.btnOK.TabIndex = 6
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'txtCode
        '
        Me.txtCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCode.Location = New System.Drawing.Point(151, 10)
        Me.txtCode.Margin = New System.Windows.Forms.Padding(2)
        Me.txtCode.MaxLength = 250
        Me.txtCode.Name = "txtCode"
        Me.txtCode.Size = New System.Drawing.Size(334, 20)
        Me.txtCode.TabIndex = 0
        Me.txtCode.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 12)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(122, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Predefined Setting Code"
        '
        'txtPrimary
        '
        Me.txtPrimary.Location = New System.Drawing.Point(151, 10)
        Me.txtPrimary.Margin = New System.Windows.Forms.Padding(2)
        Me.txtPrimary.Name = "txtPrimary"
        Me.txtPrimary.Size = New System.Drawing.Size(35, 20)
        Me.txtPrimary.TabIndex = 10
        Me.txtPrimary.Text = "-1"
        Me.txtPrimary.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 39)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(479, 366)
        Me.GroupBox1.TabIndex = 2
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
        Me.SplitContainer1.Size = New System.Drawing.Size(475, 349)
        Me.SplitContainer1.SplitterDistance = 177
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
        Me.GroupBox2.Size = New System.Drawing.Size(475, 177)
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
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle9.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvWUAllocate_Process.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle9
        Me.DgvWUAllocate_Process.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_Process.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAll_Process_No, Me.ColAll_Process_Code, Me.ColAll_Process_CodeNo, Me.ColAll_Process_Desc, Me.ColAll_Process_PrimaryProcess, Me.ColAll_Process_PrimaryCodeNo, Me.ColAll_IdentifierProcess, Me.ColAll_SetRef})
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DgvWUAllocate_Process.DefaultCellStyle = DataGridViewCellStyle10
        Me.DgvWUAllocate_Process.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWUAllocate_Process.Location = New System.Drawing.Point(2, 16)
        Me.DgvWUAllocate_Process.MultiSelect = False
        Me.DgvWUAllocate_Process.Name = "DgvWUAllocate_Process"
        Me.DgvWUAllocate_Process.ReadOnly = True
        Me.DgvWUAllocate_Process.RowHeadersVisible = False
        Me.DgvWUAllocate_Process.RowTemplate.Height = 24
        Me.DgvWUAllocate_Process.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvWUAllocate_Process.Size = New System.Drawing.Size(471, 158)
        Me.DgvWUAllocate_Process.TabIndex = 1
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
        Me.ColAll_Process_PrimaryProcess.Width = 117
        '
        'ColAll_Process_PrimaryCodeNo
        '
        Me.ColAll_Process_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColAll_Process_PrimaryCodeNo.Name = "ColAll_Process_PrimaryCodeNo"
        Me.ColAll_Process_PrimaryCodeNo.ReadOnly = True
        Me.ColAll_Process_PrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_PrimaryCodeNo.Visible = False
        Me.ColAll_Process_PrimaryCodeNo.Width = 121
        '
        'ColAll_IdentifierProcess
        '
        Me.ColAll_IdentifierProcess.HeaderText = "Identifier"
        Me.ColAll_IdentifierProcess.Name = "ColAll_IdentifierProcess"
        Me.ColAll_IdentifierProcess.ReadOnly = True
        Me.ColAll_IdentifierProcess.Visible = False
        Me.ColAll_IdentifierProcess.Width = 87
        '
        'ColAll_SetRef
        '
        Me.ColAll_SetRef.HeaderText = "Set Reference"
        Me.ColAll_SetRef.Name = "ColAll_SetRef"
        Me.ColAll_SetRef.ReadOnly = True
        Me.ColAll_SetRef.Width = 82
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.llbCheckAll)
        Me.GroupBox3.Controls.Add(Me.llbUncheckAll)
        Me.GroupBox3.Controls.Add(Me.DgvWUAllocate_WU)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.GroupBox3.Size = New System.Drawing.Size(475, 169)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Work Unit"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(350, -1)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAll.TabIndex = 5
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        Me.llbCheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(408, -1)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAll.TabIndex = 4
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        Me.llbUncheckAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'DgvWUAllocate_WU
        '
        Me.DgvWUAllocate_WU.AllowUserToAddRows = False
        Me.DgvWUAllocate_WU.AllowUserToDeleteRows = False
        Me.DgvWUAllocate_WU.AllowUserToResizeRows = False
        Me.DgvWUAllocate_WU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWUAllocate_WU.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvWUAllocate_WU.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle11
        Me.DgvWUAllocate_WU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_WU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CoAll_WU_No, Me.CoAll_WU_Allocate, Me.CoAll_WU_Code, Me.CoAll_WU_Desc, Me.CoAll_WU_ProdRate, Me.CoAll_WU_Efficiency, Me.CoAll_WU_PlanRate, Me.CoAll_WU_PrimaryWU})
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle15.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DgvWUAllocate_WU.DefaultCellStyle = DataGridViewCellStyle15
        Me.DgvWUAllocate_WU.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWUAllocate_WU.Location = New System.Drawing.Point(2, 16)
        Me.DgvWUAllocate_WU.MultiSelect = False
        Me.DgvWUAllocate_WU.Name = "DgvWUAllocate_WU"
        Me.DgvWUAllocate_WU.ReadOnly = True
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DgvWUAllocate_WU.RowHeadersDefaultCellStyle = DataGridViewCellStyle16
        Me.DgvWUAllocate_WU.RowHeadersVisible = False
        Me.DgvWUAllocate_WU.RowTemplate.Height = 24
        Me.DgvWUAllocate_WU.Size = New System.Drawing.Size(471, 150)
        Me.DgvWUAllocate_WU.TabIndex = 3
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
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_ProdRate.DefaultCellStyle = DataGridViewCellStyle12
        Me.CoAll_WU_ProdRate.HeaderText = "Rate (Unit/Hour)"
        Me.CoAll_WU_ProdRate.Name = "CoAll_WU_ProdRate"
        Me.CoAll_WU_ProdRate.ReadOnly = True
        Me.CoAll_WU_ProdRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_ProdRate.Width = 92
        '
        'CoAll_WU_Efficiency
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_Efficiency.DefaultCellStyle = DataGridViewCellStyle13
        Me.CoAll_WU_Efficiency.HeaderText = "Efficiency (%)"
        Me.CoAll_WU_Efficiency.Name = "CoAll_WU_Efficiency"
        Me.CoAll_WU_Efficiency.ReadOnly = True
        Me.CoAll_WU_Efficiency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Efficiency.Width = 76
        '
        'CoAll_WU_PlanRate
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_PlanRate.DefaultCellStyle = DataGridViewCellStyle14
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
        'FrmTrans_ProductionSchedule_NewSch4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(490, 446)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtCode)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtPrimary)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(500, 483)
        Me.Name = "FrmTrans_ProductionSchedule_NewSch4"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add New Work Unit Configuration"
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DgvWUAllocate_Process, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.DgvWUAllocate_WU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents txtCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWUAllocate_Process As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWUAllocate_WU As System.Windows.Forms.DataGridView
    Friend WithEvents CoAll_WU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Allocate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CoAll_WU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_ProdRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Efficiency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PlanRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents ColAll_Process_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_IdentifierProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_SetRef As System.Windows.Forms.DataGridViewButtonColumn
End Class
