<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTest
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
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvDown = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCombi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnStart = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtJmlWU = New System.Windows.Forms.TextBox
        Me.TxtWUUsedQty = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtMONumber = New System.Windows.Forms.TextBox
        Me.btnFind = New System.Windows.Forms.Button
        Me.cboJONumber = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.DgvWUAllocate_Process = New System.Windows.Forms.DataGridView
        Me.ColAll_Process_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Rate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_Process_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAll_UnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DgvWUAllocate_WU = New System.Windows.Forms.DataGridView
        Me.CoAll_WU_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Allocate = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.CoAll_WU_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_ProdRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_Efficiency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PlanRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_PrimaryWU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_ProduceQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_StartDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_StartTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_EndDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CoAll_WU_EndTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvWUAllocate_Process, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvWUAllocate_WU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDown
        '
        Me.dgvDown.AllowUserToAddRows = False
        Me.dgvDown.AllowUserToDeleteRows = False
        Me.dgvDown.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvDown.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvDown.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDown.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCombi})
        Me.dgvDown.Location = New System.Drawing.Point(6, 6)
        Me.dgvDown.Name = "dgvDown"
        Me.dgvDown.ReadOnly = True
        Me.dgvDown.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDown.RowTemplate.Height = 24
        Me.dgvDown.Size = New System.Drawing.Size(320, 479)
        Me.dgvDown.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        '
        'ColCombi
        '
        Me.ColCombi.HeaderText = "Combination"
        Me.ColCombi.Name = "ColCombi"
        Me.ColCombi.ReadOnly = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(472, 62)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(61, 23)
        Me.btnStart.TabIndex = 1
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(332, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Work Unit Available"
        '
        'TxtJmlWU
        '
        Me.TxtJmlWU.Location = New System.Drawing.Point(472, 6)
        Me.TxtJmlWU.Name = "TxtJmlWU"
        Me.TxtJmlWU.Size = New System.Drawing.Size(61, 22)
        Me.TxtJmlWU.TabIndex = 3
        Me.TxtJmlWU.Text = "0"
        Me.TxtJmlWU.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtWUUsedQty
        '
        Me.TxtWUUsedQty.Location = New System.Drawing.Point(472, 34)
        Me.TxtWUUsedQty.Name = "TxtWUUsedQty"
        Me.TxtWUUsedQty.Size = New System.Drawing.Size(61, 22)
        Me.TxtWUUsedQty.TabIndex = 5
        Me.TxtWUUsedQty.Text = "0"
        Me.TxtWUUsedQty.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(332, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(107, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Work Unit Used"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(661, 520)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvDown)
        Me.TabPage1.Controls.Add(Me.TxtWUUsedQty)
        Me.TabPage1.Controls.Add(Me.TxtJmlWU)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.btnStart)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(653, 491)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "TabPage1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox2)
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.cboJONumber)
        Me.TabPage2.Controls.Add(Me.btnFind)
        Me.TabPage2.Controls.Add(Me.TxtMONumber)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(653, 491)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "TabPage2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 17)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "MO Number"
        '
        'TxtMONumber
        '
        Me.TxtMONumber.Location = New System.Drawing.Point(110, 6)
        Me.TxtMONumber.Name = "TxtMONumber"
        Me.TxtMONumber.Size = New System.Drawing.Size(186, 22)
        Me.TxtMONumber.TabIndex = 1
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(302, 6)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(45, 23)
        Me.btnFind.TabIndex = 2
        Me.btnFind.Text = "Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'cboJONumber
        '
        Me.cboJONumber.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboJONumber.FormattingEnabled = True
        Me.cboJONumber.Location = New System.Drawing.Point(110, 34)
        Me.cboJONumber.Name = "cboJONumber"
        Me.cboJONumber.Size = New System.Drawing.Size(186, 24)
        Me.cboJONumber.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 37)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 17)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "JO Number"
        '
        'DgvWUAllocate_Process
        '
        Me.DgvWUAllocate_Process.AllowUserToAddRows = False
        Me.DgvWUAllocate_Process.AllowUserToDeleteRows = False
        Me.DgvWUAllocate_Process.AllowUserToResizeRows = False
        Me.DgvWUAllocate_Process.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWUAllocate_Process.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvWUAllocate_Process.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_Process.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAll_Process_No, Me.ColAll_Process_Code, Me.ColAll_Process_CodeNo, Me.ColAll_Process_Desc, Me.ColAll_Process_PrimaryProcess, Me.ColAll_Process_PrimaryCodeNo, Me.ColAll_Process_Rate, Me.ColAll_Process_Qty, Me.ColAll_UnitCode})
        Me.DgvWUAllocate_Process.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWUAllocate_Process.Location = New System.Drawing.Point(3, 18)
        Me.DgvWUAllocate_Process.Margin = New System.Windows.Forms.Padding(4)
        Me.DgvWUAllocate_Process.MultiSelect = False
        Me.DgvWUAllocate_Process.Name = "DgvWUAllocate_Process"
        Me.DgvWUAllocate_Process.ReadOnly = True
        Me.DgvWUAllocate_Process.RowHeadersVisible = False
        Me.DgvWUAllocate_Process.RowTemplate.Height = 24
        Me.DgvWUAllocate_Process.Size = New System.Drawing.Size(632, 153)
        Me.DgvWUAllocate_Process.TabIndex = 5
        '
        'ColAll_Process_No
        '
        Me.ColAll_Process_No.HeaderText = "No"
        Me.ColAll_Process_No.Name = "ColAll_Process_No"
        Me.ColAll_Process_No.ReadOnly = True
        Me.ColAll_Process_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_No.Width = 32
        '
        'ColAll_Process_Code
        '
        Me.ColAll_Process_Code.HeaderText = "Code"
        Me.ColAll_Process_Code.Name = "ColAll_Process_Code"
        Me.ColAll_Process_Code.ReadOnly = True
        Me.ColAll_Process_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Code.Width = 47
        '
        'ColAll_Process_CodeNo
        '
        Me.ColAll_Process_CodeNo.HeaderText = "Code No"
        Me.ColAll_Process_CodeNo.Name = "ColAll_Process_CodeNo"
        Me.ColAll_Process_CodeNo.ReadOnly = True
        Me.ColAll_Process_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_CodeNo.Width = 69
        '
        'ColAll_Process_Desc
        '
        Me.ColAll_Process_Desc.HeaderText = "Description"
        Me.ColAll_Process_Desc.Name = "ColAll_Process_Desc"
        Me.ColAll_Process_Desc.ReadOnly = True
        Me.ColAll_Process_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Desc.Width = 85
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
        'ColAll_Process_Rate
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAll_Process_Rate.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColAll_Process_Rate.HeaderText = "Total Rate (Unit/Hour)"
        Me.ColAll_Process_Rate.Name = "ColAll_Process_Rate"
        Me.ColAll_Process_Rate.ReadOnly = True
        Me.ColAll_Process_Rate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Rate.Width = 154
        '
        'ColAll_Process_Qty
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAll_Process_Qty.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColAll_Process_Qty.HeaderText = "Qty Produce"
        Me.ColAll_Process_Qty.Name = "ColAll_Process_Qty"
        Me.ColAll_Process_Qty.ReadOnly = True
        Me.ColAll_Process_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_Process_Qty.Width = 93
        '
        'ColAll_UnitCode
        '
        Me.ColAll_UnitCode.HeaderText = "Unit"
        Me.ColAll_UnitCode.Name = "ColAll_UnitCode"
        Me.ColAll_UnitCode.ReadOnly = True
        Me.ColAll_UnitCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAll_UnitCode.Width = 39
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DgvWUAllocate_Process)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 64)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(638, 174)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Process"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.DgvWUAllocate_WU)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 245)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(638, 239)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Work Unit"
        '
        'DgvWUAllocate_WU
        '
        Me.DgvWUAllocate_WU.AllowUserToAddRows = False
        Me.DgvWUAllocate_WU.AllowUserToDeleteRows = False
        Me.DgvWUAllocate_WU.AllowUserToResizeRows = False
        Me.DgvWUAllocate_WU.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvWUAllocate_WU.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvWUAllocate_WU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvWUAllocate_WU.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CoAll_WU_No, Me.CoAll_WU_Allocate, Me.CoAll_WU_Code, Me.CoAll_WU_Desc, Me.CoAll_WU_ProdRate, Me.CoAll_WU_Efficiency, Me.CoAll_WU_PlanRate, Me.CoAll_WU_PrimaryWU, Me.CoAll_WU_ProduceQty, Me.CoAll_WU_StartDate, Me.CoAll_WU_StartTime, Me.CoAll_WU_EndDate, Me.CoAll_WU_EndTime})
        Me.DgvWUAllocate_WU.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvWUAllocate_WU.Location = New System.Drawing.Point(3, 19)
        Me.DgvWUAllocate_WU.Margin = New System.Windows.Forms.Padding(4)
        Me.DgvWUAllocate_WU.MultiSelect = False
        Me.DgvWUAllocate_WU.Name = "DgvWUAllocate_WU"
        Me.DgvWUAllocate_WU.RowHeadersVisible = False
        Me.DgvWUAllocate_WU.RowTemplate.Height = 24
        Me.DgvWUAllocate_WU.Size = New System.Drawing.Size(632, 216)
        Me.DgvWUAllocate_WU.TabIndex = 3
        '
        'CoAll_WU_No
        '
        Me.CoAll_WU_No.HeaderText = "No"
        Me.CoAll_WU_No.Name = "CoAll_WU_No"
        Me.CoAll_WU_No.ReadOnly = True
        Me.CoAll_WU_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_No.Width = 32
        '
        'CoAll_WU_Allocate
        '
        Me.CoAll_WU_Allocate.HeaderText = "Use"
        Me.CoAll_WU_Allocate.Name = "CoAll_WU_Allocate"
        Me.CoAll_WU_Allocate.Width = 39
        '
        'CoAll_WU_Code
        '
        Me.CoAll_WU_Code.HeaderText = "Code"
        Me.CoAll_WU_Code.Name = "CoAll_WU_Code"
        Me.CoAll_WU_Code.ReadOnly = True
        Me.CoAll_WU_Code.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Code.Width = 47
        '
        'CoAll_WU_Desc
        '
        Me.CoAll_WU_Desc.HeaderText = "Description"
        Me.CoAll_WU_Desc.Name = "CoAll_WU_Desc"
        Me.CoAll_WU_Desc.ReadOnly = True
        Me.CoAll_WU_Desc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Desc.Width = 85
        '
        'CoAll_WU_ProdRate
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_ProdRate.DefaultCellStyle = DataGridViewCellStyle3
        Me.CoAll_WU_ProdRate.HeaderText = "Rate (Unit/Hour)"
        Me.CoAll_WU_ProdRate.Name = "CoAll_WU_ProdRate"
        Me.CoAll_WU_ProdRate.ReadOnly = True
        Me.CoAll_WU_ProdRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_ProdRate.Visible = False
        Me.CoAll_WU_ProdRate.Width = 118
        '
        'CoAll_WU_Efficiency
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_Efficiency.DefaultCellStyle = DataGridViewCellStyle4
        Me.CoAll_WU_Efficiency.HeaderText = "Efficiency (%)"
        Me.CoAll_WU_Efficiency.Name = "CoAll_WU_Efficiency"
        Me.CoAll_WU_Efficiency.ReadOnly = True
        Me.CoAll_WU_Efficiency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_Efficiency.Visible = False
        '
        'CoAll_WU_PlanRate
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_PlanRate.DefaultCellStyle = DataGridViewCellStyle5
        Me.CoAll_WU_PlanRate.HeaderText = "Plan Rate (Unit/Hour)"
        Me.CoAll_WU_PlanRate.Name = "CoAll_WU_PlanRate"
        Me.CoAll_WU_PlanRate.ReadOnly = True
        Me.CoAll_WU_PlanRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_PlanRate.Width = 150
        '
        'CoAll_WU_PrimaryWU
        '
        Me.CoAll_WU_PrimaryWU.HeaderText = "Primary"
        Me.CoAll_WU_PrimaryWU.Name = "CoAll_WU_PrimaryWU"
        Me.CoAll_WU_PrimaryWU.ReadOnly = True
        Me.CoAll_WU_PrimaryWU.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_PrimaryWU.Visible = False
        Me.CoAll_WU_PrimaryWU.Width = 62
        '
        'CoAll_WU_ProduceQty
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.CoAll_WU_ProduceQty.DefaultCellStyle = DataGridViewCellStyle6
        Me.CoAll_WU_ProduceQty.HeaderText = "Qty"
        Me.CoAll_WU_ProduceQty.Name = "CoAll_WU_ProduceQty"
        Me.CoAll_WU_ProduceQty.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.CoAll_WU_ProduceQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_ProduceQty.Width = 36
        '
        'CoAll_WU_StartDate
        '
        Me.CoAll_WU_StartDate.HeaderText = "Start"
        Me.CoAll_WU_StartDate.Name = "CoAll_WU_StartDate"
        Me.CoAll_WU_StartDate.ReadOnly = True
        Me.CoAll_WU_StartDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_StartDate.Width = 44
        '
        'CoAll_WU_StartTime
        '
        Me.CoAll_WU_StartTime.HeaderText = ""
        Me.CoAll_WU_StartTime.Name = "CoAll_WU_StartTime"
        Me.CoAll_WU_StartTime.ReadOnly = True
        Me.CoAll_WU_StartTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_StartTime.Width = 5
        '
        'CoAll_WU_EndDate
        '
        Me.CoAll_WU_EndDate.HeaderText = "End"
        Me.CoAll_WU_EndDate.Name = "CoAll_WU_EndDate"
        Me.CoAll_WU_EndDate.ReadOnly = True
        Me.CoAll_WU_EndDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_EndDate.Width = 39
        '
        'CoAll_WU_EndTime
        '
        Me.CoAll_WU_EndTime.HeaderText = ""
        Me.CoAll_WU_EndTime.Name = "CoAll_WU_EndTime"
        Me.CoAll_WU_EndTime.ReadOnly = True
        Me.CoAll_WU_EndTime.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.CoAll_WU_EndTime.Width = 5
        '
        'FrmTest
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(685, 544)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "FrmTest"
        Me.Text = "FrmTest"
        CType(Me.dgvDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.DgvWUAllocate_Process, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.DgvWUAllocate_WU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDown As System.Windows.Forms.DataGridView
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtJmlWU As System.Windows.Forms.TextBox
    Friend WithEvents TxtWUUsedQty As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCombi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboJONumber As System.Windows.Forms.ComboBox
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents TxtMONumber As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWUAllocate_Process As System.Windows.Forms.DataGridView
    Friend WithEvents ColAll_Process_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Rate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_Process_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAll_UnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvWUAllocate_WU As System.Windows.Forms.DataGridView
    Friend WithEvents CoAll_WU_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Allocate As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents CoAll_WU_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_ProdRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_Efficiency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PlanRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_PrimaryWU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_ProduceQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_StartDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_StartTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_EndDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CoAll_WU_EndTime As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
