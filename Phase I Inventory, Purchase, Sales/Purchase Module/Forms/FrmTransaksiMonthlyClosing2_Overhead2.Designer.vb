<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonthlyClosing2_Overhead2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiMonthlyClosing2_Overhead2))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnUP = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.dgvItemCost_CostGroup = New System.Windows.Forms.DataGridView
        Me.ColCostGroup_GroupDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostGroup_OverheadAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostGroup_AllocationStyle = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColCostGroup_RefUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostGroup_PrimaryRefUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostGroup_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCostGroup_TotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.llbRefreshBalance = New System.Windows.Forms.LinkLabel
        Me.BtnUpAccount = New System.Windows.Forms.Button
        Me.BtnDownAccount = New System.Windows.Forms.Button
        Me.dgvItemCost_CostAccountAllocation = New System.Windows.Forms.DataGridView
        Me.ColAccAll_AccountCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccAll_AccountDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccAll_EndBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccAll_AllocatedPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccAll_AllocatedAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccAll_IgnoreNegative = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAccAll_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccAll_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccAll_PrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.BtnUpScope = New System.Windows.Forms.Button
        Me.BtnDownScope = New System.Windows.Forms.Button
        Me.dgvItemCost_Scope = New System.Windows.Forms.DataGridView
        Me.ColAllScope_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllScope_ProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllScope_ProcessDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllScope_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllScope_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAllScope_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.rboGroup_Scope = New System.Windows.Forms.RadioButton
        Me.rboGroup_Account = New System.Windows.Forms.RadioButton
        Me.dgvItemCost_AllocationRecap1 = New System.Windows.Forms.DataGridView
        Me.ColAlloRecap1_CostGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap1_AccCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap1_AccDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap1_Balance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap1_AllocatedPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap1_AllocatedAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap1_IngoreNegativeBalance = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAlloRecap1_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap1_PrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemCost_AllocationRecap2 = New System.Windows.Forms.DataGridView
        Me.ColAlloRecap2_CostGroup = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap2_ProcessCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap2_ProcessDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap2_Identifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlloRecap2_PrimaryProcess = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BgwDelCostGroup = New System.ComponentModel.BackgroundWorker
        Me.BgwDelCostScope = New System.ComponentModel.BackgroundWorker
        Me.BgwDelCostAcc = New System.ComponentModel.BackgroundWorker
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItemCost_CostGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl2.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        CType(Me.dgvItemCost_CostAccountAllocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        CType(Me.dgvItemCost_Scope, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage8.SuspendLayout()
        CType(Me.dgvItemCost_AllocationRecap1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemCost_AllocationRecap2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnCancel)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnOK)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnUP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnDown)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl2)
        Me.SplitContainer1.Size = New System.Drawing.Size(663, 457)
        Me.SplitContainer1.SplitterDistance = 257
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 7
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(132, 425)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 26)
        Me.btnCancel.TabIndex = 156
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(71, 425)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 26)
        Me.btnOK.TabIndex = 155
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnUP
        '
        Me.btnUP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUP.BackColor = System.Drawing.Color.Transparent
        Me.btnUP.BackgroundImage = CType(resources.GetObject("btnUP.BackgroundImage"), System.Drawing.Image)
        Me.btnUP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnUP.ForeColor = System.Drawing.SystemColors.Control
        Me.btnUP.Location = New System.Drawing.Point(4, 425)
        Me.btnUP.Name = "btnUP"
        Me.btnUP.Size = New System.Drawing.Size(28, 26)
        Me.btnUP.TabIndex = 153
        Me.btnUP.TabStop = False
        Me.btnUP.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackColor = System.Drawing.Color.Transparent
        Me.btnDown.BackgroundImage = CType(resources.GetObject("btnDown.BackgroundImage"), System.Drawing.Image)
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDown.ForeColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(38, 425)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(28, 26)
        Me.btnDown.TabIndex = 154
        Me.btnDown.TabStop = False
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvItemCost_CostGroup)
        Me.GroupBox1.Location = New System.Drawing.Point(2, 2)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(252, 418)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Cost Group"
        '
        'dgvItemCost_CostGroup
        '
        Me.dgvItemCost_CostGroup.AllowUserToResizeColumns = False
        Me.dgvItemCost_CostGroup.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemCost_CostGroup.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItemCost_CostGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemCost_CostGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemCost_CostGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCostGroup_GroupDesc, Me.ColCostGroup_OverheadAmount, Me.ColCostGroup_AllocationStyle, Me.ColCostGroup_RefUnit, Me.ColCostGroup_Remarks, Me.ColCostGroup_PrimaryRefUnit, Me.ColCostGroup_Identifier, Me.ColCostGroup_TotalQty})
        Me.dgvItemCost_CostGroup.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemCost_CostGroup.Location = New System.Drawing.Point(2, 15)
        Me.dgvItemCost_CostGroup.MultiSelect = False
        Me.dgvItemCost_CostGroup.Name = "dgvItemCost_CostGroup"
        Me.dgvItemCost_CostGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemCost_CostGroup.RowTemplate.Height = 24
        Me.dgvItemCost_CostGroup.Size = New System.Drawing.Size(248, 401)
        Me.dgvItemCost_CostGroup.StandardTab = True
        Me.dgvItemCost_CostGroup.TabIndex = 4
        '
        'ColCostGroup_GroupDesc
        '
        Me.ColCostGroup_GroupDesc.HeaderText = "Description"
        Me.ColCostGroup_GroupDesc.MaxInputLength = 250
        Me.ColCostGroup_GroupDesc.Name = "ColCostGroup_GroupDesc"
        Me.ColCostGroup_GroupDesc.Width = 85
        '
        'ColCostGroup_OverheadAmount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColCostGroup_OverheadAmount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColCostGroup_OverheadAmount.HeaderText = "Overhead Amount"
        Me.ColCostGroup_OverheadAmount.Name = "ColCostGroup_OverheadAmount"
        Me.ColCostGroup_OverheadAmount.Width = 118
        '
        'ColCostGroup_AllocationStyle
        '
        Me.ColCostGroup_AllocationStyle.HeaderText = "Allocation Style"
        Me.ColCostGroup_AllocationStyle.Items.AddRange(New Object() {"By Process Output", "By Process Work Hour"})
        Me.ColCostGroup_AllocationStyle.Name = "ColCostGroup_AllocationStyle"
        Me.ColCostGroup_AllocationStyle.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColCostGroup_AllocationStyle.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColCostGroup_AllocationStyle.Visible = False
        Me.ColCostGroup_AllocationStyle.Width = 129
        '
        'ColCostGroup_RefUnit
        '
        Me.ColCostGroup_RefUnit.HeaderText = "Ref. Unit"
        Me.ColCostGroup_RefUnit.MaxInputLength = 250
        Me.ColCostGroup_RefUnit.Name = "ColCostGroup_RefUnit"
        Me.ColCostGroup_RefUnit.Width = 74
        '
        'ColCostGroup_Remarks
        '
        Me.ColCostGroup_Remarks.HeaderText = "Remarks"
        Me.ColCostGroup_Remarks.MaxInputLength = 250
        Me.ColCostGroup_Remarks.Name = "ColCostGroup_Remarks"
        Me.ColCostGroup_Remarks.Width = 74
        '
        'ColCostGroup_PrimaryRefUnit
        '
        Me.ColCostGroup_PrimaryRefUnit.HeaderText = "Primary Ref Unit"
        Me.ColCostGroup_PrimaryRefUnit.Name = "ColCostGroup_PrimaryRefUnit"
        Me.ColCostGroup_PrimaryRefUnit.Visible = False
        Me.ColCostGroup_PrimaryRefUnit.Width = 136
        '
        'ColCostGroup_Identifier
        '
        Me.ColCostGroup_Identifier.HeaderText = "Identifier"
        Me.ColCostGroup_Identifier.Name = "ColCostGroup_Identifier"
        Me.ColCostGroup_Identifier.ReadOnly = True
        Me.ColCostGroup_Identifier.Visible = False
        Me.ColCostGroup_Identifier.Width = 87
        '
        'ColCostGroup_TotalQty
        '
        Me.ColCostGroup_TotalQty.HeaderText = "Total Qty"
        Me.ColCostGroup_TotalQty.Name = "ColCostGroup_TotalQty"
        Me.ColCostGroup_TotalQty.Visible = False
        Me.ColCostGroup_TotalQty.Width = 91
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Controls.Add(Me.TabPage6)
        Me.TabControl2.Controls.Add(Me.TabPage8)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(0, 0)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(403, 457)
        Me.TabControl2.TabIndex = 0
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.llbRefreshBalance)
        Me.TabPage5.Controls.Add(Me.BtnUpAccount)
        Me.TabPage5.Controls.Add(Me.BtnDownAccount)
        Me.TabPage5.Controls.Add(Me.dgvItemCost_CostAccountAllocation)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage5.Size = New System.Drawing.Size(395, 431)
        Me.TabPage5.TabIndex = 0
        Me.TabPage5.Text = "Account Allocation"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'llbRefreshBalance
        '
        Me.llbRefreshBalance.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbRefreshBalance.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbRefreshBalance.AutoSize = True
        Me.llbRefreshBalance.Location = New System.Drawing.Point(72, 407)
        Me.llbRefreshBalance.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.llbRefreshBalance.Name = "llbRefreshBalance"
        Me.llbRefreshBalance.Size = New System.Drawing.Size(86, 13)
        Me.llbRefreshBalance.TabIndex = 157
        Me.llbRefreshBalance.TabStop = True
        Me.llbRefreshBalance.Text = "Refresh Balance"
        Me.llbRefreshBalance.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'BtnUpAccount
        '
        Me.BtnUpAccount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnUpAccount.BackColor = System.Drawing.Color.Transparent
        Me.BtnUpAccount.BackgroundImage = CType(resources.GetObject("BtnUpAccount.BackgroundImage"), System.Drawing.Image)
        Me.BtnUpAccount.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnUpAccount.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnUpAccount.Location = New System.Drawing.Point(5, 402)
        Me.BtnUpAccount.Name = "BtnUpAccount"
        Me.BtnUpAccount.Size = New System.Drawing.Size(28, 26)
        Me.BtnUpAccount.TabIndex = 155
        Me.BtnUpAccount.TabStop = False
        Me.BtnUpAccount.UseVisualStyleBackColor = True
        '
        'BtnDownAccount
        '
        Me.BtnDownAccount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnDownAccount.BackColor = System.Drawing.Color.Transparent
        Me.BtnDownAccount.BackgroundImage = CType(resources.GetObject("BtnDownAccount.BackgroundImage"), System.Drawing.Image)
        Me.BtnDownAccount.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnDownAccount.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnDownAccount.Location = New System.Drawing.Point(39, 402)
        Me.BtnDownAccount.Name = "BtnDownAccount"
        Me.BtnDownAccount.Size = New System.Drawing.Size(28, 26)
        Me.BtnDownAccount.TabIndex = 156
        Me.BtnDownAccount.TabStop = False
        Me.BtnDownAccount.UseVisualStyleBackColor = True
        '
        'dgvItemCost_CostAccountAllocation
        '
        Me.dgvItemCost_CostAccountAllocation.AllowUserToResizeColumns = False
        Me.dgvItemCost_CostAccountAllocation.AllowUserToResizeRows = False
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemCost_CostAccountAllocation.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvItemCost_CostAccountAllocation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemCost_CostAccountAllocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemCost_CostAccountAllocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemCost_CostAccountAllocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccAll_AccountCode, Me.ColAccAll_AccountDesc, Me.ColAccAll_EndBalance, Me.ColAccAll_AllocatedPerct, Me.ColAccAll_AllocatedAmount, Me.ColAccAll_IgnoreNegative, Me.ColAccAll_Remarks, Me.ColAccAll_Identifier, Me.ColAccAll_PrimaryAkun})
        Me.dgvItemCost_CostAccountAllocation.Location = New System.Drawing.Point(2, 2)
        Me.dgvItemCost_CostAccountAllocation.MultiSelect = False
        Me.dgvItemCost_CostAccountAllocation.Name = "dgvItemCost_CostAccountAllocation"
        Me.dgvItemCost_CostAccountAllocation.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemCost_CostAccountAllocation.RowTemplate.Height = 24
        Me.dgvItemCost_CostAccountAllocation.Size = New System.Drawing.Size(392, 396)
        Me.dgvItemCost_CostAccountAllocation.StandardTab = True
        Me.dgvItemCost_CostAccountAllocation.TabIndex = 3
        '
        'ColAccAll_AccountCode
        '
        Me.ColAccAll_AccountCode.Frozen = True
        Me.ColAccAll_AccountCode.HeaderText = "Account"
        Me.ColAccAll_AccountCode.MaxInputLength = 250
        Me.ColAccAll_AccountCode.Name = "ColAccAll_AccountCode"
        Me.ColAccAll_AccountCode.Width = 72
        '
        'ColAccAll_AccountDesc
        '
        Me.ColAccAll_AccountDesc.HeaderText = "Description"
        Me.ColAccAll_AccountDesc.MaxInputLength = 250
        Me.ColAccAll_AccountDesc.Name = "ColAccAll_AccountDesc"
        Me.ColAccAll_AccountDesc.ReadOnly = True
        Me.ColAccAll_AccountDesc.Width = 85
        '
        'ColAccAll_EndBalance
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "0"
        Me.ColAccAll_EndBalance.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColAccAll_EndBalance.HeaderText = "Balance"
        Me.ColAccAll_EndBalance.MaxInputLength = 250
        Me.ColAccAll_EndBalance.Name = "ColAccAll_EndBalance"
        Me.ColAccAll_EndBalance.ReadOnly = True
        Me.ColAccAll_EndBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccAll_EndBalance.Width = 52
        '
        'ColAccAll_AllocatedPerct
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "0.00"
        Me.ColAccAll_AllocatedPerct.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColAccAll_AllocatedPerct.HeaderText = "Allocated %"
        Me.ColAccAll_AllocatedPerct.MaxInputLength = 25
        Me.ColAccAll_AllocatedPerct.Name = "ColAccAll_AllocatedPerct"
        Me.ColAccAll_AllocatedPerct.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccAll_AllocatedPerct.Width = 68
        '
        'ColAccAll_AllocatedAmount
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColAccAll_AllocatedAmount.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColAccAll_AllocatedAmount.HeaderText = "Allocated Amount"
        Me.ColAccAll_AllocatedAmount.MaxInputLength = 250
        Me.ColAccAll_AllocatedAmount.Name = "ColAccAll_AllocatedAmount"
        Me.ColAccAll_AllocatedAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccAll_AllocatedAmount.Width = 96
        '
        'ColAccAll_IgnoreNegative
        '
        Me.ColAccAll_IgnoreNegative.HeaderText = "Ignore Negative Balance"
        Me.ColAccAll_IgnoreNegative.Name = "ColAccAll_IgnoreNegative"
        Me.ColAccAll_IgnoreNegative.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAccAll_IgnoreNegative.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColAccAll_IgnoreNegative.Width = 150
        '
        'ColAccAll_Remarks
        '
        Me.ColAccAll_Remarks.HeaderText = "Remarks"
        Me.ColAccAll_Remarks.MaxInputLength = 150
        Me.ColAccAll_Remarks.Name = "ColAccAll_Remarks"
        Me.ColAccAll_Remarks.Width = 74
        '
        'ColAccAll_Identifier
        '
        Me.ColAccAll_Identifier.HeaderText = "Identifier"
        Me.ColAccAll_Identifier.Name = "ColAccAll_Identifier"
        Me.ColAccAll_Identifier.ReadOnly = True
        Me.ColAccAll_Identifier.Visible = False
        Me.ColAccAll_Identifier.Width = 87
        '
        'ColAccAll_PrimaryAkun
        '
        DataGridViewCellStyle7.NullValue = "-1"
        Me.ColAccAll_PrimaryAkun.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColAccAll_PrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColAccAll_PrimaryAkun.Name = "ColAccAll_PrimaryAkun"
        Me.ColAccAll_PrimaryAkun.ReadOnly = True
        Me.ColAccAll_PrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccAll_PrimaryAkun.Visible = False
        Me.ColAccAll_PrimaryAkun.Width = 94
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.BtnUpScope)
        Me.TabPage6.Controls.Add(Me.BtnDownScope)
        Me.TabPage6.Controls.Add(Me.dgvItemCost_Scope)
        Me.TabPage6.Location = New System.Drawing.Point(4, 22)
        Me.TabPage6.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage6.Size = New System.Drawing.Size(560, 464)
        Me.TabPage6.TabIndex = 1
        Me.TabPage6.Text = "Process"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'BtnUpScope
        '
        Me.BtnUpScope.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnUpScope.BackColor = System.Drawing.Color.Transparent
        Me.BtnUpScope.BackgroundImage = CType(resources.GetObject("BtnUpScope.BackgroundImage"), System.Drawing.Image)
        Me.BtnUpScope.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnUpScope.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnUpScope.Location = New System.Drawing.Point(5, 438)
        Me.BtnUpScope.Name = "BtnUpScope"
        Me.BtnUpScope.Size = New System.Drawing.Size(28, 26)
        Me.BtnUpScope.TabIndex = 155
        Me.BtnUpScope.TabStop = False
        Me.BtnUpScope.UseVisualStyleBackColor = True
        '
        'BtnDownScope
        '
        Me.BtnDownScope.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnDownScope.BackColor = System.Drawing.Color.Transparent
        Me.BtnDownScope.BackgroundImage = CType(resources.GetObject("BtnDownScope.BackgroundImage"), System.Drawing.Image)
        Me.BtnDownScope.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.BtnDownScope.ForeColor = System.Drawing.SystemColors.Control
        Me.BtnDownScope.Location = New System.Drawing.Point(39, 438)
        Me.BtnDownScope.Name = "BtnDownScope"
        Me.BtnDownScope.Size = New System.Drawing.Size(28, 26)
        Me.BtnDownScope.TabIndex = 156
        Me.BtnDownScope.TabStop = False
        Me.BtnDownScope.UseVisualStyleBackColor = True
        '
        'dgvItemCost_Scope
        '
        Me.dgvItemCost_Scope.AllowUserToResizeColumns = False
        Me.dgvItemCost_Scope.AllowUserToResizeRows = False
        Me.dgvItemCost_Scope.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemCost_Scope.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemCost_Scope.ColumnHeadersHeight = 25
        Me.dgvItemCost_Scope.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemCost_Scope.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAllScope_No, Me.ColAllScope_ProcessCode, Me.ColAllScope_ProcessDesc, Me.ColAllScope_Remarks, Me.ColAllScope_Identifier, Me.ColAllScope_PrimaryProcess})
        Me.dgvItemCost_Scope.Location = New System.Drawing.Point(2, 2)
        Me.dgvItemCost_Scope.MultiSelect = False
        Me.dgvItemCost_Scope.Name = "dgvItemCost_Scope"
        Me.dgvItemCost_Scope.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemCost_Scope.RowTemplate.Height = 24
        Me.dgvItemCost_Scope.Size = New System.Drawing.Size(557, 429)
        Me.dgvItemCost_Scope.TabIndex = 3
        '
        'ColAllScope_No
        '
        Me.ColAllScope_No.Frozen = True
        Me.ColAllScope_No.HeaderText = "No"
        Me.ColAllScope_No.Name = "ColAllScope_No"
        Me.ColAllScope_No.ReadOnly = True
        Me.ColAllScope_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAllScope_No.Visible = False
        Me.ColAllScope_No.Width = 27
        '
        'ColAllScope_ProcessCode
        '
        Me.ColAllScope_ProcessCode.Frozen = True
        Me.ColAllScope_ProcessCode.HeaderText = "Process Code"
        Me.ColAllScope_ProcessCode.MaxInputLength = 250
        Me.ColAllScope_ProcessCode.Name = "ColAllScope_ProcessCode"
        Me.ColAllScope_ProcessCode.Width = 98
        '
        'ColAllScope_ProcessDesc
        '
        Me.ColAllScope_ProcessDesc.Frozen = True
        Me.ColAllScope_ProcessDesc.HeaderText = "Description"
        Me.ColAllScope_ProcessDesc.Name = "ColAllScope_ProcessDesc"
        Me.ColAllScope_ProcessDesc.ReadOnly = True
        Me.ColAllScope_ProcessDesc.Width = 85
        '
        'ColAllScope_Remarks
        '
        Me.ColAllScope_Remarks.Frozen = True
        Me.ColAllScope_Remarks.HeaderText = "Remarks"
        Me.ColAllScope_Remarks.MaxInputLength = 250
        Me.ColAllScope_Remarks.Name = "ColAllScope_Remarks"
        Me.ColAllScope_Remarks.Width = 74
        '
        'ColAllScope_Identifier
        '
        Me.ColAllScope_Identifier.Frozen = True
        Me.ColAllScope_Identifier.HeaderText = "Identifier Scope"
        Me.ColAllScope_Identifier.Name = "ColAllScope_Identifier"
        Me.ColAllScope_Identifier.ReadOnly = True
        Me.ColAllScope_Identifier.Visible = False
        Me.ColAllScope_Identifier.Width = 106
        '
        'ColAllScope_PrimaryProcess
        '
        Me.ColAllScope_PrimaryProcess.HeaderText = "Primary Acuan"
        Me.ColAllScope_PrimaryProcess.Name = "ColAllScope_PrimaryProcess"
        Me.ColAllScope_PrimaryProcess.ReadOnly = True
        Me.ColAllScope_PrimaryProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAllScope_PrimaryProcess.Visible = False
        Me.ColAllScope_PrimaryProcess.Width = 81
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.rboGroup_Scope)
        Me.TabPage8.Controls.Add(Me.rboGroup_Account)
        Me.TabPage8.Controls.Add(Me.dgvItemCost_AllocationRecap1)
        Me.TabPage8.Controls.Add(Me.dgvItemCost_AllocationRecap2)
        Me.TabPage8.Location = New System.Drawing.Point(4, 22)
        Me.TabPage8.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TabPage8.Size = New System.Drawing.Size(560, 464)
        Me.TabPage8.TabIndex = 3
        Me.TabPage8.Text = "Allocation Recap"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'rboGroup_Scope
        '
        Me.rboGroup_Scope.AutoSize = True
        Me.rboGroup_Scope.Location = New System.Drawing.Point(204, 5)
        Me.rboGroup_Scope.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboGroup_Scope.Name = "rboGroup_Scope"
        Me.rboGroup_Scope.Size = New System.Drawing.Size(167, 17)
        Me.rboGroup_Scope.TabIndex = 1
        Me.rboGroup_Scope.Text = "Cost Group - Allocation Scope"
        Me.rboGroup_Scope.UseVisualStyleBackColor = True
        '
        'rboGroup_Account
        '
        Me.rboGroup_Account.AutoSize = True
        Me.rboGroup_Account.Checked = True
        Me.rboGroup_Account.Location = New System.Drawing.Point(4, 5)
        Me.rboGroup_Account.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rboGroup_Account.Name = "rboGroup_Account"
        Me.rboGroup_Account.Size = New System.Drawing.Size(176, 17)
        Me.rboGroup_Account.TabIndex = 0
        Me.rboGroup_Account.TabStop = True
        Me.rboGroup_Account.Text = "Cost Group - Account Allocation"
        Me.rboGroup_Account.UseVisualStyleBackColor = True
        '
        'dgvItemCost_AllocationRecap1
        '
        Me.dgvItemCost_AllocationRecap1.AllowUserToAddRows = False
        Me.dgvItemCost_AllocationRecap1.AllowUserToDeleteRows = False
        Me.dgvItemCost_AllocationRecap1.AllowUserToResizeColumns = False
        Me.dgvItemCost_AllocationRecap1.AllowUserToResizeRows = False
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItemCost_AllocationRecap1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle8
        Me.dgvItemCost_AllocationRecap1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemCost_AllocationRecap1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemCost_AllocationRecap1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemCost_AllocationRecap1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAlloRecap1_CostGroup, Me.ColAlloRecap1_AccCode, Me.ColAlloRecap1_AccDesc, Me.ColAlloRecap1_Balance, Me.ColAlloRecap1_AllocatedPerct, Me.ColAlloRecap1_AllocatedAmount, Me.ColAlloRecap1_IngoreNegativeBalance, Me.ColAlloRecap1_Identifier, Me.ColAlloRecap1_PrimaryAkun})
        Me.dgvItemCost_AllocationRecap1.Location = New System.Drawing.Point(2, 33)
        Me.dgvItemCost_AllocationRecap1.Name = "dgvItemCost_AllocationRecap1"
        Me.dgvItemCost_AllocationRecap1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemCost_AllocationRecap1.RowTemplate.Height = 24
        Me.dgvItemCost_AllocationRecap1.Size = New System.Drawing.Size(557, 431)
        Me.dgvItemCost_AllocationRecap1.StandardTab = True
        Me.dgvItemCost_AllocationRecap1.TabIndex = 2
        '
        'ColAlloRecap1_CostGroup
        '
        Me.ColAlloRecap1_CostGroup.HeaderText = "Cost Group"
        Me.ColAlloRecap1_CostGroup.Name = "ColAlloRecap1_CostGroup"
        Me.ColAlloRecap1_CostGroup.Width = 85
        '
        'ColAlloRecap1_AccCode
        '
        Me.ColAlloRecap1_AccCode.HeaderText = "Account"
        Me.ColAlloRecap1_AccCode.MaxInputLength = 250
        Me.ColAlloRecap1_AccCode.Name = "ColAlloRecap1_AccCode"
        Me.ColAlloRecap1_AccCode.Width = 72
        '
        'ColAlloRecap1_AccDesc
        '
        Me.ColAlloRecap1_AccDesc.HeaderText = "Description"
        Me.ColAlloRecap1_AccDesc.MaxInputLength = 250
        Me.ColAlloRecap1_AccDesc.Name = "ColAlloRecap1_AccDesc"
        Me.ColAlloRecap1_AccDesc.ReadOnly = True
        Me.ColAlloRecap1_AccDesc.Width = 85
        '
        'ColAlloRecap1_Balance
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle9.NullValue = "0"
        Me.ColAlloRecap1_Balance.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColAlloRecap1_Balance.HeaderText = "Balance"
        Me.ColAlloRecap1_Balance.MaxInputLength = 250
        Me.ColAlloRecap1_Balance.Name = "ColAlloRecap1_Balance"
        Me.ColAlloRecap1_Balance.ReadOnly = True
        Me.ColAlloRecap1_Balance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAlloRecap1_Balance.Width = 52
        '
        'ColAlloRecap1_AllocatedPerct
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle10.NullValue = "0.00"
        Me.ColAlloRecap1_AllocatedPerct.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColAlloRecap1_AllocatedPerct.HeaderText = "Allocated %"
        Me.ColAlloRecap1_AllocatedPerct.MaxInputLength = 25
        Me.ColAlloRecap1_AllocatedPerct.Name = "ColAlloRecap1_AllocatedPerct"
        Me.ColAlloRecap1_AllocatedPerct.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAlloRecap1_AllocatedPerct.Width = 68
        '
        'ColAlloRecap1_AllocatedAmount
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle11.NullValue = "0"
        Me.ColAlloRecap1_AllocatedAmount.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColAlloRecap1_AllocatedAmount.HeaderText = "Allocated Amount"
        Me.ColAlloRecap1_AllocatedAmount.MaxInputLength = 250
        Me.ColAlloRecap1_AllocatedAmount.Name = "ColAlloRecap1_AllocatedAmount"
        Me.ColAlloRecap1_AllocatedAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAlloRecap1_AllocatedAmount.Width = 96
        '
        'ColAlloRecap1_IngoreNegativeBalance
        '
        Me.ColAlloRecap1_IngoreNegativeBalance.HeaderText = "Ignore Negative Balance"
        Me.ColAlloRecap1_IngoreNegativeBalance.Name = "ColAlloRecap1_IngoreNegativeBalance"
        Me.ColAlloRecap1_IngoreNegativeBalance.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAlloRecap1_IngoreNegativeBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColAlloRecap1_IngoreNegativeBalance.Width = 150
        '
        'ColAlloRecap1_Identifier
        '
        Me.ColAlloRecap1_Identifier.HeaderText = "Identifier"
        Me.ColAlloRecap1_Identifier.Name = "ColAlloRecap1_Identifier"
        Me.ColAlloRecap1_Identifier.ReadOnly = True
        Me.ColAlloRecap1_Identifier.Visible = False
        Me.ColAlloRecap1_Identifier.Width = 87
        '
        'ColAlloRecap1_PrimaryAkun
        '
        DataGridViewCellStyle12.NullValue = "-1"
        Me.ColAlloRecap1_PrimaryAkun.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColAlloRecap1_PrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColAlloRecap1_PrimaryAkun.Name = "ColAlloRecap1_PrimaryAkun"
        Me.ColAlloRecap1_PrimaryAkun.ReadOnly = True
        Me.ColAlloRecap1_PrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAlloRecap1_PrimaryAkun.Visible = False
        Me.ColAlloRecap1_PrimaryAkun.Width = 94
        '
        'dgvItemCost_AllocationRecap2
        '
        Me.dgvItemCost_AllocationRecap2.AllowUserToAddRows = False
        Me.dgvItemCost_AllocationRecap2.AllowUserToDeleteRows = False
        Me.dgvItemCost_AllocationRecap2.AllowUserToResizeColumns = False
        Me.dgvItemCost_AllocationRecap2.AllowUserToResizeRows = False
        Me.dgvItemCost_AllocationRecap2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItemCost_AllocationRecap2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemCost_AllocationRecap2.ColumnHeadersHeight = 25
        Me.dgvItemCost_AllocationRecap2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemCost_AllocationRecap2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAlloRecap2_CostGroup, Me.ColAlloRecap2_ProcessCode, Me.ColAlloRecap2_ProcessDesc, Me.ColAlloRecap2_Identifier, Me.ColAlloRecap2_PrimaryProcess})
        Me.dgvItemCost_AllocationRecap2.Location = New System.Drawing.Point(2, 33)
        Me.dgvItemCost_AllocationRecap2.MultiSelect = False
        Me.dgvItemCost_AllocationRecap2.Name = "dgvItemCost_AllocationRecap2"
        Me.dgvItemCost_AllocationRecap2.ReadOnly = True
        Me.dgvItemCost_AllocationRecap2.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemCost_AllocationRecap2.RowTemplate.Height = 24
        Me.dgvItemCost_AllocationRecap2.Size = New System.Drawing.Size(557, 431)
        Me.dgvItemCost_AllocationRecap2.TabIndex = 4
        '
        'ColAlloRecap2_CostGroup
        '
        Me.ColAlloRecap2_CostGroup.Frozen = True
        Me.ColAlloRecap2_CostGroup.HeaderText = "Cost Group"
        Me.ColAlloRecap2_CostGroup.Name = "ColAlloRecap2_CostGroup"
        Me.ColAlloRecap2_CostGroup.ReadOnly = True
        Me.ColAlloRecap2_CostGroup.Width = 85
        '
        'ColAlloRecap2_ProcessCode
        '
        Me.ColAlloRecap2_ProcessCode.Frozen = True
        Me.ColAlloRecap2_ProcessCode.HeaderText = "Process Code"
        Me.ColAlloRecap2_ProcessCode.MaxInputLength = 250
        Me.ColAlloRecap2_ProcessCode.Name = "ColAlloRecap2_ProcessCode"
        Me.ColAlloRecap2_ProcessCode.ReadOnly = True
        Me.ColAlloRecap2_ProcessCode.Width = 98
        '
        'ColAlloRecap2_ProcessDesc
        '
        Me.ColAlloRecap2_ProcessDesc.Frozen = True
        Me.ColAlloRecap2_ProcessDesc.HeaderText = "Description"
        Me.ColAlloRecap2_ProcessDesc.Name = "ColAlloRecap2_ProcessDesc"
        Me.ColAlloRecap2_ProcessDesc.ReadOnly = True
        Me.ColAlloRecap2_ProcessDesc.Width = 85
        '
        'ColAlloRecap2_Identifier
        '
        Me.ColAlloRecap2_Identifier.Frozen = True
        Me.ColAlloRecap2_Identifier.HeaderText = "Identifier Scope"
        Me.ColAlloRecap2_Identifier.Name = "ColAlloRecap2_Identifier"
        Me.ColAlloRecap2_Identifier.ReadOnly = True
        Me.ColAlloRecap2_Identifier.Visible = False
        Me.ColAlloRecap2_Identifier.Width = 106
        '
        'ColAlloRecap2_PrimaryProcess
        '
        Me.ColAlloRecap2_PrimaryProcess.HeaderText = "Primary Acuan"
        Me.ColAlloRecap2_PrimaryProcess.Name = "ColAlloRecap2_PrimaryProcess"
        Me.ColAlloRecap2_PrimaryProcess.ReadOnly = True
        Me.ColAlloRecap2_PrimaryProcess.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAlloRecap2_PrimaryProcess.Visible = False
        Me.ColAlloRecap2_PrimaryProcess.Width = 81
        '
        'BgwDelCostGroup
        '
        '
        'BgwDelCostScope
        '
        '
        'BgwDelCostAcc
        '
        '
        'FrmTransaksiMonthlyClosing2_Overhead2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(663, 457)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(679, 495)
        Me.Name = "FrmTransaksiMonthlyClosing2_Overhead2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overhead Cost Allocation"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.dgvItemCost_CostGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        CType(Me.dgvItemCost_CostAccountAllocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        CType(Me.dgvItemCost_Scope, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage8.ResumeLayout(False)
        Me.TabPage8.PerformLayout()
        CType(Me.dgvItemCost_AllocationRecap1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemCost_AllocationRecap2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnUP As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents dgvItemCost_CostGroup As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents llbRefreshBalance As System.Windows.Forms.LinkLabel
    Friend WithEvents BtnUpAccount As System.Windows.Forms.Button
    Friend WithEvents BtnDownAccount As System.Windows.Forms.Button
    Public WithEvents dgvItemCost_CostAccountAllocation As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents BtnUpScope As System.Windows.Forms.Button
    Friend WithEvents BtnDownScope As System.Windows.Forms.Button
    Friend WithEvents dgvItemCost_Scope As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents rboGroup_Scope As System.Windows.Forms.RadioButton
    Friend WithEvents rboGroup_Account As System.Windows.Forms.RadioButton
    Public WithEvents dgvItemCost_AllocationRecap1 As System.Windows.Forms.DataGridView
    Friend WithEvents dgvItemCost_AllocationRecap2 As System.Windows.Forms.DataGridView
    Friend WithEvents BgwDelCostGroup As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwDelCostScope As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwDelCostAcc As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColCostGroup_GroupDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostGroup_OverheadAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostGroup_AllocationStyle As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColCostGroup_RefUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostGroup_PrimaryRefUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostGroup_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCostGroup_TotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_AccountCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_AccountDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_EndBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_AllocatedPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_AllocatedAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_IgnoreNegative As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAccAll_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccAll_PrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllScope_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllScope_ProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllScope_ProcessDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllScope_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllScope_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAllScope_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_CostGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_AccCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_AccDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_Balance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_AllocatedPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_AllocatedAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_IngoreNegativeBalance As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAlloRecap1_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap1_PrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap2_CostGroup As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap2_ProcessCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap2_ProcessDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap2_Identifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlloRecap2_PrimaryProcess As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
