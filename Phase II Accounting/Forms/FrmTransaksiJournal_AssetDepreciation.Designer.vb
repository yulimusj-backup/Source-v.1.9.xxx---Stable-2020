<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_AssetDepreciation
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.gboAsset_Detail = New System.Windows.Forms.GroupBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvAsset_ItemAsset = New System.Windows.Forms.DataGridView
        Me.ColItemAssetNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_AssetCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_AssetDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_DepreciationPeriod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_DepreciationMethod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_LastBookValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_RevaluationValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_DisposalValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_SalvageValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_DepreciatedValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_EndValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_PrimaryAkumDepAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColItemAsset_PrimaryDepExpenseAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAsset = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvAsset_Item = New System.Windows.Forms.DataGridView
        Me.ColAsset_Account = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_DescAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_Debet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_Kredit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_ExchangeRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_OriginalCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_Debet_OriginalCurr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_Credit_OriginalCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_Dept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_Job = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_PrimaryDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_PrimaryJob = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_PrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAsset_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1_Asset = New System.Windows.Forms.SplitContainer
        Me.txtAsset_ReadOnly = New System.Windows.Forms.TextBox
        Me.TxtAsset_PrimaryJournal = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtRefNo = New System.Windows.Forms.TextBox
        Me.dtpAsset_Date = New System.Windows.Forms.DateTimePicker
        Me.TxtAsset_Remarks = New System.Windows.Forms.TextBox
        Me.TxtAsset_Unbalance = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnAsset_Cancel = New System.Windows.Forms.Button
        Me.btnAsset_OK = New System.Windows.Forms.Button
        Me.StatusStripAsset_1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnAsset_ViewJurnal = New System.Windows.Forms.Button
        Me.btnAsset_Delete = New System.Windows.Forms.Button
        Me.btnAsset_Print = New System.Windows.Forms.Button
        Me.btnAsset_Exit = New System.Windows.Forms.Button
        Me.btnAsset_Edit = New System.Windows.Forms.Button
        Me.btnAsset_Find = New System.Windows.Forms.Button
        Me.btnAsset_New = New System.Windows.Forms.Button
        Me.gboAsset_Detail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvAsset_ItemAsset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvAsset_Item, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1_Asset.Panel1.SuspendLayout()
        Me.SplitContainer1_Asset.Panel2.SuspendLayout()
        Me.SplitContainer1_Asset.SuspendLayout()
        Me.StatusStripAsset_1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboAsset_Detail
        '
        Me.gboAsset_Detail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboAsset_Detail.Controls.Add(Me.TabControl1)
        Me.gboAsset_Detail.Controls.Add(Me.SplitContainer1_Asset)
        Me.gboAsset_Detail.Controls.Add(Me.btnAsset_Cancel)
        Me.gboAsset_Detail.Controls.Add(Me.btnAsset_OK)
        Me.gboAsset_Detail.Location = New System.Drawing.Point(12, 3)
        Me.gboAsset_Detail.Name = "gboAsset_Detail"
        Me.gboAsset_Detail.Size = New System.Drawing.Size(860, 573)
        Me.gboAsset_Detail.TabIndex = 40
        Me.gboAsset_Detail.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(6, 69)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(848, 466)
        Me.TabControl1.TabIndex = 7
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvAsset_ItemAsset)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(840, 440)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Asset Details"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvAsset_ItemAsset
        '
        Me.dgvAsset_ItemAsset.AllowUserToAddRows = False
        Me.dgvAsset_ItemAsset.AllowUserToDeleteRows = False
        Me.dgvAsset_ItemAsset.AllowUserToResizeColumns = False
        Me.dgvAsset_ItemAsset.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvAsset_ItemAsset.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvAsset_ItemAsset.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAsset_ItemAsset.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColItemAssetNo, Me.ColItemAsset_AssetCode, Me.ColItemAsset_AssetDesc, Me.ColItemAsset_DepreciationPeriod, Me.ColItemAsset_DepreciationMethod, Me.ColItemAsset_LastBookValue, Me.ColItemAsset_RevaluationValue, Me.ColItemAsset_DisposalValue, Me.ColItemAsset_SalvageValue, Me.ColItemAsset_DepreciatedValue, Me.ColItemAsset_EndValue, Me.ColItemAsset_Remarks, Me.ColItemAsset_PrimaryAkumDepAccount, Me.ColItemAsset_PrimaryDepExpenseAccount, Me.ColPrimaryAsset})
        Me.dgvAsset_ItemAsset.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAsset_ItemAsset.Location = New System.Drawing.Point(3, 3)
        Me.dgvAsset_ItemAsset.MultiSelect = False
        Me.dgvAsset_ItemAsset.Name = "dgvAsset_ItemAsset"
        Me.dgvAsset_ItemAsset.ReadOnly = True
        Me.dgvAsset_ItemAsset.Size = New System.Drawing.Size(834, 434)
        Me.dgvAsset_ItemAsset.StandardTab = True
        Me.dgvAsset_ItemAsset.TabIndex = 2
        '
        'ColItemAssetNo
        '
        Me.ColItemAssetNo.HeaderText = "No"
        Me.ColItemAssetNo.Name = "ColItemAssetNo"
        Me.ColItemAssetNo.ReadOnly = True
        Me.ColItemAssetNo.Width = 46
        '
        'ColItemAsset_AssetCode
        '
        Me.ColItemAsset_AssetCode.HeaderText = "Asset Code"
        Me.ColItemAsset_AssetCode.Name = "ColItemAsset_AssetCode"
        Me.ColItemAsset_AssetCode.ReadOnly = True
        Me.ColItemAsset_AssetCode.Width = 86
        '
        'ColItemAsset_AssetDesc
        '
        Me.ColItemAsset_AssetDesc.HeaderText = "Description"
        Me.ColItemAsset_AssetDesc.Name = "ColItemAsset_AssetDesc"
        Me.ColItemAsset_AssetDesc.ReadOnly = True
        Me.ColItemAsset_AssetDesc.Width = 85
        '
        'ColItemAsset_DepreciationPeriod
        '
        Me.ColItemAsset_DepreciationPeriod.HeaderText = "Depreciation Period"
        Me.ColItemAsset_DepreciationPeriod.Name = "ColItemAsset_DepreciationPeriod"
        Me.ColItemAsset_DepreciationPeriod.ReadOnly = True
        Me.ColItemAsset_DepreciationPeriod.Width = 125
        '
        'ColItemAsset_DepreciationMethod
        '
        Me.ColItemAsset_DepreciationMethod.HeaderText = "Depreciation Method"
        Me.ColItemAsset_DepreciationMethod.Name = "ColItemAsset_DepreciationMethod"
        Me.ColItemAsset_DepreciationMethod.ReadOnly = True
        Me.ColItemAsset_DepreciationMethod.Width = 131
        '
        'ColItemAsset_LastBookValue
        '
        Me.ColItemAsset_LastBookValue.HeaderText = "Last Book Value"
        Me.ColItemAsset_LastBookValue.Name = "ColItemAsset_LastBookValue"
        Me.ColItemAsset_LastBookValue.ReadOnly = True
        Me.ColItemAsset_LastBookValue.Width = 110
        '
        'ColItemAsset_RevaluationValue
        '
        Me.ColItemAsset_RevaluationValue.HeaderText = "Revaluation"
        Me.ColItemAsset_RevaluationValue.Name = "ColItemAsset_RevaluationValue"
        Me.ColItemAsset_RevaluationValue.ReadOnly = True
        Me.ColItemAsset_RevaluationValue.Width = 89
        '
        'ColItemAsset_DisposalValue
        '
        Me.ColItemAsset_DisposalValue.HeaderText = "Partial Disposal"
        Me.ColItemAsset_DisposalValue.Name = "ColItemAsset_DisposalValue"
        Me.ColItemAsset_DisposalValue.ReadOnly = True
        Me.ColItemAsset_DisposalValue.Width = 104
        '
        'ColItemAsset_SalvageValue
        '
        Me.ColItemAsset_SalvageValue.HeaderText = "Salvage Value"
        Me.ColItemAsset_SalvageValue.Name = "ColItemAsset_SalvageValue"
        Me.ColItemAsset_SalvageValue.ReadOnly = True
        Me.ColItemAsset_SalvageValue.Width = 101
        '
        'ColItemAsset_DepreciatedValue
        '
        Me.ColItemAsset_DepreciatedValue.HeaderText = "Depreciated Value"
        Me.ColItemAsset_DepreciatedValue.Name = "ColItemAsset_DepreciatedValue"
        Me.ColItemAsset_DepreciatedValue.ReadOnly = True
        Me.ColItemAsset_DepreciatedValue.Width = 120
        '
        'ColItemAsset_EndValue
        '
        Me.ColItemAsset_EndValue.HeaderText = "End Value"
        Me.ColItemAsset_EndValue.Name = "ColItemAsset_EndValue"
        Me.ColItemAsset_EndValue.ReadOnly = True
        Me.ColItemAsset_EndValue.Width = 81
        '
        'ColItemAsset_Remarks
        '
        Me.ColItemAsset_Remarks.HeaderText = "Remarks"
        Me.ColItemAsset_Remarks.Name = "ColItemAsset_Remarks"
        Me.ColItemAsset_Remarks.ReadOnly = True
        Me.ColItemAsset_Remarks.Width = 74
        '
        'ColItemAsset_PrimaryAkumDepAccount
        '
        Me.ColItemAsset_PrimaryAkumDepAccount.HeaderText = "Primary Akum Dep"
        Me.ColItemAsset_PrimaryAkumDepAccount.Name = "ColItemAsset_PrimaryAkumDepAccount"
        Me.ColItemAsset_PrimaryAkumDepAccount.ReadOnly = True
        Me.ColItemAsset_PrimaryAkumDepAccount.Visible = False
        Me.ColItemAsset_PrimaryAkumDepAccount.Width = 119
        '
        'ColItemAsset_PrimaryDepExpenseAccount
        '
        Me.ColItemAsset_PrimaryDepExpenseAccount.HeaderText = "Primary Dep Expense Account"
        Me.ColItemAsset_PrimaryDepExpenseAccount.Name = "ColItemAsset_PrimaryDepExpenseAccount"
        Me.ColItemAsset_PrimaryDepExpenseAccount.ReadOnly = True
        Me.ColItemAsset_PrimaryDepExpenseAccount.Visible = False
        Me.ColItemAsset_PrimaryDepExpenseAccount.Width = 176
        '
        'ColPrimaryAsset
        '
        Me.ColPrimaryAsset.HeaderText = "Primary Asset"
        Me.ColPrimaryAsset.Name = "ColPrimaryAsset"
        Me.ColPrimaryAsset.ReadOnly = True
        Me.ColPrimaryAsset.Visible = False
        Me.ColPrimaryAsset.Width = 95
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvAsset_Item)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(840, 440)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Account Details"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvAsset_Item
        '
        Me.dgvAsset_Item.AllowUserToAddRows = False
        Me.dgvAsset_Item.AllowUserToDeleteRows = False
        Me.dgvAsset_Item.AllowUserToResizeColumns = False
        Me.dgvAsset_Item.AllowUserToResizeRows = False
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvAsset_Item.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvAsset_Item.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAsset_Item.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAsset_Account, Me.ColAsset_DescAccount, Me.ColAsset_Debet, Me.ColAsset_Kredit, Me.ColAsset_ExchangeRate, Me.ColAsset_OriginalCurrency, Me.ColAsset_Debet_OriginalCurr, Me.ColAsset_Credit_OriginalCurrency, Me.ColAsset_Remarks, Me.ColAsset_Dept, Me.ColAsset_Job, Me.ColAsset_PrimaryDept, Me.ColAsset_PrimaryJob, Me.ColAsset_PrimaryAkun, Me.ColAsset_PrimaryCurrency})
        Me.dgvAsset_Item.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAsset_Item.Location = New System.Drawing.Point(3, 3)
        Me.dgvAsset_Item.MultiSelect = False
        Me.dgvAsset_Item.Name = "dgvAsset_Item"
        Me.dgvAsset_Item.Size = New System.Drawing.Size(834, 434)
        Me.dgvAsset_Item.StandardTab = True
        Me.dgvAsset_Item.TabIndex = 1
        '
        'ColAsset_Account
        '
        Me.ColAsset_Account.HeaderText = "Account"
        Me.ColAsset_Account.MaxInputLength = 50
        Me.ColAsset_Account.Name = "ColAsset_Account"
        Me.ColAsset_Account.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_Account.Width = 53
        '
        'ColAsset_DescAccount
        '
        Me.ColAsset_DescAccount.HeaderText = "Description"
        Me.ColAsset_DescAccount.MaxInputLength = 150
        Me.ColAsset_DescAccount.Name = "ColAsset_DescAccount"
        Me.ColAsset_DescAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_DescAccount.Width = 66
        '
        'ColAsset_Debet
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.NullValue = "0"
        Me.ColAsset_Debet.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColAsset_Debet.HeaderText = "Debet"
        Me.ColAsset_Debet.MaxInputLength = 25
        Me.ColAsset_Debet.Name = "ColAsset_Debet"
        Me.ColAsset_Debet.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_Debet.Width = 42
        '
        'ColAsset_Kredit
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle4.NullValue = "0"
        Me.ColAsset_Kredit.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColAsset_Kredit.HeaderText = "Credit"
        Me.ColAsset_Kredit.MaxInputLength = 25
        Me.ColAsset_Kredit.Name = "ColAsset_Kredit"
        Me.ColAsset_Kredit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_Kredit.Width = 40
        '
        'ColAsset_ExchangeRate
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.NullValue = "1"
        Me.ColAsset_ExchangeRate.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColAsset_ExchangeRate.HeaderText = "Exchange Rate"
        Me.ColAsset_ExchangeRate.Name = "ColAsset_ExchangeRate"
        Me.ColAsset_ExchangeRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_ExchangeRate.Visible = False
        Me.ColAsset_ExchangeRate.Width = 87
        '
        'ColAsset_OriginalCurrency
        '
        Me.ColAsset_OriginalCurrency.HeaderText = "Currency"
        Me.ColAsset_OriginalCurrency.Name = "ColAsset_OriginalCurrency"
        Me.ColAsset_OriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_OriginalCurrency.Visible = False
        Me.ColAsset_OriginalCurrency.Width = 55
        '
        'ColAsset_Debet_OriginalCurr
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle6.NullValue = "0"
        Me.ColAsset_Debet_OriginalCurr.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColAsset_Debet_OriginalCurr.HeaderText = "Debet"
        Me.ColAsset_Debet_OriginalCurr.Name = "ColAsset_Debet_OriginalCurr"
        Me.ColAsset_Debet_OriginalCurr.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_Debet_OriginalCurr.Visible = False
        Me.ColAsset_Debet_OriginalCurr.Width = 42
        '
        'ColAsset_Credit_OriginalCurrency
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.NullValue = "0"
        Me.ColAsset_Credit_OriginalCurrency.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColAsset_Credit_OriginalCurrency.HeaderText = "Credit"
        Me.ColAsset_Credit_OriginalCurrency.Name = "ColAsset_Credit_OriginalCurrency"
        Me.ColAsset_Credit_OriginalCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_Credit_OriginalCurrency.Visible = False
        Me.ColAsset_Credit_OriginalCurrency.Width = 40
        '
        'ColAsset_Remarks
        '
        Me.ColAsset_Remarks.HeaderText = "Remarks"
        Me.ColAsset_Remarks.MaxInputLength = 150
        Me.ColAsset_Remarks.Name = "ColAsset_Remarks"
        Me.ColAsset_Remarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_Remarks.Width = 55
        '
        'ColAsset_Dept
        '
        Me.ColAsset_Dept.HeaderText = "Departement"
        Me.ColAsset_Dept.Name = "ColAsset_Dept"
        Me.ColAsset_Dept.Width = 93
        '
        'ColAsset_Job
        '
        Me.ColAsset_Job.HeaderText = "Job"
        Me.ColAsset_Job.Name = "ColAsset_Job"
        Me.ColAsset_Job.Width = 49
        '
        'ColAsset_PrimaryDept
        '
        Me.ColAsset_PrimaryDept.HeaderText = "Primary Dept"
        Me.ColAsset_PrimaryDept.Name = "ColAsset_PrimaryDept"
        Me.ColAsset_PrimaryDept.Visible = False
        Me.ColAsset_PrimaryDept.Width = 92
        '
        'ColAsset_PrimaryJob
        '
        Me.ColAsset_PrimaryJob.HeaderText = "Primary Job"
        Me.ColAsset_PrimaryJob.Name = "ColAsset_PrimaryJob"
        Me.ColAsset_PrimaryJob.Visible = False
        Me.ColAsset_PrimaryJob.Width = 86
        '
        'ColAsset_PrimaryAkun
        '
        Me.ColAsset_PrimaryAkun.HeaderText = "PrimaryAkun"
        Me.ColAsset_PrimaryAkun.Name = "ColAsset_PrimaryAkun"
        Me.ColAsset_PrimaryAkun.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAsset_PrimaryAkun.Visible = False
        Me.ColAsset_PrimaryAkun.Width = 72
        '
        'ColAsset_PrimaryCurrency
        '
        Me.ColAsset_PrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColAsset_PrimaryCurrency.Name = "ColAsset_PrimaryCurrency"
        Me.ColAsset_PrimaryCurrency.Visible = False
        Me.ColAsset_PrimaryCurrency.Width = 111
        '
        'SplitContainer1_Asset
        '
        Me.SplitContainer1_Asset.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1_Asset.Location = New System.Drawing.Point(6, 8)
        Me.SplitContainer1_Asset.Name = "SplitContainer1_Asset"
        '
        'SplitContainer1_Asset.Panel1
        '
        Me.SplitContainer1_Asset.Panel1.Controls.Add(Me.txtAsset_ReadOnly)
        Me.SplitContainer1_Asset.Panel1.Controls.Add(Me.TxtAsset_PrimaryJournal)
        Me.SplitContainer1_Asset.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1_Asset.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1_Asset.Panel1.Controls.Add(Me.TxtRefNo)
        Me.SplitContainer1_Asset.Panel1.Controls.Add(Me.dtpAsset_Date)
        '
        'SplitContainer1_Asset.Panel2
        '
        Me.SplitContainer1_Asset.Panel2.Controls.Add(Me.TxtAsset_Remarks)
        Me.SplitContainer1_Asset.Panel2.Controls.Add(Me.TxtAsset_Unbalance)
        Me.SplitContainer1_Asset.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1_Asset.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1_Asset.Size = New System.Drawing.Size(843, 55)
        Me.SplitContainer1_Asset.SplitterDistance = 421
        Me.SplitContainer1_Asset.TabIndex = 0
        Me.SplitContainer1_Asset.TabStop = False
        '
        'txtAsset_ReadOnly
        '
        Me.txtAsset_ReadOnly.Location = New System.Drawing.Point(75, 3)
        Me.txtAsset_ReadOnly.Name = "txtAsset_ReadOnly"
        Me.txtAsset_ReadOnly.ReadOnly = True
        Me.txtAsset_ReadOnly.Size = New System.Drawing.Size(23, 20)
        Me.txtAsset_ReadOnly.TabIndex = 211
        Me.txtAsset_ReadOnly.TabStop = False
        Me.txtAsset_ReadOnly.Visible = False
        '
        'TxtAsset_PrimaryJournal
        '
        Me.TxtAsset_PrimaryJournal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAsset_PrimaryJournal.Location = New System.Drawing.Point(69, 26)
        Me.TxtAsset_PrimaryJournal.Name = "TxtAsset_PrimaryJournal"
        Me.TxtAsset_PrimaryJournal.Size = New System.Drawing.Size(31, 20)
        Me.TxtAsset_PrimaryJournal.TabIndex = 163
        Me.TxtAsset_PrimaryJournal.TabStop = False
        Me.TxtAsset_PrimaryJournal.Text = "-1"
        Me.TxtAsset_PrimaryJournal.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Document No"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'TxtRefNo
        '
        Me.TxtRefNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRefNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtRefNo.Location = New System.Drawing.Point(104, 3)
        Me.TxtRefNo.MaxLength = 50
        Me.TxtRefNo.Name = "TxtRefNo"
        Me.TxtRefNo.ReadOnly = True
        Me.TxtRefNo.Size = New System.Drawing.Size(314, 20)
        Me.TxtRefNo.TabIndex = 0
        Me.TxtRefNo.TabStop = False
        '
        'dtpAsset_Date
        '
        Me.dtpAsset_Date.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpAsset_Date.CustomFormat = "dd MMM yyyy"
        Me.dtpAsset_Date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAsset_Date.Location = New System.Drawing.Point(104, 28)
        Me.dtpAsset_Date.Name = "dtpAsset_Date"
        Me.dtpAsset_Date.Size = New System.Drawing.Size(314, 20)
        Me.dtpAsset_Date.TabIndex = 1
        '
        'TxtAsset_Remarks
        '
        Me.TxtAsset_Remarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAsset_Remarks.Location = New System.Drawing.Point(106, 5)
        Me.TxtAsset_Remarks.MaxLength = 150
        Me.TxtAsset_Remarks.Name = "TxtAsset_Remarks"
        Me.TxtAsset_Remarks.Size = New System.Drawing.Size(309, 20)
        Me.TxtAsset_Remarks.TabIndex = 0
        '
        'TxtAsset_Unbalance
        '
        Me.TxtAsset_Unbalance.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAsset_Unbalance.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtAsset_Unbalance.Location = New System.Drawing.Point(106, 31)
        Me.TxtAsset_Unbalance.MaxLength = 50
        Me.TxtAsset_Unbalance.Name = "TxtAsset_Unbalance"
        Me.TxtAsset_Unbalance.ReadOnly = True
        Me.TxtAsset_Unbalance.Size = New System.Drawing.Size(309, 20)
        Me.TxtAsset_Unbalance.TabIndex = 1
        Me.TxtAsset_Unbalance.TabStop = False
        Me.TxtAsset_Unbalance.Text = "0"
        Me.TxtAsset_Unbalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(60, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Description"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 164
        Me.Label2.Text = "Unbalance"
        '
        'btnAsset_Cancel
        '
        Me.btnAsset_Cancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_Cancel.Location = New System.Drawing.Point(436, 541)
        Me.btnAsset_Cancel.Name = "btnAsset_Cancel"
        Me.btnAsset_Cancel.Size = New System.Drawing.Size(104, 26)
        Me.btnAsset_Cancel.TabIndex = 6
        Me.btnAsset_Cancel.Text = "&Cancel"
        Me.btnAsset_Cancel.UseVisualStyleBackColor = True
        '
        'btnAsset_OK
        '
        Me.btnAsset_OK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_OK.Location = New System.Drawing.Point(320, 541)
        Me.btnAsset_OK.Name = "btnAsset_OK"
        Me.btnAsset_OK.Size = New System.Drawing.Size(104, 26)
        Me.btnAsset_OK.TabIndex = 5
        Me.btnAsset_OK.Text = "&OK"
        Me.btnAsset_OK.UseVisualStyleBackColor = True
        '
        'StatusStripAsset_1
        '
        Me.StatusStripAsset_1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStripAsset_1.Location = New System.Drawing.Point(0, 620)
        Me.StatusStripAsset_1.Name = "StatusStripAsset_1"
        Me.StatusStripAsset_1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStripAsset_1.TabIndex = 48
        Me.StatusStripAsset_1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'btnAsset_ViewJurnal
        '
        Me.btnAsset_ViewJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAsset_ViewJurnal.Location = New System.Drawing.Point(12, 582)
        Me.btnAsset_ViewJurnal.Name = "btnAsset_ViewJurnal"
        Me.btnAsset_ViewJurnal.Size = New System.Drawing.Size(80, 26)
        Me.btnAsset_ViewJurnal.TabIndex = 41
        Me.btnAsset_ViewJurnal.TabStop = False
        Me.btnAsset_ViewJurnal.Text = "View Journal"
        Me.btnAsset_ViewJurnal.UseVisualStyleBackColor = True
        '
        'btnAsset_Delete
        '
        Me.btnAsset_Delete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_Delete.Location = New System.Drawing.Point(446, 582)
        Me.btnAsset_Delete.Name = "btnAsset_Delete"
        Me.btnAsset_Delete.Size = New System.Drawing.Size(80, 26)
        Me.btnAsset_Delete.TabIndex = 45
        Me.btnAsset_Delete.Text = "&Delete"
        Me.btnAsset_Delete.UseVisualStyleBackColor = True
        '
        'btnAsset_Print
        '
        Me.btnAsset_Print.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_Print.Location = New System.Drawing.Point(534, 582)
        Me.btnAsset_Print.Name = "btnAsset_Print"
        Me.btnAsset_Print.Size = New System.Drawing.Size(80, 26)
        Me.btnAsset_Print.TabIndex = 46
        Me.btnAsset_Print.Text = "&Print"
        Me.btnAsset_Print.UseVisualStyleBackColor = True
        '
        'btnAsset_Exit
        '
        Me.btnAsset_Exit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_Exit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnAsset_Exit.Location = New System.Drawing.Point(622, 582)
        Me.btnAsset_Exit.Name = "btnAsset_Exit"
        Me.btnAsset_Exit.Size = New System.Drawing.Size(80, 26)
        Me.btnAsset_Exit.TabIndex = 47
        Me.btnAsset_Exit.Text = "E&xit"
        Me.btnAsset_Exit.UseVisualStyleBackColor = True
        '
        'btnAsset_Edit
        '
        Me.btnAsset_Edit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_Edit.Location = New System.Drawing.Point(358, 582)
        Me.btnAsset_Edit.Name = "btnAsset_Edit"
        Me.btnAsset_Edit.Size = New System.Drawing.Size(80, 26)
        Me.btnAsset_Edit.TabIndex = 44
        Me.btnAsset_Edit.Text = "&Edit"
        Me.btnAsset_Edit.UseVisualStyleBackColor = True
        '
        'btnAsset_Find
        '
        Me.btnAsset_Find.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_Find.Location = New System.Drawing.Point(182, 582)
        Me.btnAsset_Find.Name = "btnAsset_Find"
        Me.btnAsset_Find.Size = New System.Drawing.Size(80, 26)
        Me.btnAsset_Find.TabIndex = 42
        Me.btnAsset_Find.Text = "&Find"
        Me.btnAsset_Find.UseVisualStyleBackColor = True
        '
        'btnAsset_New
        '
        Me.btnAsset_New.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnAsset_New.Location = New System.Drawing.Point(270, 582)
        Me.btnAsset_New.Name = "btnAsset_New"
        Me.btnAsset_New.Size = New System.Drawing.Size(80, 26)
        Me.btnAsset_New.TabIndex = 43
        Me.btnAsset_New.Text = "&New"
        Me.btnAsset_New.UseVisualStyleBackColor = True
        '
        'FrmTransaksiJournal_AssetDepreciation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(884, 642)
        Me.Controls.Add(Me.gboAsset_Detail)
        Me.Controls.Add(Me.StatusStripAsset_1)
        Me.Controls.Add(Me.btnAsset_ViewJurnal)
        Me.Controls.Add(Me.btnAsset_Delete)
        Me.Controls.Add(Me.btnAsset_Print)
        Me.Controls.Add(Me.btnAsset_Exit)
        Me.Controls.Add(Me.btnAsset_Edit)
        Me.Controls.Add(Me.btnAsset_Find)
        Me.Controls.Add(Me.btnAsset_New)
        Me.Name = "FrmTransaksiJournal_AssetDepreciation"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Asset Depreciation"
        Me.gboAsset_Detail.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvAsset_ItemAsset, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvAsset_Item, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1_Asset.Panel1.ResumeLayout(False)
        Me.SplitContainer1_Asset.Panel1.PerformLayout()
        Me.SplitContainer1_Asset.Panel2.ResumeLayout(False)
        Me.SplitContainer1_Asset.Panel2.PerformLayout()
        Me.SplitContainer1_Asset.ResumeLayout(False)
        Me.StatusStripAsset_1.ResumeLayout(False)
        Me.StatusStripAsset_1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboAsset_Detail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1_Asset As System.Windows.Forms.SplitContainer
    Friend WithEvents txtAsset_ReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtAsset_PrimaryJournal As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtRefNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpAsset_Date As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtAsset_Remarks As System.Windows.Forms.TextBox
    Friend WithEvents TxtAsset_Unbalance As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnAsset_Cancel As System.Windows.Forms.Button
    Friend WithEvents btnAsset_OK As System.Windows.Forms.Button
    Public WithEvents dgvAsset_Item As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStripAsset_1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnAsset_ViewJurnal As System.Windows.Forms.Button
    Friend WithEvents btnAsset_Delete As System.Windows.Forms.Button
    Friend WithEvents btnAsset_Print As System.Windows.Forms.Button
    Friend WithEvents btnAsset_Exit As System.Windows.Forms.Button
    Friend WithEvents btnAsset_Edit As System.Windows.Forms.Button
    Friend WithEvents btnAsset_Find As System.Windows.Forms.Button
    Friend WithEvents btnAsset_New As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents dgvAsset_ItemAsset As System.Windows.Forms.DataGridView
    Friend WithEvents ColAsset_Account As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_DescAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_Debet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_Kredit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_ExchangeRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_OriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_Debet_OriginalCurr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_Credit_OriginalCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_Dept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_Job As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_PrimaryDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_PrimaryJob As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_PrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAsset_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAssetNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_AssetCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_AssetDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_DepreciationPeriod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_DepreciationMethod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_LastBookValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_RevaluationValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_DisposalValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_SalvageValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_DepreciatedValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_EndValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_PrimaryAkumDepAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColItemAsset_PrimaryDepExpenseAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAsset As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
