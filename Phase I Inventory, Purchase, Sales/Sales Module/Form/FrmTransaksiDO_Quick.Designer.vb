<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDO_Quick
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiDO_Quick))
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
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvCreatedBy = New System.Windows.Forms.DataGridView
        Me.ColUserNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUser_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvCustomer = New System.Windows.Forms.DataGridView
        Me.ColCust_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCust_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvItemProductGroup = New System.Windows.Forms.DataGridView
        Me.ColProductGroup_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Level = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColProductGroup_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductGroup_PrimaryAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.txtUserCode = New System.Windows.Forms.TextBox
        Me.btnFindUser = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtLocCode = New System.Windows.Forms.TextBox
        Me.txtPrimaryWarehouse = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.rbInventoryBooking = New System.Windows.Forms.RadioButton
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.rbSO = New System.Windows.Forms.RadioButton
        Me.BtnFindLoc = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.LlbClearFilter = New System.Windows.Forms.LinkLabel
        Me.chkDirectSave = New System.Windows.Forms.CheckBox
        Me.ChkIgnoreLowTolerance = New System.Windows.Forms.CheckBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtCancellationReason = New System.Windows.Forms.TextBox
        Me.cboCustomerFormat = New System.Windows.Forms.ComboBox
        Me.GBPeriod = New System.Windows.Forms.GroupBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.btnFilter = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.TxtTransactionNumber = New System.Windows.Forms.TextBox
        Me.btnFindTransaction = New System.Windows.Forms.Button
        Me.lblFind = New System.Windows.Forms.Label
        Me.TxtPrimaryTransaction = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.TxtPrimaryLoc = New System.Windows.Forms.TextBox
        Me.txtPrimaryUser = New System.Windows.Forms.TextBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDocumentNO = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTolerance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyDO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColChkDO = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColRemarksItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPoduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryTransaction = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaruUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefPrec_QtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimalMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchControlled = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryControlled = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocationControlled = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier_CustAddress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFakturPajak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Colidentifier_address_customer_ship = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Colidentifier_address_ship_to = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary_Customer_Ship_To = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.prg01 = New System.Windows.Forms.ProgressBar
        Me.LblProgress = New System.Windows.Forms.Label
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BgwCustomer = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.dgvItemReject = New System.Windows.Forms.DataGridView
        Me.colReject_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_CustomerCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_DocumentNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.colReject_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_ExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_Plus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_Currency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_UnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_DiscPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_DiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_QtyDO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_ChkCreateDO = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colReject_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryTransaction = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryCustomer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_QtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_RatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_RatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_DefprecDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_UnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_DecimalMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_IsBatch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_IsExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_IsLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_IdentifierCustAddress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_Tax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_FakturPajak = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_IdentifierAddressCustomerShipTo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_IdentifierAddressShipTo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimaryCustomerShipTo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_PrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReject_SONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvCreatedBy, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBPeriod.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemReject, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        resources.ApplyResources(Me.SplitContainer1, "SplitContainer1")
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtUserCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindUser)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtLocCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryWarehouse)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.rbInventoryBooking)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbUncheckAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.llbCheckAll)
        Me.SplitContainer1.Panel2.Controls.Add(Me.rbSO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.BtnFindLoc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LlbClearFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkDirectSave)
        Me.SplitContainer1.Panel2.Controls.Add(Me.ChkIgnoreLowTolerance)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCancellationReason)
        Me.SplitContainer1.Panel2.Controls.Add(Me.cboCustomerFormat)
        Me.SplitContainer1.Panel2.Controls.Add(Me.GBPeriod)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnStart)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTransactionNumber)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnFindTransaction)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblFind)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryTransaction)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryLoc)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtPrimaryUser)
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        resources.ApplyResources(Me.TabControl1, "TabControl1")
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvCreatedBy)
        resources.ApplyResources(Me.TabPage1, "TabPage1")
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvCreatedBy
        '
        Me.dgvCreatedBy.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCreatedBy.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCreatedBy.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColUserNo, Me.ColUserCode, Me.ColUserName, Me.ColUser_Primary})
        resources.ApplyResources(Me.dgvCreatedBy, "dgvCreatedBy")
        Me.dgvCreatedBy.MultiSelect = False
        Me.dgvCreatedBy.Name = "dgvCreatedBy"
        Me.dgvCreatedBy.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCreatedBy.RowTemplate.Height = 24
        '
        'ColUserNo
        '
        resources.ApplyResources(Me.ColUserNo, "ColUserNo")
        Me.ColUserNo.Name = "ColUserNo"
        '
        'ColUserCode
        '
        resources.ApplyResources(Me.ColUserCode, "ColUserCode")
        Me.ColUserCode.Name = "ColUserCode"
        '
        'ColUserName
        '
        resources.ApplyResources(Me.ColUserName, "ColUserName")
        Me.ColUserName.Name = "ColUserName"
        '
        'ColUser_Primary
        '
        resources.ApplyResources(Me.ColUser_Primary, "ColUser_Primary")
        Me.ColUser_Primary.Name = "ColUser_Primary"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvCustomer)
        resources.ApplyResources(Me.TabPage2, "TabPage2")
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvCustomer
        '
        Me.DgvCustomer.AllowUserToResizeColumns = False
        Me.DgvCustomer.AllowUserToResizeRows = False
        Me.DgvCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCustomer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCust_No, Me.ColCust_Code, Me.ColCust_Name, Me.ColCust_Remarks, Me.ColCust_Primary})
        resources.ApplyResources(Me.DgvCustomer, "DgvCustomer")
        Me.DgvCustomer.MultiSelect = False
        Me.DgvCustomer.Name = "DgvCustomer"
        Me.DgvCustomer.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCustomer.RowTemplate.Height = 24
        '
        'ColCust_No
        '
        resources.ApplyResources(Me.ColCust_No, "ColCust_No")
        Me.ColCust_No.Name = "ColCust_No"
        Me.ColCust_No.ReadOnly = True
        '
        'ColCust_Code
        '
        resources.ApplyResources(Me.ColCust_Code, "ColCust_Code")
        Me.ColCust_Code.Name = "ColCust_Code"
        '
        'ColCust_Name
        '
        resources.ApplyResources(Me.ColCust_Name, "ColCust_Name")
        Me.ColCust_Name.Name = "ColCust_Name"
        '
        'ColCust_Remarks
        '
        resources.ApplyResources(Me.ColCust_Remarks, "ColCust_Remarks")
        Me.ColCust_Remarks.Name = "ColCust_Remarks"
        Me.ColCust_Remarks.ReadOnly = True
        '
        'ColCust_Primary
        '
        resources.ApplyResources(Me.ColCust_Primary, "ColCust_Primary")
        Me.ColCust_Primary.Name = "ColCust_Primary"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvItemProductGroup)
        resources.ApplyResources(Me.TabPage3, "TabPage3")
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvItemProductGroup
        '
        Me.dgvItemProductGroup.AllowUserToResizeColumns = False
        Me.dgvItemProductGroup.AllowUserToResizeRows = False
        Me.dgvItemProductGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemProductGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemProductGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductGroup_No, Me.ColProductGroup_Level, Me.ColProductGroup_CodeNo, Me.ColProductGroup_Desc, Me.ColProductGroup_Remarks, Me.ColProductGroup_PrimaryAcuan})
        resources.ApplyResources(Me.dgvItemProductGroup, "dgvItemProductGroup")
        Me.dgvItemProductGroup.MultiSelect = False
        Me.dgvItemProductGroup.Name = "dgvItemProductGroup"
        Me.dgvItemProductGroup.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemProductGroup.RowTemplate.Height = 24
        '
        'ColProductGroup_No
        '
        resources.ApplyResources(Me.ColProductGroup_No, "ColProductGroup_No")
        Me.ColProductGroup_No.Name = "ColProductGroup_No"
        Me.ColProductGroup_No.ReadOnly = True
        '
        'ColProductGroup_Level
        '
        resources.ApplyResources(Me.ColProductGroup_Level, "ColProductGroup_Level")
        Me.ColProductGroup_Level.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.ColProductGroup_Level.Name = "ColProductGroup_Level"
        '
        'ColProductGroup_CodeNo
        '
        resources.ApplyResources(Me.ColProductGroup_CodeNo, "ColProductGroup_CodeNo")
        Me.ColProductGroup_CodeNo.Name = "ColProductGroup_CodeNo"
        '
        'ColProductGroup_Desc
        '
        resources.ApplyResources(Me.ColProductGroup_Desc, "ColProductGroup_Desc")
        Me.ColProductGroup_Desc.Name = "ColProductGroup_Desc"
        Me.ColProductGroup_Desc.ReadOnly = True
        '
        'ColProductGroup_Remarks
        '
        resources.ApplyResources(Me.ColProductGroup_Remarks, "ColProductGroup_Remarks")
        Me.ColProductGroup_Remarks.Name = "ColProductGroup_Remarks"
        Me.ColProductGroup_Remarks.ReadOnly = True
        '
        'ColProductGroup_PrimaryAcuan
        '
        resources.ApplyResources(Me.ColProductGroup_PrimaryAcuan, "ColProductGroup_PrimaryAcuan")
        Me.ColProductGroup_PrimaryAcuan.Name = "ColProductGroup_PrimaryAcuan"
        Me.ColProductGroup_PrimaryAcuan.ReadOnly = True
        '
        'txtUserCode
        '
        Me.txtUserCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        resources.ApplyResources(Me.txtUserCode, "txtUserCode")
        Me.txtUserCode.Name = "txtUserCode"
        '
        'btnFindUser
        '
        resources.ApplyResources(Me.btnFindUser, "btnFindUser")
        Me.btnFindUser.Name = "btnFindUser"
        Me.btnFindUser.TabStop = False
        Me.btnFindUser.UseVisualStyleBackColor = True
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'TxtLocCode
        '
        Me.TxtLocCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        resources.ApplyResources(Me.TxtLocCode, "TxtLocCode")
        Me.TxtLocCode.Name = "TxtLocCode"
        '
        'txtPrimaryWarehouse
        '
        resources.ApplyResources(Me.txtPrimaryWarehouse, "txtPrimaryWarehouse")
        Me.txtPrimaryWarehouse.Name = "txtPrimaryWarehouse"
        Me.txtPrimaryWarehouse.TabStop = False
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'rbInventoryBooking
        '
        resources.ApplyResources(Me.rbInventoryBooking, "rbInventoryBooking")
        Me.rbInventoryBooking.Name = "rbInventoryBooking"
        Me.rbInventoryBooking.TabStop = True
        Me.rbInventoryBooking.UseVisualStyleBackColor = True
        '
        'llbUncheckAll
        '
        resources.ApplyResources(Me.llbUncheckAll, "llbUncheckAll")
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.TabStop = True
        '
        'llbCheckAll
        '
        resources.ApplyResources(Me.llbCheckAll, "llbCheckAll")
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.TabStop = True
        '
        'rbSO
        '
        resources.ApplyResources(Me.rbSO, "rbSO")
        Me.rbSO.Name = "rbSO"
        Me.rbSO.TabStop = True
        Me.rbSO.UseVisualStyleBackColor = True
        '
        'BtnFindLoc
        '
        resources.ApplyResources(Me.BtnFindLoc, "BtnFindLoc")
        Me.BtnFindLoc.Name = "BtnFindLoc"
        Me.BtnFindLoc.TabStop = False
        Me.BtnFindLoc.UseVisualStyleBackColor = True
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'LlbClearFilter
        '
        resources.ApplyResources(Me.LlbClearFilter, "LlbClearFilter")
        Me.LlbClearFilter.Name = "LlbClearFilter"
        Me.LlbClearFilter.TabStop = True
        '
        'chkDirectSave
        '
        resources.ApplyResources(Me.chkDirectSave, "chkDirectSave")
        Me.chkDirectSave.Name = "chkDirectSave"
        Me.chkDirectSave.UseVisualStyleBackColor = True
        '
        'ChkIgnoreLowTolerance
        '
        resources.ApplyResources(Me.ChkIgnoreLowTolerance, "ChkIgnoreLowTolerance")
        Me.ChkIgnoreLowTolerance.Name = "ChkIgnoreLowTolerance"
        Me.ChkIgnoreLowTolerance.UseVisualStyleBackColor = True
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'TxtCancellationReason
        '
        resources.ApplyResources(Me.TxtCancellationReason, "TxtCancellationReason")
        Me.TxtCancellationReason.Name = "TxtCancellationReason"
        '
        'cboCustomerFormat
        '
        Me.cboCustomerFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerFormat.FormattingEnabled = True
        Me.cboCustomerFormat.Items.AddRange(New Object() {resources.GetString("cboCustomerFormat.Items"), resources.GetString("cboCustomerFormat.Items1"), resources.GetString("cboCustomerFormat.Items2")})
        resources.ApplyResources(Me.cboCustomerFormat, "cboCustomerFormat")
        Me.cboCustomerFormat.Name = "cboCustomerFormat"
        '
        'GBPeriod
        '
        Me.GBPeriod.Controls.Add(Me.dtpFrom)
        Me.GBPeriod.Controls.Add(Me.Label4)
        Me.GBPeriod.Controls.Add(Me.Label3)
        Me.GBPeriod.Controls.Add(Me.dtpTo)
        resources.ApplyResources(Me.GBPeriod, "GBPeriod")
        Me.GBPeriod.Name = "GBPeriod"
        Me.GBPeriod.TabStop = False
        '
        'dtpFrom
        '
        resources.ApplyResources(Me.dtpFrom, "dtpFrom")
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Name = "dtpFrom"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'dtpTo
        '
        resources.ApplyResources(Me.dtpTo, "dtpTo")
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Name = "dtpTo"
        '
        'btnFilter
        '
        resources.ApplyResources(Me.btnFilter, "btnFilter")
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        resources.ApplyResources(Me.btnStart, "btnStart")
        Me.btnStart.Name = "btnStart"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'TxtTransactionNumber
        '
        Me.TxtTransactionNumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        resources.ApplyResources(Me.TxtTransactionNumber, "TxtTransactionNumber")
        Me.TxtTransactionNumber.Name = "TxtTransactionNumber"
        '
        'btnFindTransaction
        '
        resources.ApplyResources(Me.btnFindTransaction, "btnFindTransaction")
        Me.btnFindTransaction.Name = "btnFindTransaction"
        Me.btnFindTransaction.TabStop = False
        Me.btnFindTransaction.UseVisualStyleBackColor = True
        '
        'lblFind
        '
        resources.ApplyResources(Me.lblFind, "lblFind")
        Me.lblFind.Name = "lblFind"
        '
        'TxtPrimaryTransaction
        '
        resources.ApplyResources(Me.TxtPrimaryTransaction, "TxtPrimaryTransaction")
        Me.TxtPrimaryTransaction.Name = "TxtPrimaryTransaction"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        resources.ApplyResources(Me.btnClose, "btnClose")
        Me.btnClose.Name = "btnClose"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'TxtPrimaryLoc
        '
        resources.ApplyResources(Me.TxtPrimaryLoc, "TxtPrimaryLoc")
        Me.TxtPrimaryLoc.Name = "TxtPrimaryLoc"
        '
        'txtPrimaryUser
        '
        resources.ApplyResources(Me.txtPrimaryUser, "txtPrimaryUser")
        Me.txtPrimaryUser.Name = "txtPrimaryUser"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        resources.ApplyResources(Me.dgvItem, "dgvItem")
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCustomer, Me.ColDocumentNO, Me.ColDate, Me.ColCodeNo, Me.ColDesc, Me.ColBatchNo, Me.ColExpiryDate, Me.ColQty, Me.ColTolerance, Me.ColCurrency, Me.ColUnitPrice, Me.ColDiscPerct, Me.ColDiscAmount, Me.ColAmount, Me.ColQtyDO, Me.ColUnit, Me.ColChkDO, Me.ColRemarksItem, Me.ColPrimaryPoduct, Me.ColPrimaryUnitMulti, Me.ColPrimaryTransaction, Me.ColPrimaryCustomer, Me.ColPrimaryLocation, Me.ColPrimaryWarehouse, Me.ColPrimaryCurrency, Me.ColQtyDefault, Me.ColPrimaruUnitDefault, Me.ColRatioDefault, Me.ColRatioReference, Me.ColDefPrec_QtyDefault, Me.ColUnitDefault, Me.ColDecimalMulti, Me.ColBatchControlled, Me.ColExpiryControlled, Me.ColLocationControlled, Me.ColIdentifier_CustAddress, Me.ColTax, Me.ColFakturPajak, Me.Colidentifier_address_customer_ship, Me.Colidentifier_address_ship_to, Me.ColPrimary_Customer_Ship_To, Me.ColPrimarySO, Me.ColSONo})
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        '
        'ColNo
        '
        resources.ApplyResources(Me.ColNo, "ColNo")
        Me.ColNo.Name = "ColNo"
        '
        'ColCustomer
        '
        resources.ApplyResources(Me.ColCustomer, "ColCustomer")
        Me.ColCustomer.Name = "ColCustomer"
        '
        'ColDocumentNO
        '
        Me.ColDocumentNO.ActiveLinkColor = System.Drawing.Color.Blue
        resources.ApplyResources(Me.ColDocumentNO, "ColDocumentNO")
        Me.ColDocumentNO.Name = "ColDocumentNO"
        Me.ColDocumentNO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDocumentNO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColDocumentNO.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'ColDate
        '
        DataGridViewCellStyle1.Format = "dd MMM yyyy"
        Me.ColDate.DefaultCellStyle = DataGridViewCellStyle1
        resources.ApplyResources(Me.ColDate, "ColDate")
        Me.ColDate.Name = "ColDate"
        '
        'ColCodeNo
        '
        resources.ApplyResources(Me.ColCodeNo, "ColCodeNo")
        Me.ColCodeNo.Name = "ColCodeNo"
        '
        'ColDesc
        '
        resources.ApplyResources(Me.ColDesc, "ColDesc")
        Me.ColDesc.Name = "ColDesc"
        '
        'ColBatchNo
        '
        resources.ApplyResources(Me.ColBatchNo, "ColBatchNo")
        Me.ColBatchNo.Name = "ColBatchNo"
        '
        'ColExpiryDate
        '
        resources.ApplyResources(Me.ColExpiryDate, "ColExpiryDate")
        Me.ColExpiryDate.Name = "ColExpiryDate"
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        resources.ApplyResources(Me.ColQty, "ColQty")
        Me.ColQty.Name = "ColQty"
        '
        'ColTolerance
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTolerance.DefaultCellStyle = DataGridViewCellStyle3
        resources.ApplyResources(Me.ColTolerance, "ColTolerance")
        Me.ColTolerance.Name = "ColTolerance"
        '
        'ColCurrency
        '
        resources.ApplyResources(Me.ColCurrency, "ColCurrency")
        Me.ColCurrency.Name = "ColCurrency"
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle4
        resources.ApplyResources(Me.ColUnitPrice, "ColUnitPrice")
        Me.ColUnitPrice.Name = "ColUnitPrice"
        '
        'ColDiscPerct
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscPerct.DefaultCellStyle = DataGridViewCellStyle5
        resources.ApplyResources(Me.ColDiscPerct, "ColDiscPerct")
        Me.ColDiscPerct.Name = "ColDiscPerct"
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle6
        resources.ApplyResources(Me.ColDiscAmount, "ColDiscAmount")
        Me.ColDiscAmount.Name = "ColDiscAmount"
        '
        'ColAmount
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle7
        resources.ApplyResources(Me.ColAmount, "ColAmount")
        Me.ColAmount.Name = "ColAmount"
        '
        'ColQtyDO
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyDO.DefaultCellStyle = DataGridViewCellStyle8
        resources.ApplyResources(Me.ColQtyDO, "ColQtyDO")
        Me.ColQtyDO.Name = "ColQtyDO"
        '
        'ColUnit
        '
        resources.ApplyResources(Me.ColUnit, "ColUnit")
        Me.ColUnit.Name = "ColUnit"
        '
        'ColChkDO
        '
        resources.ApplyResources(Me.ColChkDO, "ColChkDO")
        Me.ColChkDO.Name = "ColChkDO"
        '
        'ColRemarksItem
        '
        resources.ApplyResources(Me.ColRemarksItem, "ColRemarksItem")
        Me.ColRemarksItem.Name = "ColRemarksItem"
        '
        'ColPrimaryPoduct
        '
        resources.ApplyResources(Me.ColPrimaryPoduct, "ColPrimaryPoduct")
        Me.ColPrimaryPoduct.Name = "ColPrimaryPoduct"
        '
        'ColPrimaryUnitMulti
        '
        resources.ApplyResources(Me.ColPrimaryUnitMulti, "ColPrimaryUnitMulti")
        Me.ColPrimaryUnitMulti.Name = "ColPrimaryUnitMulti"
        '
        'ColPrimaryTransaction
        '
        resources.ApplyResources(Me.ColPrimaryTransaction, "ColPrimaryTransaction")
        Me.ColPrimaryTransaction.Name = "ColPrimaryTransaction"
        '
        'ColPrimaryCustomer
        '
        resources.ApplyResources(Me.ColPrimaryCustomer, "ColPrimaryCustomer")
        Me.ColPrimaryCustomer.Name = "ColPrimaryCustomer"
        Me.ColPrimaryCustomer.ReadOnly = True
        '
        'ColPrimaryLocation
        '
        resources.ApplyResources(Me.ColPrimaryLocation, "ColPrimaryLocation")
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.ReadOnly = True
        '
        'ColPrimaryWarehouse
        '
        resources.ApplyResources(Me.ColPrimaryWarehouse, "ColPrimaryWarehouse")
        Me.ColPrimaryWarehouse.Name = "ColPrimaryWarehouse"
        Me.ColPrimaryWarehouse.ReadOnly = True
        '
        'ColPrimaryCurrency
        '
        resources.ApplyResources(Me.ColPrimaryCurrency, "ColPrimaryCurrency")
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.ReadOnly = True
        '
        'ColQtyDefault
        '
        resources.ApplyResources(Me.ColQtyDefault, "ColQtyDefault")
        Me.ColQtyDefault.Name = "ColQtyDefault"
        '
        'ColPrimaruUnitDefault
        '
        resources.ApplyResources(Me.ColPrimaruUnitDefault, "ColPrimaruUnitDefault")
        Me.ColPrimaruUnitDefault.Name = "ColPrimaruUnitDefault"
        '
        'ColRatioDefault
        '
        resources.ApplyResources(Me.ColRatioDefault, "ColRatioDefault")
        Me.ColRatioDefault.Name = "ColRatioDefault"
        '
        'ColRatioReference
        '
        resources.ApplyResources(Me.ColRatioReference, "ColRatioReference")
        Me.ColRatioReference.Name = "ColRatioReference"
        '
        'ColDefPrec_QtyDefault
        '
        resources.ApplyResources(Me.ColDefPrec_QtyDefault, "ColDefPrec_QtyDefault")
        Me.ColDefPrec_QtyDefault.Name = "ColDefPrec_QtyDefault"
        '
        'ColUnitDefault
        '
        resources.ApplyResources(Me.ColUnitDefault, "ColUnitDefault")
        Me.ColUnitDefault.Name = "ColUnitDefault"
        '
        'ColDecimalMulti
        '
        resources.ApplyResources(Me.ColDecimalMulti, "ColDecimalMulti")
        Me.ColDecimalMulti.Name = "ColDecimalMulti"
        Me.ColDecimalMulti.ReadOnly = True
        '
        'ColBatchControlled
        '
        resources.ApplyResources(Me.ColBatchControlled, "ColBatchControlled")
        Me.ColBatchControlled.Name = "ColBatchControlled"
        Me.ColBatchControlled.ReadOnly = True
        '
        'ColExpiryControlled
        '
        resources.ApplyResources(Me.ColExpiryControlled, "ColExpiryControlled")
        Me.ColExpiryControlled.Name = "ColExpiryControlled"
        Me.ColExpiryControlled.ReadOnly = True
        '
        'ColLocationControlled
        '
        resources.ApplyResources(Me.ColLocationControlled, "ColLocationControlled")
        Me.ColLocationControlled.Name = "ColLocationControlled"
        Me.ColLocationControlled.ReadOnly = True
        '
        'ColIdentifier_CustAddress
        '
        resources.ApplyResources(Me.ColIdentifier_CustAddress, "ColIdentifier_CustAddress")
        Me.ColIdentifier_CustAddress.Name = "ColIdentifier_CustAddress"
        '
        'ColTax
        '
        resources.ApplyResources(Me.ColTax, "ColTax")
        Me.ColTax.Name = "ColTax"
        '
        'ColFakturPajak
        '
        resources.ApplyResources(Me.ColFakturPajak, "ColFakturPajak")
        Me.ColFakturPajak.Name = "ColFakturPajak"
        '
        'Colidentifier_address_customer_ship
        '
        resources.ApplyResources(Me.Colidentifier_address_customer_ship, "Colidentifier_address_customer_ship")
        Me.Colidentifier_address_customer_ship.Name = "Colidentifier_address_customer_ship"
        '
        'Colidentifier_address_ship_to
        '
        resources.ApplyResources(Me.Colidentifier_address_ship_to, "Colidentifier_address_ship_to")
        Me.Colidentifier_address_ship_to.Name = "Colidentifier_address_ship_to"
        '
        'ColPrimary_Customer_Ship_To
        '
        resources.ApplyResources(Me.ColPrimary_Customer_Ship_To, "ColPrimary_Customer_Ship_To")
        Me.ColPrimary_Customer_Ship_To.Name = "ColPrimary_Customer_Ship_To"
        '
        'ColPrimarySO
        '
        resources.ApplyResources(Me.ColPrimarySO, "ColPrimarySO")
        Me.ColPrimarySO.Name = "ColPrimarySO"
        '
        'ColSONo
        '
        resources.ApplyResources(Me.ColSONo, "ColSONo")
        Me.ColSONo.Name = "ColSONo"
        '
        'prg01
        '
        resources.ApplyResources(Me.prg01, "prg01")
        Me.prg01.Name = "prg01"
        '
        'LblProgress
        '
        resources.ApplyResources(Me.LblProgress, "LblProgress")
        Me.LblProgress.Name = "LblProgress"
        '
        'BackgroundWorker1
        '
        '
        'BgwCustomer
        '
        '
        'BackgroundWorker3
        '
        '
        'BackgroundWorker2
        '
        '
        'dgvItemReject
        '
        Me.dgvItemReject.AllowUserToAddRows = False
        Me.dgvItemReject.AllowUserToDeleteRows = False
        Me.dgvItemReject.AllowUserToResizeColumns = False
        Me.dgvItemReject.AllowUserToResizeRows = False
        resources.ApplyResources(Me.dgvItemReject, "dgvItemReject")
        Me.dgvItemReject.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemReject.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItemReject.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colReject_No, Me.colReject_CustomerCode, Me.colReject_DocumentNo, Me.colReject_Date, Me.colReject_CodeNo, Me.colReject_Description, Me.colReject_BatchNo, Me.colReject_ExpiryDate, Me.colReject_Qty, Me.colReject_Plus, Me.colReject_Currency, Me.colReject_UnitPrice, Me.colReject_DiscPerc, Me.colReject_DiscAmount, Me.colReject_Amount, Me.colReject_QtyDO, Me.colReject_Unit, Me.colReject_ChkCreateDO, Me.colReject_Remarks, Me.colReject_PrimaryProduct, Me.colReject_PrimaryUnitMulti, Me.colReject_PrimaryTransaction, Me.colReject_PrimaryCustomer, Me.colReject_PrimaryLocation, Me.colReject_PrimaryWarehouse, Me.colReject_PrimaryCurrency, Me.colReject_QtyDefault, Me.colReject_PrimaryUnitDefault, Me.colReject_RatioDefault, Me.colReject_RatioReference, Me.colReject_DefprecDefault, Me.colReject_UnitDefault, Me.colReject_DecimalMulti, Me.colReject_IsBatch, Me.colReject_IsExpiry, Me.colReject_IsLocation, Me.colReject_IdentifierCustAddress, Me.colReject_Tax, Me.colReject_FakturPajak, Me.colReject_IdentifierAddressCustomerShipTo, Me.colReject_IdentifierAddressShipTo, Me.colReject_PrimaryCustomerShipTo, Me.colReject_PrimarySO, Me.colReject_SONo})
        Me.dgvItemReject.Name = "dgvItemReject"
        Me.dgvItemReject.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItemReject.RowTemplate.Height = 24
        Me.dgvItemReject.TabStop = False
        '
        'colReject_No
        '
        resources.ApplyResources(Me.colReject_No, "colReject_No")
        Me.colReject_No.Name = "colReject_No"
        '
        'colReject_CustomerCode
        '
        resources.ApplyResources(Me.colReject_CustomerCode, "colReject_CustomerCode")
        Me.colReject_CustomerCode.Name = "colReject_CustomerCode"
        '
        'colReject_DocumentNo
        '
        Me.colReject_DocumentNo.ActiveLinkColor = System.Drawing.Color.Blue
        resources.ApplyResources(Me.colReject_DocumentNo, "colReject_DocumentNo")
        Me.colReject_DocumentNo.Name = "colReject_DocumentNo"
        Me.colReject_DocumentNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.colReject_DocumentNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.colReject_DocumentNo.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'colReject_Date
        '
        DataGridViewCellStyle9.Format = "dd MMM yyyy"
        Me.colReject_Date.DefaultCellStyle = DataGridViewCellStyle9
        resources.ApplyResources(Me.colReject_Date, "colReject_Date")
        Me.colReject_Date.Name = "colReject_Date"
        '
        'colReject_CodeNo
        '
        resources.ApplyResources(Me.colReject_CodeNo, "colReject_CodeNo")
        Me.colReject_CodeNo.Name = "colReject_CodeNo"
        '
        'colReject_Description
        '
        resources.ApplyResources(Me.colReject_Description, "colReject_Description")
        Me.colReject_Description.Name = "colReject_Description"
        '
        'colReject_BatchNo
        '
        resources.ApplyResources(Me.colReject_BatchNo, "colReject_BatchNo")
        Me.colReject_BatchNo.Name = "colReject_BatchNo"
        '
        'colReject_ExpiryDate
        '
        resources.ApplyResources(Me.colReject_ExpiryDate, "colReject_ExpiryDate")
        Me.colReject_ExpiryDate.Name = "colReject_ExpiryDate"
        '
        'colReject_Qty
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colReject_Qty.DefaultCellStyle = DataGridViewCellStyle10
        resources.ApplyResources(Me.colReject_Qty, "colReject_Qty")
        Me.colReject_Qty.Name = "colReject_Qty"
        '
        'colReject_Plus
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colReject_Plus.DefaultCellStyle = DataGridViewCellStyle11
        resources.ApplyResources(Me.colReject_Plus, "colReject_Plus")
        Me.colReject_Plus.Name = "colReject_Plus"
        '
        'colReject_Currency
        '
        resources.ApplyResources(Me.colReject_Currency, "colReject_Currency")
        Me.colReject_Currency.Name = "colReject_Currency"
        '
        'colReject_UnitPrice
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colReject_UnitPrice.DefaultCellStyle = DataGridViewCellStyle12
        resources.ApplyResources(Me.colReject_UnitPrice, "colReject_UnitPrice")
        Me.colReject_UnitPrice.Name = "colReject_UnitPrice"
        '
        'colReject_DiscPerc
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colReject_DiscPerc.DefaultCellStyle = DataGridViewCellStyle13
        resources.ApplyResources(Me.colReject_DiscPerc, "colReject_DiscPerc")
        Me.colReject_DiscPerc.Name = "colReject_DiscPerc"
        '
        'colReject_DiscAmount
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colReject_DiscAmount.DefaultCellStyle = DataGridViewCellStyle14
        resources.ApplyResources(Me.colReject_DiscAmount, "colReject_DiscAmount")
        Me.colReject_DiscAmount.Name = "colReject_DiscAmount"
        '
        'colReject_Amount
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colReject_Amount.DefaultCellStyle = DataGridViewCellStyle15
        resources.ApplyResources(Me.colReject_Amount, "colReject_Amount")
        Me.colReject_Amount.Name = "colReject_Amount"
        '
        'colReject_QtyDO
        '
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colReject_QtyDO.DefaultCellStyle = DataGridViewCellStyle16
        resources.ApplyResources(Me.colReject_QtyDO, "colReject_QtyDO")
        Me.colReject_QtyDO.Name = "colReject_QtyDO"
        '
        'colReject_Unit
        '
        resources.ApplyResources(Me.colReject_Unit, "colReject_Unit")
        Me.colReject_Unit.Name = "colReject_Unit"
        '
        'colReject_ChkCreateDO
        '
        resources.ApplyResources(Me.colReject_ChkCreateDO, "colReject_ChkCreateDO")
        Me.colReject_ChkCreateDO.Name = "colReject_ChkCreateDO"
        '
        'colReject_Remarks
        '
        resources.ApplyResources(Me.colReject_Remarks, "colReject_Remarks")
        Me.colReject_Remarks.Name = "colReject_Remarks"
        '
        'colReject_PrimaryProduct
        '
        resources.ApplyResources(Me.colReject_PrimaryProduct, "colReject_PrimaryProduct")
        Me.colReject_PrimaryProduct.Name = "colReject_PrimaryProduct"
        '
        'colReject_PrimaryUnitMulti
        '
        resources.ApplyResources(Me.colReject_PrimaryUnitMulti, "colReject_PrimaryUnitMulti")
        Me.colReject_PrimaryUnitMulti.Name = "colReject_PrimaryUnitMulti"
        '
        'colReject_PrimaryTransaction
        '
        resources.ApplyResources(Me.colReject_PrimaryTransaction, "colReject_PrimaryTransaction")
        Me.colReject_PrimaryTransaction.Name = "colReject_PrimaryTransaction"
        '
        'colReject_PrimaryCustomer
        '
        resources.ApplyResources(Me.colReject_PrimaryCustomer, "colReject_PrimaryCustomer")
        Me.colReject_PrimaryCustomer.Name = "colReject_PrimaryCustomer"
        Me.colReject_PrimaryCustomer.ReadOnly = True
        '
        'colReject_PrimaryLocation
        '
        resources.ApplyResources(Me.colReject_PrimaryLocation, "colReject_PrimaryLocation")
        Me.colReject_PrimaryLocation.Name = "colReject_PrimaryLocation"
        Me.colReject_PrimaryLocation.ReadOnly = True
        '
        'colReject_PrimaryWarehouse
        '
        resources.ApplyResources(Me.colReject_PrimaryWarehouse, "colReject_PrimaryWarehouse")
        Me.colReject_PrimaryWarehouse.Name = "colReject_PrimaryWarehouse"
        Me.colReject_PrimaryWarehouse.ReadOnly = True
        '
        'colReject_PrimaryCurrency
        '
        resources.ApplyResources(Me.colReject_PrimaryCurrency, "colReject_PrimaryCurrency")
        Me.colReject_PrimaryCurrency.Name = "colReject_PrimaryCurrency"
        Me.colReject_PrimaryCurrency.ReadOnly = True
        '
        'colReject_QtyDefault
        '
        resources.ApplyResources(Me.colReject_QtyDefault, "colReject_QtyDefault")
        Me.colReject_QtyDefault.Name = "colReject_QtyDefault"
        '
        'colReject_PrimaryUnitDefault
        '
        resources.ApplyResources(Me.colReject_PrimaryUnitDefault, "colReject_PrimaryUnitDefault")
        Me.colReject_PrimaryUnitDefault.Name = "colReject_PrimaryUnitDefault"
        '
        'colReject_RatioDefault
        '
        resources.ApplyResources(Me.colReject_RatioDefault, "colReject_RatioDefault")
        Me.colReject_RatioDefault.Name = "colReject_RatioDefault"
        '
        'colReject_RatioReference
        '
        resources.ApplyResources(Me.colReject_RatioReference, "colReject_RatioReference")
        Me.colReject_RatioReference.Name = "colReject_RatioReference"
        '
        'colReject_DefprecDefault
        '
        resources.ApplyResources(Me.colReject_DefprecDefault, "colReject_DefprecDefault")
        Me.colReject_DefprecDefault.Name = "colReject_DefprecDefault"
        '
        'colReject_UnitDefault
        '
        resources.ApplyResources(Me.colReject_UnitDefault, "colReject_UnitDefault")
        Me.colReject_UnitDefault.Name = "colReject_UnitDefault"
        '
        'colReject_DecimalMulti
        '
        resources.ApplyResources(Me.colReject_DecimalMulti, "colReject_DecimalMulti")
        Me.colReject_DecimalMulti.Name = "colReject_DecimalMulti"
        Me.colReject_DecimalMulti.ReadOnly = True
        '
        'colReject_IsBatch
        '
        resources.ApplyResources(Me.colReject_IsBatch, "colReject_IsBatch")
        Me.colReject_IsBatch.Name = "colReject_IsBatch"
        Me.colReject_IsBatch.ReadOnly = True
        '
        'colReject_IsExpiry
        '
        resources.ApplyResources(Me.colReject_IsExpiry, "colReject_IsExpiry")
        Me.colReject_IsExpiry.Name = "colReject_IsExpiry"
        Me.colReject_IsExpiry.ReadOnly = True
        '
        'colReject_IsLocation
        '
        resources.ApplyResources(Me.colReject_IsLocation, "colReject_IsLocation")
        Me.colReject_IsLocation.Name = "colReject_IsLocation"
        Me.colReject_IsLocation.ReadOnly = True
        '
        'colReject_IdentifierCustAddress
        '
        resources.ApplyResources(Me.colReject_IdentifierCustAddress, "colReject_IdentifierCustAddress")
        Me.colReject_IdentifierCustAddress.Name = "colReject_IdentifierCustAddress"
        '
        'colReject_Tax
        '
        resources.ApplyResources(Me.colReject_Tax, "colReject_Tax")
        Me.colReject_Tax.Name = "colReject_Tax"
        '
        'colReject_FakturPajak
        '
        resources.ApplyResources(Me.colReject_FakturPajak, "colReject_FakturPajak")
        Me.colReject_FakturPajak.Name = "colReject_FakturPajak"
        '
        'colReject_IdentifierAddressCustomerShipTo
        '
        resources.ApplyResources(Me.colReject_IdentifierAddressCustomerShipTo, "colReject_IdentifierAddressCustomerShipTo")
        Me.colReject_IdentifierAddressCustomerShipTo.Name = "colReject_IdentifierAddressCustomerShipTo"
        '
        'colReject_IdentifierAddressShipTo
        '
        resources.ApplyResources(Me.colReject_IdentifierAddressShipTo, "colReject_IdentifierAddressShipTo")
        Me.colReject_IdentifierAddressShipTo.Name = "colReject_IdentifierAddressShipTo"
        '
        'colReject_PrimaryCustomerShipTo
        '
        resources.ApplyResources(Me.colReject_PrimaryCustomerShipTo, "colReject_PrimaryCustomerShipTo")
        Me.colReject_PrimaryCustomerShipTo.Name = "colReject_PrimaryCustomerShipTo"
        '
        'colReject_PrimarySO
        '
        resources.ApplyResources(Me.colReject_PrimarySO, "colReject_PrimarySO")
        Me.colReject_PrimarySO.Name = "colReject_PrimarySO"
        '
        'colReject_SONo
        '
        resources.ApplyResources(Me.colReject_SONo, "colReject_SONo")
        Me.colReject_SONo.Name = "colReject_SONo"
        '
        'FrmTransaksiDO_Quick
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.prg01)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.dgvItemReject)
        Me.Name = "FrmTransaksiDO_Quick"
        Me.ShowInTaskbar = False
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvCreatedBy, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.DgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvItemProductGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GBPeriod.ResumeLayout(False)
        Me.GBPeriod.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemReject, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblFind As System.Windows.Forms.Label
    Friend WithEvents TxtTransactionNumber As System.Windows.Forms.TextBox
    Friend WithEvents btnFindTransaction As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryTransaction As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents GBPeriod As System.Windows.Forms.GroupBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents prg01 As System.Windows.Forms.ProgressBar
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents dgvCreatedBy As System.Windows.Forms.DataGridView
    Friend WithEvents ChkIgnoreLowTolerance As System.Windows.Forms.CheckBox
    Friend WithEvents cboCustomerFormat As System.Windows.Forms.ComboBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtCancellationReason As System.Windows.Forms.TextBox
    Friend WithEvents chkDirectSave As System.Windows.Forms.CheckBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents LlbClearFilter As System.Windows.Forms.LinkLabel
    Friend WithEvents DgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents dgvItemProductGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColProductGroup_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Level As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColProductGroup_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductGroup_PrimaryAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BgwCustomer As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColUserNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUser_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnFindLoc As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtLocCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryLoc As System.Windows.Forms.TextBox
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColCust_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCust_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents rbInventoryBooking As System.Windows.Forms.RadioButton
    Friend WithEvents rbSO As System.Windows.Forms.RadioButton
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDocumentNO As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTolerance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyDO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColChkDO As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColRemarksItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPoduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryTransaction As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaruUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefPrec_QtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimalMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchControlled As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryControlled As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocationControlled As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier_CustAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFakturPajak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Colidentifier_address_customer_ship As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Colidentifier_address_ship_to As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary_Customer_Ship_To As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents dgvItemReject As System.Windows.Forms.DataGridView
    Friend WithEvents colReject_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_CustomerCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_DocumentNo As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents colReject_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_ExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_Plus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_Currency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_UnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_DiscPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_DiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_QtyDO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_ChkCreateDO As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colReject_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryTransaction As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryCustomer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_QtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_RatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_RatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_DefprecDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_UnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_DecimalMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_IsBatch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_IsExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_IsLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_IdentifierCustAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_Tax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_FakturPajak As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_IdentifierAddressCustomerShipTo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_IdentifierAddressShipTo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimaryCustomerShipTo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_PrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReject_SONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtPrimaryWarehouse As System.Windows.Forms.TextBox
    Friend WithEvents txtUserCode As System.Windows.Forms.TextBox
    Friend WithEvents btnFindUser As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryUser As System.Windows.Forms.TextBox
End Class
