<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard
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
        Me.components = New System.ComponentModel.Container
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvMP = New System.Windows.Forms.DataGridView
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TxtFilter_PrimaryMO = New System.Windows.Forms.TextBox
        Me.TxtFilter_MONo = New System.Windows.Forms.TextBox
        Me.LblFilter_FindMO = New System.Windows.Forms.Label
        Me.btnFilter_FindMO = New System.Windows.Forms.Button
        Me.btnFilter_FindCodeNo = New System.Windows.Forms.Button
        Me.TxtFilter_PrimaryPurchReq = New System.Windows.Forms.TextBox
        Me.TxtFilter_PurcReqNo = New System.Windows.Forms.TextBox
        Me.LblFilter_PurcReq = New System.Windows.Forms.Label
        Me.btnFilter_FindPurchReq = New System.Windows.Forms.Button
        Me.GboFilter_PurcReqDate = New System.Windows.Forms.GroupBox
        Me.chkFilter_PurhReqDate = New System.Windows.Forms.CheckBox
        Me.DtpFilter_PurchReq_To = New System.Windows.Forms.DateTimePicker
        Me.Label34 = New System.Windows.Forms.Label
        Me.DtpFilter_PurcReq_From = New System.Windows.Forms.DateTimePicker
        Me.Label35 = New System.Windows.Forms.Label
        Me.Txtfilter_PrimaryRefParty = New System.Windows.Forms.TextBox
        Me.Txtfilter_RefPartyCode = New System.Windows.Forms.TextBox
        Me.LblRefParty = New System.Windows.Forms.Label
        Me.chkFilter_PrOBItem = New System.Windows.Forms.CheckBox
        Me.chkFilter_SOBItem = New System.Windows.Forms.CheckBox
        Me.chkFilter_FollowUpOnly = New System.Windows.Forms.CheckBox
        Me.chkFilter_Service = New System.Windows.Forms.CheckBox
        Me.chkFilter_Goods = New System.Windows.Forms.CheckBox
        Me.TxtMP_QtyInquiry = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TxtFilter_PrimaryCustomer = New System.Windows.Forms.TextBox
        Me.TxtFilter_PrimaryProductGroup = New System.Windows.Forms.TextBox
        Me.TxtFilter_PrimaryPreferedSupplier = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.llbMP_Filter = New System.Windows.Forms.LinkLabel
        Me.TxtFilter_ProductGroupCode = New System.Windows.Forms.TextBox
        Me.LblFilter_PoductGroupDesc = New System.Windows.Forms.Label
        Me.TxtFilter_Desc = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtFilter_CodeNo = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.LblFilter_UserCreateSO = New System.Windows.Forms.Label
        Me.TxtFilter_CustCode = New System.Windows.Forms.TextBox
        Me.LblFilter_Cust = New System.Windows.Forms.Label
        Me.cboFilter_ProductGroupLevel = New System.Windows.Forms.ComboBox
        Me.LblFilter_ProductGroup = New System.Windows.Forms.Label
        Me.btnFilter_FindPreferedSupplier = New System.Windows.Forms.Button
        Me.TxtFilter_PreferedSupplier = New System.Windows.Forms.TextBox
        Me.LblFilter_PreferedSupplier = New System.Windows.Forms.Label
        Me.TxtFilter_PreferedSupplierDesc = New System.Windows.Forms.TextBox
        Me.btnFilter_FindRefParty = New System.Windows.Forms.Button
        Me.Txtfilter_RefPartyName = New System.Windows.Forms.TextBox
        Me.btnFilter_FindProductGroupLevel = New System.Windows.Forms.Button
        Me.btnFilter_FindCust = New System.Windows.Forms.Button
        Me.TxtFilter_ProductGroupLevelDesc = New System.Windows.Forms.TextBox
        Me.TxtFilter_CustName = New System.Windows.Forms.TextBox
        Me.TxtFilter_UserCreateSO = New System.Windows.Forms.TextBox
        Me.btnFilter_FindUser = New System.Windows.Forms.Button
        Me.LblFilter_ProcurementStrategy = New System.Windows.Forms.Label
        Me.GboFilter_MODate = New System.Windows.Forms.GroupBox
        Me.chkMODateFilter = New System.Windows.Forms.CheckBox
        Me.DtpFilter_MO_To = New System.Windows.Forms.DateTimePicker
        Me.Label37 = New System.Windows.Forms.Label
        Me.DtpFilter_MO_From = New System.Windows.Forms.DateTimePicker
        Me.Label38 = New System.Windows.Forms.Label
        Me.GboFilter_SODate = New System.Windows.Forms.GroupBox
        Me.chkSODateFilter = New System.Windows.Forms.CheckBox
        Me.DtpFilter_To = New System.Windows.Forms.DateTimePicker
        Me.Label21 = New System.Windows.Forms.Label
        Me.DtpFilter_From = New System.Windows.Forms.DateTimePicker
        Me.Label7 = New System.Windows.Forms.Label
        Me.CboFilter_ProcurementStrategy = New System.Windows.Forms.ComboBox
        Me.TxtFilter_NamaUserCreateSO = New System.Windows.Forms.TextBox
        Me.chkFilter_PrOBLargerThanPOB = New System.Windows.Forms.CheckBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.chkMP_RequiredOnly = New System.Windows.Forms.CheckBox
        Me.btnMP_Print = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.btnMP_Edit = New System.Windows.Forms.Button
        Me.btnMP_New = New System.Windows.Forms.Button
        Me.StatusStrip4 = New System.Windows.Forms.StatusStrip
        Me.tslMP_CreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslMP_CreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslMP_ModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslMP_ModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboMP = New System.Windows.Forms.GroupBox
        Me.llmbMP_SetQtyOrderEqualQtyPrOB = New System.Windows.Forms.LinkLabel
        Me.llbMP_CopyMP = New System.Windows.Forms.LinkLabel
        Me.llbMP_CekAll = New System.Windows.Forms.LinkLabel
        Me.llbMP_AddToPO = New System.Windows.Forms.LinkLabel
        Me.dtpMP_Filter = New System.Windows.Forms.DateTimePicker
        Me.llbMP_CreatePO = New System.Windows.Forms.LinkLabel
        Me.btnMP_Cancel = New System.Windows.Forms.Button
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer
        Me.TxtMP_DaysPerMonth = New System.Windows.Forms.TextBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.dtpMP_RefDate = New System.Windows.Forms.DateTimePicker
        Me.dtpMP = New System.Windows.Forms.DateTimePicker
        Me.llbMP_Calculate = New System.Windows.Forms.LinkLabel
        Me.TxtMP_Primary = New System.Windows.Forms.TextBox
        Me.TxtMP_DocNo = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.llbAuthDetail = New System.Windows.Forms.LinkLabel
        Me.llbMP_LocationSetting = New System.Windows.Forms.LinkLabel
        Me.TxtMP_Remarks = New System.Windows.Forms.TextBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.chkClose = New System.Windows.Forms.CheckBox
        Me.btnMP_OK = New System.Windows.Forms.Button
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Label42 = New System.Windows.Forms.Label
        Me.TxtMP_QtyRMReqSpare = New System.Windows.Forms.TextBox
        Me.Label41 = New System.Windows.Forms.Label
        Me.TxtMP_QtySOBSpare = New System.Windows.Forms.TextBox
        Me.Label40 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPrOBSpare = New System.Windows.Forms.TextBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.TxtMP_QtyRMReq = New System.Windows.Forms.TextBox
        Me.Label31 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPrOB = New System.Windows.Forms.TextBox
        Me.TxtMP_QtySafetyStock = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.TxtMP_QtySOB = New System.Windows.Forms.TextBox
        Me.TxtMP_QtyOptimumStock = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPendingDO = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TxtMP_QtyFree = New System.Windows.Forms.TextBox
        Me.Label39 = New System.Windows.Forms.Label
        Me.TxtMP_QtyMOB = New System.Windows.Forms.TextBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtMP_QtyRequired = New System.Windows.Forms.TextBox
        Me.TxtMP_QtyOnHand = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.TxtMP_QtyPOB = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TxtMP_QtyGIT = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtMP_AvgPerMonth = New System.Windows.Forms.TextBox
        Me.cboMPOrderStrategy = New System.Windows.Forms.ComboBox
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.TxtMP_DiscFromPerct = New System.Windows.Forms.TextBox
        Me.TxtMP_DiscAmount = New System.Windows.Forms.TextBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.LlbMP_UnitPrice = New System.Windows.Forms.LinkLabel
        Me.TxtMP_CurrencyCode = New System.Windows.Forms.TextBox
        Me.TxtMP_UnitPrice = New System.Windows.Forms.TextBox
        Me.btnMP_FindCurrency = New System.Windows.Forms.Button
        Me.Label22 = New System.Windows.Forms.Label
        Me.TxtMP_PreferedSupplier = New System.Windows.Forms.TextBox
        Me.TxtMP_Discount = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.TxtMP_NettPrice = New System.Windows.Forms.TextBox
        Me.btnMP_FindSupplier = New System.Windows.Forms.Button
        Me.TxtMP_PrimaryCurrency = New System.Windows.Forms.TextBox
        Me.TxtMP_PrimarySupplier = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.llbMP_Tolerance = New System.Windows.Forms.LinkLabel
        Me.TxtMP_PreferedSupplierDesc = New System.Windows.Forms.TextBox
        Me.llbMP_Ratio = New System.Windows.Forms.LinkLabel
        Me.TxtMP_QtyRequired2 = New System.Windows.Forms.TextBox
        Me.LblMP_Unit1 = New System.Windows.Forms.Label
        Me.TxtMP_QtyOrder = New System.Windows.Forms.TextBox
        Me.LblRatio = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.TxtMP_SubTotal = New System.Windows.Forms.TextBox
        Me.TxtMP_PrimaryOrderUnit = New System.Windows.Forms.TextBox
        Me.cboMP_OrderUnit = New System.Windows.Forms.ComboBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.btnMP_Find = New System.Windows.Forms.Button
        Me.btnMP_Delete = New System.Windows.Forms.Button
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.btnSF_Edit = New System.Windows.Forms.Button
        Me.btnSF_New = New System.Windows.Forms.Button
        Me.btnSF_Find = New System.Windows.Forms.Button
        Me.btnSF_Print = New System.Windows.Forms.Button
        Me.btnSF_Delete = New System.Windows.Forms.Button
        Me.StatusStrip3 = New System.Windows.Forms.StatusStrip
        Me.tslSF_CreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslSF_CreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslSF_ModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslSF_ModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboSF = New System.Windows.Forms.GroupBox
        Me.llbSF_CekAll = New System.Windows.Forms.LinkLabel
        Me.btnSF_Cancel = New System.Windows.Forms.Button
        Me.btnSF_OK = New System.Windows.Forms.Button
        Me.llbSF_CreateDO = New System.Windows.Forms.LinkLabel
        Me.dgvSF = New System.Windows.Forms.DataGridView
        Me.ColSF_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_Type = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_SONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_CustomerPONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_SODate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_Customer = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_QtySOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_QtyAvailable = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_PromiseToDeliver = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_PromiseToDeliver_Str = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_EstimateDelivery = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_EstimateDelivery_Str = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColSF_CreateDO = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColSF_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_DeliveryStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_PrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_PrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_PrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_QtyReadyStock = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_TypeInt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSF_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer6 = New System.Windows.Forms.SplitContainer
        Me.TxtSF_StockInPurchaseNo = New System.Windows.Forms.TextBox
        Me.LblSF_StockInPurchaseNo = New System.Windows.Forms.Label
        Me.llbSF_Calculate = New System.Windows.Forms.LinkLabel
        Me.TxtSF_Urutan = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpSF = New System.Windows.Forms.DateTimePicker
        Me.TxtSF_Primary = New System.Windows.Forms.TextBox
        Me.TxtSF_DocNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.dtpAvailable = New System.Windows.Forms.DateTimePicker
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.chkSF_LateOnly = New System.Windows.Forms.CheckBox
        Me.TxtSF_Remarks = New System.Windows.Forms.TextBox
        Me.cboAvailability = New System.Windows.Forms.ComboBox
        Me.llbColumnSettings = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.bwRefDate = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.BackgroundWorker6 = New System.ComponentModel.BackgroundWorker
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.ColMP_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOnHand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyPOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyGIT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyInquiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyPendingDO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtySOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyPrOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtySafetyStock = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyLeadTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyRequired = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColMP_QtyOrderMaximum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrderMinimum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrderOptimum = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyMax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyMin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOpt = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrder = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_OrderUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_OrderPoint = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColMP_NextOrderPoint = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColMP_ProcurementType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMP_ProcurementStrategy = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMP_OrderStrategy = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColMP_FollowUp = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColMP_PreferedSupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Currency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_UnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_Disc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_NettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_SubTotal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimarySupplier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryOrderUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecOrderUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_OrderPointDelay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_NextOrderPointDelay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_LeadTimeDays = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_QtyOrderMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_PrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DefPrecUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_RatioMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_RatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMP_DetailFollowUp = New System.Windows.Forms.DataGridViewLinkColumn
        CType(Me.dgvMP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GboFilter_PurcReqDate.SuspendLayout()
        Me.GboFilter_MODate.SuspendLayout()
        Me.GboFilter_SODate.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.StatusStrip4.SuspendLayout()
        Me.gboMP.SuspendLayout()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.StatusStrip3.SuspendLayout()
        Me.gboSF.SuspendLayout()
        CType(Me.dgvSF, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer6.Panel1.SuspendLayout()
        Me.SplitContainer6.Panel2.SuspendLayout()
        Me.SplitContainer6.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvMP
        '
        Me.dgvMP.AllowUserToAddRows = False
        Me.dgvMP.AllowUserToDeleteRows = False
        Me.dgvMP.AllowUserToResizeColumns = False
        Me.dgvMP.AllowUserToResizeRows = False
        Me.dgvMP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvMP.ColumnHeadersHeight = 25
        Me.dgvMP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvMP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColMP_No, Me.ColMP_CodeNo, Me.ColMP_Desc, Me.ColMP_Unit, Me.ColMP_QtyOnHand, Me.ColMP_QtyPOB, Me.ColMP_QtyGIT, Me.ColMP_QtyInquiry, Me.ColMP_QtyPendingDO, Me.ColMP_QtySOB, Me.ColMP_QtyPrOB, Me.ColMP_QtySafetyStock, Me.ColMP_QtyLeadTime, Me.ColMP_QtyRequired, Me.ColMP_QtyOrderMaximum, Me.ColMP_QtyOrderMinimum, Me.ColMP_QtyOrderOptimum, Me.ColMP_QtyMax, Me.ColMP_QtyMin, Me.ColMP_QtyOpt, Me.ColMP_QtyOrder, Me.ColMP_OrderUnit, Me.ColMP_OrderPoint, Me.ColMP_NextOrderPoint, Me.ColMP_ProcurementType, Me.ColMP_ProcurementStrategy, Me.ColMP_OrderStrategy, Me.ColMP_FollowUp, Me.ColMP_PreferedSupplier, Me.ColMP_Currency, Me.ColMP_UnitPrice, Me.ColMP_Disc, Me.ColMP_NettPrice, Me.ColMP_SubTotal, Me.ColMP_PrimaryCurrency, Me.ColMP_PrimarySupplier, Me.ColMP_PrimaryCodeNo, Me.ColMP_DefPrecCurrency, Me.ColMP_DefPrecUnit, Me.ColMP_PrimaryOrderUnit, Me.ColMP_DefPrecOrderUnit, Me.ColMP_OrderPointDelay, Me.ColMP_NextOrderPointDelay, Me.ColMP_LeadTimeDays, Me.ColMP_QtyOrderMulti, Me.ColMP_PrimaryUnitMulti, Me.ColMP_DefPrecUnitMulti, Me.ColMP_RatioMulti, Me.ColMP_RatioDefault, Me.ColMP_DiscAmount, Me.ColMP_DiscFromPerct, Me.ColMP_DetailFollowUp})
        Me.dgvMP.Location = New System.Drawing.Point(2, 84)
        Me.dgvMP.MultiSelect = False
        Me.dgvMP.Name = "dgvMP"
        Me.dgvMP.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvMP.RowTemplate.Height = 24
        Me.dgvMP.Size = New System.Drawing.Size(569, 139)
        Me.dgvMP.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PrimaryMO)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_MONo)
        Me.GroupBox1.Controls.Add(Me.LblFilter_FindMO)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindMO)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindCodeNo)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PrimaryPurchReq)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PurcReqNo)
        Me.GroupBox1.Controls.Add(Me.LblFilter_PurcReq)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindPurchReq)
        Me.GroupBox1.Controls.Add(Me.GboFilter_PurcReqDate)
        Me.GroupBox1.Controls.Add(Me.Txtfilter_PrimaryRefParty)
        Me.GroupBox1.Controls.Add(Me.Txtfilter_RefPartyCode)
        Me.GroupBox1.Controls.Add(Me.LblRefParty)
        Me.GroupBox1.Controls.Add(Me.chkFilter_PrOBItem)
        Me.GroupBox1.Controls.Add(Me.chkFilter_SOBItem)
        Me.GroupBox1.Controls.Add(Me.chkFilter_FollowUpOnly)
        Me.GroupBox1.Controls.Add(Me.chkFilter_Service)
        Me.GroupBox1.Controls.Add(Me.chkFilter_Goods)
        Me.GroupBox1.Controls.Add(Me.TxtMP_QtyInquiry)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PrimaryCustomer)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PrimaryProductGroup)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PrimaryPreferedSupplier)
        Me.GroupBox1.Controls.Add(Me.TxtReadOnly)
        Me.GroupBox1.Controls.Add(Me.llbMP_Filter)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_ProductGroupCode)
        Me.GroupBox1.Controls.Add(Me.LblFilter_PoductGroupDesc)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_Desc)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_CodeNo)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.LblFilter_UserCreateSO)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_CustCode)
        Me.GroupBox1.Controls.Add(Me.LblFilter_Cust)
        Me.GroupBox1.Controls.Add(Me.cboFilter_ProductGroupLevel)
        Me.GroupBox1.Controls.Add(Me.LblFilter_ProductGroup)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindPreferedSupplier)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PreferedSupplier)
        Me.GroupBox1.Controls.Add(Me.LblFilter_PreferedSupplier)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_PreferedSupplierDesc)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindRefParty)
        Me.GroupBox1.Controls.Add(Me.Txtfilter_RefPartyName)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindProductGroupLevel)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindCust)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_ProductGroupLevelDesc)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_CustName)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_UserCreateSO)
        Me.GroupBox1.Controls.Add(Me.btnFilter_FindUser)
        Me.GroupBox1.Controls.Add(Me.LblFilter_ProcurementStrategy)
        Me.GroupBox1.Controls.Add(Me.GboFilter_MODate)
        Me.GroupBox1.Controls.Add(Me.GboFilter_SODate)
        Me.GroupBox1.Controls.Add(Me.CboFilter_ProcurementStrategy)
        Me.GroupBox1.Controls.Add(Me.TxtFilter_NamaUserCreateSO)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(227, 515)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Data Filtering"
        '
        'TxtFilter_PrimaryMO
        '
        Me.TxtFilter_PrimaryMO.Location = New System.Drawing.Point(22, 530)
        Me.TxtFilter_PrimaryMO.Name = "TxtFilter_PrimaryMO"
        Me.TxtFilter_PrimaryMO.Size = New System.Drawing.Size(22, 20)
        Me.TxtFilter_PrimaryMO.TabIndex = 113
        Me.TxtFilter_PrimaryMO.TabStop = False
        Me.TxtFilter_PrimaryMO.Text = "-1"
        Me.TxtFilter_PrimaryMO.Visible = False
        '
        'TxtFilter_MONo
        '
        Me.TxtFilter_MONo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_MONo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_MONo.Location = New System.Drawing.Point(72, 530)
        Me.TxtFilter_MONo.Name = "TxtFilter_MONo"
        Me.TxtFilter_MONo.Size = New System.Drawing.Size(97, 20)
        Me.TxtFilter_MONo.TabIndex = 114
        '
        'LblFilter_FindMO
        '
        Me.LblFilter_FindMO.AutoSize = True
        Me.LblFilter_FindMO.Location = New System.Drawing.Point(7, 531)
        Me.LblFilter_FindMO.Name = "LblFilter_FindMO"
        Me.LblFilter_FindMO.Size = New System.Drawing.Size(41, 13)
        Me.LblFilter_FindMO.TabIndex = 116
        Me.LblFilter_FindMO.Text = "MO No"
        '
        'btnFilter_FindMO
        '
        Me.btnFilter_FindMO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindMO.Location = New System.Drawing.Point(177, 530)
        Me.btnFilter_FindMO.Name = "btnFilter_FindMO"
        Me.btnFilter_FindMO.Size = New System.Drawing.Size(39, 18)
        Me.btnFilter_FindMO.TabIndex = 115
        Me.btnFilter_FindMO.TabStop = False
        Me.btnFilter_FindMO.Text = "Find"
        Me.btnFilter_FindMO.UseVisualStyleBackColor = True
        '
        'btnFilter_FindCodeNo
        '
        Me.btnFilter_FindCodeNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindCodeNo.Location = New System.Drawing.Point(177, 87)
        Me.btnFilter_FindCodeNo.Name = "btnFilter_FindCodeNo"
        Me.btnFilter_FindCodeNo.Size = New System.Drawing.Size(39, 23)
        Me.btnFilter_FindCodeNo.TabIndex = 111
        Me.btnFilter_FindCodeNo.TabStop = False
        Me.btnFilter_FindCodeNo.Text = "Find"
        Me.btnFilter_FindCodeNo.UseVisualStyleBackColor = True
        '
        'TxtFilter_PrimaryPurchReq
        '
        Me.TxtFilter_PrimaryPurchReq.Location = New System.Drawing.Point(22, 505)
        Me.TxtFilter_PrimaryPurchReq.Name = "TxtFilter_PrimaryPurchReq"
        Me.TxtFilter_PrimaryPurchReq.Size = New System.Drawing.Size(22, 20)
        Me.TxtFilter_PrimaryPurchReq.TabIndex = 107
        Me.TxtFilter_PrimaryPurchReq.TabStop = False
        Me.TxtFilter_PrimaryPurchReq.Text = "-1"
        Me.TxtFilter_PrimaryPurchReq.Visible = False
        '
        'TxtFilter_PurcReqNo
        '
        Me.TxtFilter_PurcReqNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_PurcReqNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_PurcReqNo.Location = New System.Drawing.Point(72, 505)
        Me.TxtFilter_PurcReqNo.Name = "TxtFilter_PurcReqNo"
        Me.TxtFilter_PurcReqNo.Size = New System.Drawing.Size(97, 20)
        Me.TxtFilter_PurcReqNo.TabIndex = 108
        '
        'LblFilter_PurcReq
        '
        Me.LblFilter_PurcReq.AutoSize = True
        Me.LblFilter_PurcReq.Location = New System.Drawing.Point(7, 507)
        Me.LblFilter_PurcReq.Name = "LblFilter_PurcReq"
        Me.LblFilter_PurcReq.Size = New System.Drawing.Size(61, 13)
        Me.LblFilter_PurcReq.TabIndex = 110
        Me.LblFilter_PurcReq.Text = "Purch. Req"
        '
        'btnFilter_FindPurchReq
        '
        Me.btnFilter_FindPurchReq.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindPurchReq.Location = New System.Drawing.Point(177, 505)
        Me.btnFilter_FindPurchReq.Name = "btnFilter_FindPurchReq"
        Me.btnFilter_FindPurchReq.Size = New System.Drawing.Size(39, 18)
        Me.btnFilter_FindPurchReq.TabIndex = 109
        Me.btnFilter_FindPurchReq.TabStop = False
        Me.btnFilter_FindPurchReq.Text = "Find"
        Me.btnFilter_FindPurchReq.UseVisualStyleBackColor = True
        '
        'GboFilter_PurcReqDate
        '
        Me.GboFilter_PurcReqDate.Controls.Add(Me.chkFilter_PurhReqDate)
        Me.GboFilter_PurcReqDate.Controls.Add(Me.DtpFilter_PurchReq_To)
        Me.GboFilter_PurcReqDate.Controls.Add(Me.Label34)
        Me.GboFilter_PurcReqDate.Controls.Add(Me.DtpFilter_PurcReq_From)
        Me.GboFilter_PurcReqDate.Controls.Add(Me.Label35)
        Me.GboFilter_PurcReqDate.Location = New System.Drawing.Point(4, 436)
        Me.GboFilter_PurcReqDate.Name = "GboFilter_PurcReqDate"
        Me.GboFilter_PurcReqDate.Size = New System.Drawing.Size(219, 65)
        Me.GboFilter_PurcReqDate.TabIndex = 21
        Me.GboFilter_PurcReqDate.TabStop = False
        '
        'chkFilter_PurhReqDate
        '
        Me.chkFilter_PurhReqDate.AutoSize = True
        Me.chkFilter_PurhReqDate.Location = New System.Drawing.Point(10, -1)
        Me.chkFilter_PurhReqDate.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFilter_PurhReqDate.Name = "chkFilter_PurhReqDate"
        Me.chkFilter_PurhReqDate.Size = New System.Drawing.Size(109, 17)
        Me.chkFilter_PurhReqDate.TabIndex = 0
        Me.chkFilter_PurhReqDate.Text = "Purch. Req. Date"
        Me.chkFilter_PurhReqDate.UseVisualStyleBackColor = True
        '
        'DtpFilter_PurchReq_To
        '
        Me.DtpFilter_PurchReq_To.CustomFormat = "dd MMM yyyy"
        Me.DtpFilter_PurchReq_To.Enabled = False
        Me.DtpFilter_PurchReq_To.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFilter_PurchReq_To.Location = New System.Drawing.Point(82, 41)
        Me.DtpFilter_PurchReq_To.Name = "DtpFilter_PurchReq_To"
        Me.DtpFilter_PurchReq_To.Size = New System.Drawing.Size(130, 20)
        Me.DtpFilter_PurchReq_To.TabIndex = 2
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Enabled = False
        Me.Label34.Location = New System.Drawing.Point(8, 46)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(20, 13)
        Me.Label34.TabIndex = 2
        Me.Label34.Text = "To"
        '
        'DtpFilter_PurcReq_From
        '
        Me.DtpFilter_PurcReq_From.CustomFormat = "dd MMM yyyy"
        Me.DtpFilter_PurcReq_From.Enabled = False
        Me.DtpFilter_PurcReq_From.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFilter_PurcReq_From.Location = New System.Drawing.Point(82, 18)
        Me.DtpFilter_PurcReq_From.Name = "DtpFilter_PurcReq_From"
        Me.DtpFilter_PurcReq_From.Size = New System.Drawing.Size(130, 20)
        Me.DtpFilter_PurcReq_From.TabIndex = 1
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Enabled = False
        Me.Label35.Location = New System.Drawing.Point(8, 22)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(30, 13)
        Me.Label35.TabIndex = 0
        Me.Label35.Text = "From"
        '
        'Txtfilter_PrimaryRefParty
        '
        Me.Txtfilter_PrimaryRefParty.Location = New System.Drawing.Point(75, 389)
        Me.Txtfilter_PrimaryRefParty.Name = "Txtfilter_PrimaryRefParty"
        Me.Txtfilter_PrimaryRefParty.Size = New System.Drawing.Size(22, 20)
        Me.Txtfilter_PrimaryRefParty.TabIndex = 106
        Me.Txtfilter_PrimaryRefParty.TabStop = False
        Me.Txtfilter_PrimaryRefParty.Text = "-1"
        Me.Txtfilter_PrimaryRefParty.Visible = False
        '
        'Txtfilter_RefPartyCode
        '
        Me.Txtfilter_RefPartyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txtfilter_RefPartyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.Txtfilter_RefPartyCode.Location = New System.Drawing.Point(102, 388)
        Me.Txtfilter_RefPartyCode.Name = "Txtfilter_RefPartyCode"
        Me.Txtfilter_RefPartyCode.Size = New System.Drawing.Size(69, 20)
        Me.Txtfilter_RefPartyCode.TabIndex = 19
        '
        'LblRefParty
        '
        Me.LblRefParty.AutoSize = True
        Me.LblRefParty.Location = New System.Drawing.Point(7, 392)
        Me.LblRefParty.Name = "LblRefParty"
        Me.LblRefParty.Size = New System.Drawing.Size(54, 13)
        Me.LblRefParty.TabIndex = 105
        Me.LblRefParty.Text = "Ref. Party"
        '
        'chkFilter_PrOBItem
        '
        Me.chkFilter_PrOBItem.AutoSize = True
        Me.chkFilter_PrOBItem.Checked = True
        Me.chkFilter_PrOBItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_PrOBItem.Location = New System.Drawing.Point(136, 63)
        Me.chkFilter_PrOBItem.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFilter_PrOBItem.Name = "chkFilter_PrOBItem"
        Me.chkFilter_PrOBItem.Size = New System.Drawing.Size(93, 17)
        Me.chkFilter_PrOBItem.TabIndex = 4
        Me.chkFilter_PrOBItem.Text = "Item w/ Pr.OB"
        Me.chkFilter_PrOBItem.UseVisualStyleBackColor = True
        '
        'chkFilter_SOBItem
        '
        Me.chkFilter_SOBItem.AutoSize = True
        Me.chkFilter_SOBItem.Checked = True
        Me.chkFilter_SOBItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_SOBItem.Location = New System.Drawing.Point(136, 41)
        Me.chkFilter_SOBItem.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFilter_SOBItem.Name = "chkFilter_SOBItem"
        Me.chkFilter_SOBItem.Size = New System.Drawing.Size(87, 17)
        Me.chkFilter_SOBItem.TabIndex = 3
        Me.chkFilter_SOBItem.Text = "Item w/ SOB"
        Me.chkFilter_SOBItem.UseVisualStyleBackColor = True
        '
        'chkFilter_FollowUpOnly
        '
        Me.chkFilter_FollowUpOnly.AutoSize = True
        Me.chkFilter_FollowUpOnly.Location = New System.Drawing.Point(9, 41)
        Me.chkFilter_FollowUpOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFilter_FollowUpOnly.Name = "chkFilter_FollowUpOnly"
        Me.chkFilter_FollowUpOnly.Size = New System.Drawing.Size(97, 17)
        Me.chkFilter_FollowUpOnly.TabIndex = 2
        Me.chkFilter_FollowUpOnly.Text = "Follow Up Only"
        Me.chkFilter_FollowUpOnly.UseVisualStyleBackColor = True
        '
        'chkFilter_Service
        '
        Me.chkFilter_Service.AutoSize = True
        Me.chkFilter_Service.Checked = True
        Me.chkFilter_Service.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_Service.Location = New System.Drawing.Point(135, 19)
        Me.chkFilter_Service.Name = "chkFilter_Service"
        Me.chkFilter_Service.Size = New System.Drawing.Size(62, 17)
        Me.chkFilter_Service.TabIndex = 1
        Me.chkFilter_Service.Text = "Service"
        Me.chkFilter_Service.UseVisualStyleBackColor = True
        '
        'chkFilter_Goods
        '
        Me.chkFilter_Goods.AutoSize = True
        Me.chkFilter_Goods.Checked = True
        Me.chkFilter_Goods.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkFilter_Goods.Location = New System.Drawing.Point(9, 19)
        Me.chkFilter_Goods.Name = "chkFilter_Goods"
        Me.chkFilter_Goods.Size = New System.Drawing.Size(57, 17)
        Me.chkFilter_Goods.TabIndex = 0
        Me.chkFilter_Goods.Text = "Goods"
        Me.chkFilter_Goods.UseVisualStyleBackColor = True
        '
        'TxtMP_QtyInquiry
        '
        Me.TxtMP_QtyInquiry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyInquiry.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyInquiry.Location = New System.Drawing.Point(91, 449)
        Me.TxtMP_QtyInquiry.Name = "TxtMP_QtyInquiry"
        Me.TxtMP_QtyInquiry.ReadOnly = True
        Me.TxtMP_QtyInquiry.Size = New System.Drawing.Size(139, 20)
        Me.TxtMP_QtyInquiry.TabIndex = 0
        Me.TxtMP_QtyInquiry.TabStop = False
        Me.TxtMP_QtyInquiry.Text = "0"
        Me.TxtMP_QtyInquiry.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtMP_QtyInquiry.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(0, 452)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(67, 13)
        Me.Label12.TabIndex = 55
        Me.Label12.Text = "Sales Inquiry"
        Me.Label12.Visible = False
        '
        'TxtFilter_PrimaryCustomer
        '
        Me.TxtFilter_PrimaryCustomer.Location = New System.Drawing.Point(75, 139)
        Me.TxtFilter_PrimaryCustomer.Name = "TxtFilter_PrimaryCustomer"
        Me.TxtFilter_PrimaryCustomer.Size = New System.Drawing.Size(22, 20)
        Me.TxtFilter_PrimaryCustomer.TabIndex = 88
        Me.TxtFilter_PrimaryCustomer.TabStop = False
        Me.TxtFilter_PrimaryCustomer.Text = "-1"
        Me.TxtFilter_PrimaryCustomer.Visible = False
        '
        'TxtFilter_PrimaryProductGroup
        '
        Me.TxtFilter_PrimaryProductGroup.Location = New System.Drawing.Point(75, 214)
        Me.TxtFilter_PrimaryProductGroup.Name = "TxtFilter_PrimaryProductGroup"
        Me.TxtFilter_PrimaryProductGroup.Size = New System.Drawing.Size(22, 20)
        Me.TxtFilter_PrimaryProductGroup.TabIndex = 6
        Me.TxtFilter_PrimaryProductGroup.TabStop = False
        Me.TxtFilter_PrimaryProductGroup.Text = "-1"
        Me.TxtFilter_PrimaryProductGroup.Visible = False
        '
        'TxtFilter_PrimaryPreferedSupplier
        '
        Me.TxtFilter_PrimaryPreferedSupplier.Location = New System.Drawing.Point(75, 139)
        Me.TxtFilter_PrimaryPreferedSupplier.Name = "TxtFilter_PrimaryPreferedSupplier"
        Me.TxtFilter_PrimaryPreferedSupplier.Size = New System.Drawing.Size(22, 20)
        Me.TxtFilter_PrimaryPreferedSupplier.TabIndex = 74
        Me.TxtFilter_PrimaryPreferedSupplier.TabStop = False
        Me.TxtFilter_PrimaryPreferedSupplier.Text = "-1"
        Me.TxtFilter_PrimaryPreferedSupplier.Visible = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(12, 418)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(103, 20)
        Me.TxtReadOnly.TabIndex = 18
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'llbMP_Filter
        '
        Me.llbMP_Filter.AutoSize = True
        Me.llbMP_Filter.Location = New System.Drawing.Point(8, 398)
        Me.llbMP_Filter.Name = "llbMP_Filter"
        Me.llbMP_Filter.Size = New System.Drawing.Size(29, 13)
        Me.llbMP_Filter.TabIndex = 22
        Me.llbMP_Filter.TabStop = True
        Me.llbMP_Filter.Text = "Filter"
        Me.llbMP_Filter.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtFilter_ProductGroupCode
        '
        Me.TxtFilter_ProductGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_ProductGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_ProductGroupCode.Location = New System.Drawing.Point(102, 213)
        Me.TxtFilter_ProductGroupCode.Name = "TxtFilter_ProductGroupCode"
        Me.TxtFilter_ProductGroupCode.Size = New System.Drawing.Size(69, 20)
        Me.TxtFilter_ProductGroupCode.TabIndex = 11
        '
        'LblFilter_PoductGroupDesc
        '
        Me.LblFilter_PoductGroupDesc.AutoSize = True
        Me.LblFilter_PoductGroupDesc.Location = New System.Drawing.Point(7, 215)
        Me.LblFilter_PoductGroupDesc.Name = "LblFilter_PoductGroupDesc"
        Me.LblFilter_PoductGroupDesc.Size = New System.Drawing.Size(76, 13)
        Me.LblFilter_PoductGroupDesc.TabIndex = 53
        Me.LblFilter_PoductGroupDesc.Text = "Product Group"
        '
        'TxtFilter_Desc
        '
        Me.TxtFilter_Desc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_Desc.Location = New System.Drawing.Point(102, 114)
        Me.TxtFilter_Desc.Name = "TxtFilter_Desc"
        Me.TxtFilter_Desc.Size = New System.Drawing.Size(115, 20)
        Me.TxtFilter_Desc.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 117)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = "Description"
        '
        'TxtFilter_CodeNo
        '
        Me.TxtFilter_CodeNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_CodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_CodeNo.Location = New System.Drawing.Point(102, 89)
        Me.TxtFilter_CodeNo.Name = "TxtFilter_CodeNo"
        Me.TxtFilter_CodeNo.Size = New System.Drawing.Size(69, 20)
        Me.TxtFilter_CodeNo.TabIndex = 5
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(7, 92)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(49, 13)
        Me.Label8.TabIndex = 49
        Me.Label8.Text = "Code No"
        '
        'LblFilter_UserCreateSO
        '
        Me.LblFilter_UserCreateSO.AutoSize = True
        Me.LblFilter_UserCreateSO.Location = New System.Drawing.Point(7, 265)
        Me.LblFilter_UserCreateSO.Name = "LblFilter_UserCreateSO"
        Me.LblFilter_UserCreateSO.Size = New System.Drawing.Size(59, 13)
        Me.LblFilter_UserCreateSO.TabIndex = 10
        Me.LblFilter_UserCreateSO.Text = "Created By"
        '
        'TxtFilter_CustCode
        '
        Me.TxtFilter_CustCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_CustCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_CustCode.Location = New System.Drawing.Point(102, 137)
        Me.TxtFilter_CustCode.Name = "TxtFilter_CustCode"
        Me.TxtFilter_CustCode.Size = New System.Drawing.Size(69, 20)
        Me.TxtFilter_CustCode.TabIndex = 7
        '
        'LblFilter_Cust
        '
        Me.LblFilter_Cust.AutoSize = True
        Me.LblFilter_Cust.Location = New System.Drawing.Point(7, 141)
        Me.LblFilter_Cust.Name = "LblFilter_Cust"
        Me.LblFilter_Cust.Size = New System.Drawing.Size(51, 13)
        Me.LblFilter_Cust.TabIndex = 10
        Me.LblFilter_Cust.Text = "Customer"
        '
        'cboFilter_ProductGroupLevel
        '
        Me.cboFilter_ProductGroupLevel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboFilter_ProductGroupLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFilter_ProductGroupLevel.FormattingEnabled = True
        Me.cboFilter_ProductGroupLevel.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category"})
        Me.cboFilter_ProductGroupLevel.Location = New System.Drawing.Point(120, 187)
        Me.cboFilter_ProductGroupLevel.Name = "cboFilter_ProductGroupLevel"
        Me.cboFilter_ProductGroupLevel.Size = New System.Drawing.Size(97, 21)
        Me.cboFilter_ProductGroupLevel.TabIndex = 10
        '
        'LblFilter_ProductGroup
        '
        Me.LblFilter_ProductGroup.AutoSize = True
        Me.LblFilter_ProductGroup.Location = New System.Drawing.Point(7, 189)
        Me.LblFilter_ProductGroup.Name = "LblFilter_ProductGroup"
        Me.LblFilter_ProductGroup.Size = New System.Drawing.Size(105, 13)
        Me.LblFilter_ProductGroup.TabIndex = 3
        Me.LblFilter_ProductGroup.Text = "Product Group Level"
        '
        'btnFilter_FindPreferedSupplier
        '
        Me.btnFilter_FindPreferedSupplier.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindPreferedSupplier.Location = New System.Drawing.Point(177, 135)
        Me.btnFilter_FindPreferedSupplier.Name = "btnFilter_FindPreferedSupplier"
        Me.btnFilter_FindPreferedSupplier.Size = New System.Drawing.Size(39, 23)
        Me.btnFilter_FindPreferedSupplier.TabIndex = 8
        Me.btnFilter_FindPreferedSupplier.TabStop = False
        Me.btnFilter_FindPreferedSupplier.Text = "Find"
        Me.btnFilter_FindPreferedSupplier.UseVisualStyleBackColor = True
        '
        'TxtFilter_PreferedSupplier
        '
        Me.TxtFilter_PreferedSupplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_PreferedSupplier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_PreferedSupplier.Location = New System.Drawing.Point(102, 137)
        Me.TxtFilter_PreferedSupplier.Name = "TxtFilter_PreferedSupplier"
        Me.TxtFilter_PreferedSupplier.Size = New System.Drawing.Size(69, 20)
        Me.TxtFilter_PreferedSupplier.TabIndex = 7
        '
        'LblFilter_PreferedSupplier
        '
        Me.LblFilter_PreferedSupplier.AutoSize = True
        Me.LblFilter_PreferedSupplier.Location = New System.Drawing.Point(7, 141)
        Me.LblFilter_PreferedSupplier.Name = "LblFilter_PreferedSupplier"
        Me.LblFilter_PreferedSupplier.Size = New System.Drawing.Size(88, 13)
        Me.LblFilter_PreferedSupplier.TabIndex = 0
        Me.LblFilter_PreferedSupplier.Text = "Prefered Supplier"
        '
        'TxtFilter_PreferedSupplierDesc
        '
        Me.TxtFilter_PreferedSupplierDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_PreferedSupplierDesc.Location = New System.Drawing.Point(9, 162)
        Me.TxtFilter_PreferedSupplierDesc.Name = "TxtFilter_PreferedSupplierDesc"
        Me.TxtFilter_PreferedSupplierDesc.ReadOnly = True
        Me.TxtFilter_PreferedSupplierDesc.Size = New System.Drawing.Size(207, 20)
        Me.TxtFilter_PreferedSupplierDesc.TabIndex = 17
        Me.TxtFilter_PreferedSupplierDesc.TabStop = False
        '
        'btnFilter_FindRefParty
        '
        Me.btnFilter_FindRefParty.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindRefParty.Location = New System.Drawing.Point(177, 385)
        Me.btnFilter_FindRefParty.Name = "btnFilter_FindRefParty"
        Me.btnFilter_FindRefParty.Size = New System.Drawing.Size(39, 23)
        Me.btnFilter_FindRefParty.TabIndex = 20
        Me.btnFilter_FindRefParty.TabStop = False
        Me.btnFilter_FindRefParty.Text = "Find"
        Me.btnFilter_FindRefParty.UseVisualStyleBackColor = True
        '
        'Txtfilter_RefPartyName
        '
        Me.Txtfilter_RefPartyName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txtfilter_RefPartyName.Location = New System.Drawing.Point(9, 412)
        Me.Txtfilter_RefPartyName.Name = "Txtfilter_RefPartyName"
        Me.Txtfilter_RefPartyName.ReadOnly = True
        Me.Txtfilter_RefPartyName.Size = New System.Drawing.Size(207, 20)
        Me.Txtfilter_RefPartyName.TabIndex = 104
        Me.Txtfilter_RefPartyName.TabStop = False
        '
        'btnFilter_FindProductGroupLevel
        '
        Me.btnFilter_FindProductGroupLevel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindProductGroupLevel.Location = New System.Drawing.Point(177, 210)
        Me.btnFilter_FindProductGroupLevel.Name = "btnFilter_FindProductGroupLevel"
        Me.btnFilter_FindProductGroupLevel.Size = New System.Drawing.Size(39, 23)
        Me.btnFilter_FindProductGroupLevel.TabIndex = 12
        Me.btnFilter_FindProductGroupLevel.TabStop = False
        Me.btnFilter_FindProductGroupLevel.Text = "Find"
        Me.btnFilter_FindProductGroupLevel.UseVisualStyleBackColor = True
        '
        'btnFilter_FindCust
        '
        Me.btnFilter_FindCust.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindCust.Location = New System.Drawing.Point(177, 135)
        Me.btnFilter_FindCust.Name = "btnFilter_FindCust"
        Me.btnFilter_FindCust.Size = New System.Drawing.Size(39, 23)
        Me.btnFilter_FindCust.TabIndex = 8
        Me.btnFilter_FindCust.TabStop = False
        Me.btnFilter_FindCust.Text = "Find"
        Me.btnFilter_FindCust.UseVisualStyleBackColor = True
        '
        'TxtFilter_ProductGroupLevelDesc
        '
        Me.TxtFilter_ProductGroupLevelDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_ProductGroupLevelDesc.Location = New System.Drawing.Point(9, 238)
        Me.TxtFilter_ProductGroupLevelDesc.Name = "TxtFilter_ProductGroupLevelDesc"
        Me.TxtFilter_ProductGroupLevelDesc.ReadOnly = True
        Me.TxtFilter_ProductGroupLevelDesc.Size = New System.Drawing.Size(207, 20)
        Me.TxtFilter_ProductGroupLevelDesc.TabIndex = 13
        Me.TxtFilter_ProductGroupLevelDesc.TabStop = False
        '
        'TxtFilter_CustName
        '
        Me.TxtFilter_CustName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_CustName.Location = New System.Drawing.Point(9, 162)
        Me.TxtFilter_CustName.Name = "TxtFilter_CustName"
        Me.TxtFilter_CustName.ReadOnly = True
        Me.TxtFilter_CustName.Size = New System.Drawing.Size(207, 20)
        Me.TxtFilter_CustName.TabIndex = 4
        Me.TxtFilter_CustName.TabStop = False
        '
        'TxtFilter_UserCreateSO
        '
        Me.TxtFilter_UserCreateSO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_UserCreateSO.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtFilter_UserCreateSO.Location = New System.Drawing.Point(103, 262)
        Me.TxtFilter_UserCreateSO.Name = "TxtFilter_UserCreateSO"
        Me.TxtFilter_UserCreateSO.Size = New System.Drawing.Size(69, 20)
        Me.TxtFilter_UserCreateSO.TabIndex = 15
        '
        'btnFilter_FindUser
        '
        Me.btnFilter_FindUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFilter_FindUser.Location = New System.Drawing.Point(177, 259)
        Me.btnFilter_FindUser.Name = "btnFilter_FindUser"
        Me.btnFilter_FindUser.Size = New System.Drawing.Size(39, 23)
        Me.btnFilter_FindUser.TabIndex = 16
        Me.btnFilter_FindUser.TabStop = False
        Me.btnFilter_FindUser.Text = "Find"
        Me.btnFilter_FindUser.UseVisualStyleBackColor = True
        '
        'LblFilter_ProcurementStrategy
        '
        Me.LblFilter_ProcurementStrategy.AutoSize = True
        Me.LblFilter_ProcurementStrategy.Location = New System.Drawing.Point(11, 264)
        Me.LblFilter_ProcurementStrategy.Name = "LblFilter_ProcurementStrategy"
        Me.LblFilter_ProcurementStrategy.Size = New System.Drawing.Size(109, 13)
        Me.LblFilter_ProcurementStrategy.TabIndex = 11
        Me.LblFilter_ProcurementStrategy.Text = "Procurement Strategy"
        '
        'GboFilter_MODate
        '
        Me.GboFilter_MODate.Controls.Add(Me.chkMODateFilter)
        Me.GboFilter_MODate.Controls.Add(Me.DtpFilter_MO_To)
        Me.GboFilter_MODate.Controls.Add(Me.Label37)
        Me.GboFilter_MODate.Controls.Add(Me.DtpFilter_MO_From)
        Me.GboFilter_MODate.Controls.Add(Me.Label38)
        Me.GboFilter_MODate.Location = New System.Drawing.Point(9, 316)
        Me.GboFilter_MODate.Name = "GboFilter_MODate"
        Me.GboFilter_MODate.Size = New System.Drawing.Size(219, 65)
        Me.GboFilter_MODate.TabIndex = 112
        Me.GboFilter_MODate.TabStop = False
        '
        'chkMODateFilter
        '
        Me.chkMODateFilter.AutoSize = True
        Me.chkMODateFilter.Location = New System.Drawing.Point(10, -1)
        Me.chkMODateFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.chkMODateFilter.Name = "chkMODateFilter"
        Me.chkMODateFilter.Size = New System.Drawing.Size(69, 17)
        Me.chkMODateFilter.TabIndex = 0
        Me.chkMODateFilter.Text = "MO Date"
        Me.chkMODateFilter.UseVisualStyleBackColor = True
        '
        'DtpFilter_MO_To
        '
        Me.DtpFilter_MO_To.CustomFormat = "dd MMM yyyy"
        Me.DtpFilter_MO_To.Enabled = False
        Me.DtpFilter_MO_To.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFilter_MO_To.Location = New System.Drawing.Point(82, 38)
        Me.DtpFilter_MO_To.Name = "DtpFilter_MO_To"
        Me.DtpFilter_MO_To.Size = New System.Drawing.Size(130, 20)
        Me.DtpFilter_MO_To.TabIndex = 2
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Enabled = False
        Me.Label37.Location = New System.Drawing.Point(8, 42)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(20, 13)
        Me.Label37.TabIndex = 2
        Me.Label37.Text = "To"
        '
        'DtpFilter_MO_From
        '
        Me.DtpFilter_MO_From.CustomFormat = "dd MMM yyyy"
        Me.DtpFilter_MO_From.Enabled = False
        Me.DtpFilter_MO_From.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFilter_MO_From.Location = New System.Drawing.Point(82, 15)
        Me.DtpFilter_MO_From.Name = "DtpFilter_MO_From"
        Me.DtpFilter_MO_From.Size = New System.Drawing.Size(130, 20)
        Me.DtpFilter_MO_From.TabIndex = 1
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Enabled = False
        Me.Label38.Location = New System.Drawing.Point(8, 19)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(30, 13)
        Me.Label38.TabIndex = 0
        Me.Label38.Text = "From"
        '
        'GboFilter_SODate
        '
        Me.GboFilter_SODate.Controls.Add(Me.chkSODateFilter)
        Me.GboFilter_SODate.Controls.Add(Me.DtpFilter_To)
        Me.GboFilter_SODate.Controls.Add(Me.Label21)
        Me.GboFilter_SODate.Controls.Add(Me.DtpFilter_From)
        Me.GboFilter_SODate.Controls.Add(Me.Label7)
        Me.GboFilter_SODate.Location = New System.Drawing.Point(9, 316)
        Me.GboFilter_SODate.Name = "GboFilter_SODate"
        Me.GboFilter_SODate.Size = New System.Drawing.Size(219, 65)
        Me.GboFilter_SODate.TabIndex = 18
        Me.GboFilter_SODate.TabStop = False
        '
        'chkSODateFilter
        '
        Me.chkSODateFilter.AutoSize = True
        Me.chkSODateFilter.Location = New System.Drawing.Point(10, -1)
        Me.chkSODateFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSODateFilter.Name = "chkSODateFilter"
        Me.chkSODateFilter.Size = New System.Drawing.Size(67, 17)
        Me.chkSODateFilter.TabIndex = 0
        Me.chkSODateFilter.Text = "SO Date"
        Me.chkSODateFilter.UseVisualStyleBackColor = True
        '
        'DtpFilter_To
        '
        Me.DtpFilter_To.CustomFormat = "dd MMM yyyy"
        Me.DtpFilter_To.Enabled = False
        Me.DtpFilter_To.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFilter_To.Location = New System.Drawing.Point(82, 38)
        Me.DtpFilter_To.Name = "DtpFilter_To"
        Me.DtpFilter_To.Size = New System.Drawing.Size(130, 20)
        Me.DtpFilter_To.TabIndex = 2
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Enabled = False
        Me.Label21.Location = New System.Drawing.Point(8, 42)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(20, 13)
        Me.Label21.TabIndex = 2
        Me.Label21.Text = "To"
        '
        'DtpFilter_From
        '
        Me.DtpFilter_From.CustomFormat = "dd MMM yyyy"
        Me.DtpFilter_From.Enabled = False
        Me.DtpFilter_From.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFilter_From.Location = New System.Drawing.Point(82, 15)
        Me.DtpFilter_From.Name = "DtpFilter_From"
        Me.DtpFilter_From.Size = New System.Drawing.Size(130, 20)
        Me.DtpFilter_From.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Enabled = False
        Me.Label7.Location = New System.Drawing.Point(8, 19)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(30, 13)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "From"
        '
        'CboFilter_ProcurementStrategy
        '
        Me.CboFilter_ProcurementStrategy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboFilter_ProcurementStrategy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFilter_ProcurementStrategy.FormattingEnabled = True
        Me.CboFilter_ProcurementStrategy.Items.AddRange(New Object() {"Both", "Build to stock", "Build to order"})
        Me.CboFilter_ProcurementStrategy.Location = New System.Drawing.Point(127, 261)
        Me.CboFilter_ProcurementStrategy.Name = "CboFilter_ProcurementStrategy"
        Me.CboFilter_ProcurementStrategy.Size = New System.Drawing.Size(91, 21)
        Me.CboFilter_ProcurementStrategy.TabIndex = 14
        '
        'TxtFilter_NamaUserCreateSO
        '
        Me.TxtFilter_NamaUserCreateSO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtFilter_NamaUserCreateSO.Location = New System.Drawing.Point(9, 290)
        Me.TxtFilter_NamaUserCreateSO.Name = "TxtFilter_NamaUserCreateSO"
        Me.TxtFilter_NamaUserCreateSO.ReadOnly = True
        Me.TxtFilter_NamaUserCreateSO.Size = New System.Drawing.Size(207, 20)
        Me.TxtFilter_NamaUserCreateSO.TabIndex = 17
        Me.TxtFilter_NamaUserCreateSO.TabStop = False
        '
        'chkFilter_PrOBLargerThanPOB
        '
        Me.chkFilter_PrOBLargerThanPOB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkFilter_PrOBLargerThanPOB.AutoSize = True
        Me.chkFilter_PrOBLargerThanPOB.Location = New System.Drawing.Point(5, 406)
        Me.chkFilter_PrOBLargerThanPOB.Margin = New System.Windows.Forms.Padding(2)
        Me.chkFilter_PrOBLargerThanPOB.Name = "chkFilter_PrOBLargerThanPOB"
        Me.chkFilter_PrOBLargerThanPOB.Size = New System.Drawing.Size(127, 17)
        Me.chkFilter_PrOBLargerThanPOB.TabIndex = 3
        Me.chkFilter_PrOBLargerThanPOB.Text = "Item w/ Pr.OB > POB"
        Me.chkFilter_PrOBLargerThanPOB.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(181, 469)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(0, 85)
        Me.btnClose.TabIndex = 22
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'chkMP_RequiredOnly
        '
        Me.chkMP_RequiredOnly.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkMP_RequiredOnly.AutoSize = True
        Me.chkMP_RequiredOnly.Checked = True
        Me.chkMP_RequiredOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMP_RequiredOnly.Location = New System.Drawing.Point(5, 427)
        Me.chkMP_RequiredOnly.Name = "chkMP_RequiredOnly"
        Me.chkMP_RequiredOnly.Size = New System.Drawing.Size(79, 17)
        Me.chkMP_RequiredOnly.TabIndex = 4
        Me.chkMP_RequiredOnly.Text = "Order Point"
        Me.chkMP_RequiredOnly.UseVisualStyleBackColor = True
        '
        'btnMP_Print
        '
        Me.btnMP_Print.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnMP_Print.Location = New System.Drawing.Point(400, 456)
        Me.btnMP_Print.Name = "btnMP_Print"
        Me.btnMP_Print.Size = New System.Drawing.Size(69, 23)
        Me.btnMP_Print.TabIndex = 23
        Me.btnMP_Print.Text = "Print"
        Me.btnMP_Print.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(11, 3)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer1.Size = New System.Drawing.Size(821, 531)
        Me.SplitContainer1.SplitterDistance = 233
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 31
        Me.SplitContainer1.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(585, 531)
        Me.TabControl1.TabIndex = 34
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnMP_Edit)
        Me.TabPage1.Controls.Add(Me.btnMP_New)
        Me.TabPage1.Controls.Add(Me.StatusStrip4)
        Me.TabPage1.Controls.Add(Me.gboMP)
        Me.TabPage1.Controls.Add(Me.btnMP_Find)
        Me.TabPage1.Controls.Add(Me.btnMP_Print)
        Me.TabPage1.Controls.Add(Me.btnMP_Delete)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(577, 505)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Material Planning"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnMP_Edit
        '
        Me.btnMP_Edit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnMP_Edit.Location = New System.Drawing.Point(250, 456)
        Me.btnMP_Edit.Name = "btnMP_Edit"
        Me.btnMP_Edit.Size = New System.Drawing.Size(69, 23)
        Me.btnMP_Edit.TabIndex = 46
        Me.btnMP_Edit.Text = "Edit"
        Me.btnMP_Edit.UseVisualStyleBackColor = True
        '
        'btnMP_New
        '
        Me.btnMP_New.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnMP_New.Location = New System.Drawing.Point(176, 456)
        Me.btnMP_New.Name = "btnMP_New"
        Me.btnMP_New.Size = New System.Drawing.Size(69, 23)
        Me.btnMP_New.TabIndex = 45
        Me.btnMP_New.Text = "New"
        Me.btnMP_New.UseVisualStyleBackColor = True
        '
        'StatusStrip4
        '
        Me.StatusStrip4.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslMP_CreatedBy, Me.tslMP_CreatedDate, Me.tslMP_ModifiedBy, Me.tslMP_ModifiedDate})
        Me.StatusStrip4.Location = New System.Drawing.Point(2, 481)
        Me.StatusStrip4.Name = "StatusStrip4"
        Me.StatusStrip4.Size = New System.Drawing.Size(573, 22)
        Me.StatusStrip4.TabIndex = 40
        Me.StatusStrip4.Text = "StatusStrip4"
        '
        'tslMP_CreatedBy
        '
        Me.tslMP_CreatedBy.Name = "tslMP_CreatedBy"
        Me.tslMP_CreatedBy.Size = New System.Drawing.Size(139, 17)
        Me.tslMP_CreatedBy.Spring = True
        Me.tslMP_CreatedBy.Text = "Created By:"
        '
        'tslMP_CreatedDate
        '
        Me.tslMP_CreatedDate.Name = "tslMP_CreatedDate"
        Me.tslMP_CreatedDate.Size = New System.Drawing.Size(139, 17)
        Me.tslMP_CreatedDate.Spring = True
        Me.tslMP_CreatedDate.Text = "Created Date:"
        '
        'tslMP_ModifiedBy
        '
        Me.tslMP_ModifiedBy.Name = "tslMP_ModifiedBy"
        Me.tslMP_ModifiedBy.Size = New System.Drawing.Size(139, 17)
        Me.tslMP_ModifiedBy.Spring = True
        Me.tslMP_ModifiedBy.Text = "Modified By:"
        '
        'tslMP_ModifiedDate
        '
        Me.tslMP_ModifiedDate.Name = "tslMP_ModifiedDate"
        Me.tslMP_ModifiedDate.Size = New System.Drawing.Size(139, 17)
        Me.tslMP_ModifiedDate.Spring = True
        Me.tslMP_ModifiedDate.Text = "Modified Date:"
        '
        'gboMP
        '
        Me.gboMP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboMP.Controls.Add(Me.llmbMP_SetQtyOrderEqualQtyPrOB)
        Me.gboMP.Controls.Add(Me.chkFilter_PrOBLargerThanPOB)
        Me.gboMP.Controls.Add(Me.llbMP_CopyMP)
        Me.gboMP.Controls.Add(Me.llbMP_CekAll)
        Me.gboMP.Controls.Add(Me.llbMP_AddToPO)
        Me.gboMP.Controls.Add(Me.dtpMP_Filter)
        Me.gboMP.Controls.Add(Me.llbMP_CreatePO)
        Me.gboMP.Controls.Add(Me.btnMP_Cancel)
        Me.gboMP.Controls.Add(Me.SplitContainer5)
        Me.gboMP.Controls.Add(Me.chkMP_RequiredOnly)
        Me.gboMP.Controls.Add(Me.btnMP_OK)
        Me.gboMP.Controls.Add(Me.TabControl2)
        Me.gboMP.Controls.Add(Me.dgvMP)
        Me.gboMP.Location = New System.Drawing.Point(3, 1)
        Me.gboMP.Margin = New System.Windows.Forms.Padding(2)
        Me.gboMP.Name = "gboMP"
        Me.gboMP.Padding = New System.Windows.Forms.Padding(2)
        Me.gboMP.Size = New System.Drawing.Size(572, 450)
        Me.gboMP.TabIndex = 44
        Me.gboMP.TabStop = False
        '
        'llmbMP_SetQtyOrderEqualQtyPrOB
        '
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.AutoSize = True
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.Location = New System.Drawing.Point(441, 404)
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.Name = "llmbMP_SetQtyOrderEqualQtyPrOB"
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.Size = New System.Drawing.Size(127, 13)
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.TabIndex = 89
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.TabStop = True
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.Text = "Set Qty Order = Qty PrOB"
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.Visible = False
        Me.llmbMP_SetQtyOrderEqualQtyPrOB.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbMP_CopyMP
        '
        Me.llbMP_CopyMP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbMP_CopyMP.AutoSize = True
        Me.llbMP_CopyMP.Location = New System.Drawing.Point(515, 427)
        Me.llbMP_CopyMP.Name = "llbMP_CopyMP"
        Me.llbMP_CopyMP.Size = New System.Drawing.Size(57, 13)
        Me.llbMP_CopyMP.TabIndex = 8
        Me.llbMP_CopyMP.TabStop = True
        Me.llbMP_CopyMP.Text = "Copy Data"
        Me.llbMP_CopyMP.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbMP_CekAll
        '
        Me.llbMP_CekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbMP_CekAll.AutoSize = True
        Me.llbMP_CekAll.Location = New System.Drawing.Point(374, 230)
        Me.llbMP_CekAll.Name = "llbMP_CekAll"
        Me.llbMP_CekAll.Size = New System.Drawing.Size(52, 13)
        Me.llbMP_CekAll.TabIndex = 88
        Me.llbMP_CekAll.TabStop = True
        Me.llbMP_CekAll.Text = "Check All"
        Me.llbMP_CekAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbMP_AddToPO
        '
        Me.llbMP_AddToPO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbMP_AddToPO.AutoSize = True
        Me.llbMP_AddToPO.Location = New System.Drawing.Point(443, 230)
        Me.llbMP_AddToPO.Name = "llbMP_AddToPO"
        Me.llbMP_AddToPO.Size = New System.Drawing.Size(60, 13)
        Me.llbMP_AddToPO.TabIndex = 87
        Me.llbMP_AddToPO.TabStop = True
        Me.llbMP_AddToPO.Text = "Add To PO"
        Me.llbMP_AddToPO.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'dtpMP_Filter
        '
        Me.dtpMP_Filter.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dtpMP_Filter.CustomFormat = "dd MMM yyyy"
        Me.dtpMP_Filter.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpMP_Filter.Location = New System.Drawing.Point(88, 425)
        Me.dtpMP_Filter.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpMP_Filter.Name = "dtpMP_Filter"
        Me.dtpMP_Filter.Size = New System.Drawing.Size(100, 20)
        Me.dtpMP_Filter.TabIndex = 5
        '
        'llbMP_CreatePO
        '
        Me.llbMP_CreatePO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbMP_CreatePO.AutoSize = True
        Me.llbMP_CreatePO.Location = New System.Drawing.Point(509, 230)
        Me.llbMP_CreatePO.Name = "llbMP_CreatePO"
        Me.llbMP_CreatePO.Size = New System.Drawing.Size(56, 13)
        Me.llbMP_CreatePO.TabIndex = 85
        Me.llbMP_CreatePO.TabStop = True
        Me.llbMP_CreatePO.Text = "Create PO"
        Me.llbMP_CreatePO.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'btnMP_Cancel
        '
        Me.btnMP_Cancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnMP_Cancel.Location = New System.Drawing.Point(287, 422)
        Me.btnMP_Cancel.Name = "btnMP_Cancel"
        Me.btnMP_Cancel.Size = New System.Drawing.Size(69, 23)
        Me.btnMP_Cancel.TabIndex = 7
        Me.btnMP_Cancel.Text = "Cancel"
        Me.btnMP_Cancel.UseVisualStyleBackColor = True
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer5.Location = New System.Drawing.Point(4, 11)
        Me.SplitContainer5.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer5.Name = "SplitContainer5"
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.TxtMP_DaysPerMonth)
        Me.SplitContainer5.Panel1.Controls.Add(Me.Label29)
        Me.SplitContainer5.Panel1.Controls.Add(Me.Label27)
        Me.SplitContainer5.Panel1.Controls.Add(Me.dtpMP_RefDate)
        Me.SplitContainer5.Panel1.Controls.Add(Me.dtpMP)
        Me.SplitContainer5.Panel1.Controls.Add(Me.llbMP_Calculate)
        Me.SplitContainer5.Panel1.Controls.Add(Me.TxtMP_Primary)
        Me.SplitContainer5.Panel1.Controls.Add(Me.TxtMP_DocNo)
        Me.SplitContainer5.Panel1.Controls.Add(Me.Label26)
        '
        'SplitContainer5.Panel2
        '
        Me.SplitContainer5.Panel2.Controls.Add(Me.llbAuthDetail)
        Me.SplitContainer5.Panel2.Controls.Add(Me.llbMP_LocationSetting)
        Me.SplitContainer5.Panel2.Controls.Add(Me.TxtMP_Remarks)
        Me.SplitContainer5.Panel2.Controls.Add(Me.Label28)
        Me.SplitContainer5.Panel2.Controls.Add(Me.chkClose)
        Me.SplitContainer5.Size = New System.Drawing.Size(572, 67)
        Me.SplitContainer5.SplitterDistance = 269
        Me.SplitContainer5.SplitterWidth = 3
        Me.SplitContainer5.TabIndex = 0
        Me.SplitContainer5.TabStop = False
        '
        'TxtMP_DaysPerMonth
        '
        Me.TxtMP_DaysPerMonth.Location = New System.Drawing.Point(70, 23)
        Me.TxtMP_DaysPerMonth.Name = "TxtMP_DaysPerMonth"
        Me.TxtMP_DaysPerMonth.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_DaysPerMonth.TabIndex = 73
        Me.TxtMP_DaysPerMonth.TabStop = False
        Me.TxtMP_DaysPerMonth.Text = "0"
        Me.TxtMP_DaysPerMonth.Visible = False
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(1, 47)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(53, 13)
        Me.Label29.TabIndex = 76
        Me.Label29.Text = "Ref. Date"
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(1, 27)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(30, 13)
        Me.Label27.TabIndex = 72
        Me.Label27.Text = "Date"
        '
        'dtpMP_RefDate
        '
        Me.dtpMP_RefDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpMP_RefDate.CustomFormat = "dd MMM yyyy"
        Me.dtpMP_RefDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpMP_RefDate.Location = New System.Drawing.Point(96, 46)
        Me.dtpMP_RefDate.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpMP_RefDate.Name = "dtpMP_RefDate"
        Me.dtpMP_RefDate.Size = New System.Drawing.Size(171, 20)
        Me.dtpMP_RefDate.TabIndex = 75
        '
        'dtpMP
        '
        Me.dtpMP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpMP.CustomFormat = "dd MMM yyyy"
        Me.dtpMP.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpMP.Location = New System.Drawing.Point(96, 23)
        Me.dtpMP.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpMP.Name = "dtpMP"
        Me.dtpMP.Size = New System.Drawing.Size(171, 20)
        Me.dtpMP.TabIndex = 71
        '
        'llbMP_Calculate
        '
        Me.llbMP_Calculate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbMP_Calculate.AutoSize = True
        Me.llbMP_Calculate.Location = New System.Drawing.Point(217, 2)
        Me.llbMP_Calculate.Name = "llbMP_Calculate"
        Me.llbMP_Calculate.Size = New System.Drawing.Size(51, 13)
        Me.llbMP_Calculate.TabIndex = 86
        Me.llbMP_Calculate.TabStop = True
        Me.llbMP_Calculate.Text = "Calculate"
        Me.llbMP_Calculate.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtMP_Primary
        '
        Me.TxtMP_Primary.Location = New System.Drawing.Point(69, 0)
        Me.TxtMP_Primary.Name = "TxtMP_Primary"
        Me.TxtMP_Primary.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_Primary.TabIndex = 70
        Me.TxtMP_Primary.TabStop = False
        Me.TxtMP_Primary.Text = "-1"
        Me.TxtMP_Primary.Visible = False
        '
        'TxtMP_DocNo
        '
        Me.TxtMP_DocNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_DocNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMP_DocNo.Location = New System.Drawing.Point(96, 0)
        Me.TxtMP_DocNo.Name = "TxtMP_DocNo"
        Me.TxtMP_DocNo.Size = New System.Drawing.Size(116, 20)
        Me.TxtMP_DocNo.TabIndex = 52
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(1, 2)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(73, 13)
        Me.Label26.TabIndex = 51
        Me.Label26.Text = "Document No"
        '
        'llbAuthDetail
        '
        Me.llbAuthDetail.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbAuthDetail.AutoSize = True
        Me.llbAuthDetail.Location = New System.Drawing.Point(106, 47)
        Me.llbAuthDetail.Name = "llbAuthDetail"
        Me.llbAuthDetail.Size = New System.Drawing.Size(98, 13)
        Me.llbAuthDetail.TabIndex = 89
        Me.llbAuthDetail.TabStop = True
        Me.llbAuthDetail.Text = "Authorization Detail"
        Me.llbAuthDetail.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'llbMP_LocationSetting
        '
        Me.llbMP_LocationSetting.AutoSize = True
        Me.llbMP_LocationSetting.Location = New System.Drawing.Point(7, 47)
        Me.llbMP_LocationSetting.Name = "llbMP_LocationSetting"
        Me.llbMP_LocationSetting.Size = New System.Drawing.Size(84, 13)
        Me.llbMP_LocationSetting.TabIndex = 87
        Me.llbMP_LocationSetting.TabStop = True
        Me.llbMP_LocationSetting.Text = "Location Setting"
        Me.llbMP_LocationSetting.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtMP_Remarks
        '
        Me.TxtMP_Remarks.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_Remarks.Location = New System.Drawing.Point(55, 0)
        Me.TxtMP_Remarks.Multiline = True
        Me.TxtMP_Remarks.Name = "TxtMP_Remarks"
        Me.TxtMP_Remarks.Size = New System.Drawing.Size(240, 42)
        Me.TxtMP_Remarks.TabIndex = 74
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(1, 2)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(49, 13)
        Me.Label28.TabIndex = 73
        Me.Label28.Text = "Remarks"
        '
        'chkClose
        '
        Me.chkClose.AutoSize = True
        Me.chkClose.Location = New System.Drawing.Point(220, 46)
        Me.chkClose.Margin = New System.Windows.Forms.Padding(2)
        Me.chkClose.Name = "chkClose"
        Me.chkClose.Size = New System.Drawing.Size(69, 17)
        Me.chkClose.TabIndex = 88
        Me.chkClose.Text = "Is Closed"
        Me.chkClose.UseVisualStyleBackColor = True
        '
        'btnMP_OK
        '
        Me.btnMP_OK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnMP_OK.Location = New System.Drawing.Point(213, 422)
        Me.btnMP_OK.Name = "btnMP_OK"
        Me.btnMP_OK.Size = New System.Drawing.Size(69, 23)
        Me.btnMP_OK.TabIndex = 6
        Me.btnMP_OK.Text = "OK"
        Me.btnMP_OK.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Location = New System.Drawing.Point(2, 228)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(569, 175)
        Me.TabControl2.TabIndex = 2
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.SplitContainer2)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(561, 149)
        Me.TabPage4.TabIndex = 0
        Me.TabPage4.Text = "Detail Qty"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.Location = New System.Drawing.Point(2, 2)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label42)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyRMReqSpare)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label41)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtySOBSpare)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label40)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyPrOBSpare)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label33)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyRMReq)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label31)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyPrOB)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtySafetyStock)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label16)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtySOB)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyOptimumStock)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label14)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TxtMP_QtyPendingDO)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label17)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyFree)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label39)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyMOB)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label36)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label30)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyRequired)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyOnHand)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label18)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyPOB)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_QtyGIT)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer2.Panel2.Controls.Add(Me.TxtMP_AvgPerMonth)
        Me.SplitContainer2.Panel2.Controls.Add(Me.cboMPOrderStrategy)
        Me.SplitContainer2.Size = New System.Drawing.Size(563, 158)
        Me.SplitContainer2.SplitterDistance = 259
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 71
        Me.SplitContainer2.TabStop = False
        '
        'Label42
        '
        Me.Label42.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(177, 78)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(13, 13)
        Me.Label42.TabIndex = 77
        Me.Label42.Text = "+"
        '
        'TxtMP_QtyRMReqSpare
        '
        Me.TxtMP_QtyRMReqSpare.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRMReqSpare.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyRMReqSpare.Location = New System.Drawing.Point(195, 76)
        Me.TxtMP_QtyRMReqSpare.Name = "TxtMP_QtyRMReqSpare"
        Me.TxtMP_QtyRMReqSpare.ReadOnly = True
        Me.TxtMP_QtyRMReqSpare.Size = New System.Drawing.Size(61, 20)
        Me.TxtMP_QtyRMReqSpare.TabIndex = 76
        Me.TxtMP_QtyRMReqSpare.Text = "0"
        Me.TxtMP_QtyRMReqSpare.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label41
        '
        Me.Label41.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(177, 30)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(13, 13)
        Me.Label41.TabIndex = 75
        Me.Label41.Text = "+"
        '
        'TxtMP_QtySOBSpare
        '
        Me.TxtMP_QtySOBSpare.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtySOBSpare.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtySOBSpare.Location = New System.Drawing.Point(195, 28)
        Me.TxtMP_QtySOBSpare.Name = "TxtMP_QtySOBSpare"
        Me.TxtMP_QtySOBSpare.ReadOnly = True
        Me.TxtMP_QtySOBSpare.Size = New System.Drawing.Size(61, 20)
        Me.TxtMP_QtySOBSpare.TabIndex = 74
        Me.TxtMP_QtySOBSpare.Text = "0"
        Me.TxtMP_QtySOBSpare.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label40
        '
        Me.Label40.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(177, 6)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(13, 13)
        Me.Label40.TabIndex = 73
        Me.Label40.Text = "+"
        '
        'TxtMP_QtyPrOBSpare
        '
        Me.TxtMP_QtyPrOBSpare.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPrOBSpare.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPrOBSpare.Location = New System.Drawing.Point(195, 4)
        Me.TxtMP_QtyPrOBSpare.Name = "TxtMP_QtyPrOBSpare"
        Me.TxtMP_QtyPrOBSpare.ReadOnly = True
        Me.TxtMP_QtyPrOBSpare.Size = New System.Drawing.Size(61, 20)
        Me.TxtMP_QtyPrOBSpare.TabIndex = 72
        Me.TxtMP_QtyPrOBSpare.Text = "0"
        Me.TxtMP_QtyPrOBSpare.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(3, 79)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(95, 13)
        Me.Label33.TabIndex = 71
        Me.Label33.Text = "Raw Material Req."
        '
        'TxtMP_QtyRMReq
        '
        Me.TxtMP_QtyRMReq.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRMReq.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyRMReq.Location = New System.Drawing.Point(101, 76)
        Me.TxtMP_QtyRMReq.Name = "TxtMP_QtyRMReq"
        Me.TxtMP_QtyRMReq.ReadOnly = True
        Me.TxtMP_QtyRMReq.Size = New System.Drawing.Size(71, 20)
        Me.TxtMP_QtyRMReq.TabIndex = 70
        Me.TxtMP_QtyRMReq.Text = "0"
        Me.TxtMP_QtyRMReq.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(3, 5)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(81, 13)
        Me.Label31.TabIndex = 69
        Me.Label31.Text = "Purch. Request"
        '
        'TxtMP_QtyPrOB
        '
        Me.TxtMP_QtyPrOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPrOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPrOB.Location = New System.Drawing.Point(101, 4)
        Me.TxtMP_QtyPrOB.Name = "TxtMP_QtyPrOB"
        Me.TxtMP_QtyPrOB.ReadOnly = True
        Me.TxtMP_QtyPrOB.Size = New System.Drawing.Size(71, 20)
        Me.TxtMP_QtyPrOB.TabIndex = 0
        Me.TxtMP_QtyPrOB.Text = "0"
        Me.TxtMP_QtyPrOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtySafetyStock
        '
        Me.TxtMP_QtySafetyStock.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtySafetyStock.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtySafetyStock.Location = New System.Drawing.Point(101, 126)
        Me.TxtMP_QtySafetyStock.Name = "TxtMP_QtySafetyStock"
        Me.TxtMP_QtySafetyStock.ReadOnly = True
        Me.TxtMP_QtySafetyStock.Size = New System.Drawing.Size(155, 20)
        Me.TxtMP_QtySafetyStock.TabIndex = 4
        Me.TxtMP_QtySafetyStock.Text = "0"
        Me.TxtMP_QtySafetyStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 30)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(75, 13)
        Me.Label11.TabIndex = 57
        Me.Label11.Text = "Sales Balance"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(3, 128)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(68, 13)
        Me.Label16.TabIndex = 67
        Me.Label16.Text = "Safety Stock"
        '
        'TxtMP_QtySOB
        '
        Me.TxtMP_QtySOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtySOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtySOB.Location = New System.Drawing.Point(101, 28)
        Me.TxtMP_QtySOB.Name = "TxtMP_QtySOB"
        Me.TxtMP_QtySOB.ReadOnly = True
        Me.TxtMP_QtySOB.Size = New System.Drawing.Size(71, 20)
        Me.TxtMP_QtySOB.TabIndex = 1
        Me.TxtMP_QtySOB.Text = "0"
        Me.TxtMP_QtySOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtyOptimumStock
        '
        Me.TxtMP_QtyOptimumStock.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyOptimumStock.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyOptimumStock.Location = New System.Drawing.Point(101, 101)
        Me.TxtMP_QtyOptimumStock.Name = "TxtMP_QtyOptimumStock"
        Me.TxtMP_QtyOptimumStock.ReadOnly = True
        Me.TxtMP_QtyOptimumStock.Size = New System.Drawing.Size(155, 20)
        Me.TxtMP_QtyOptimumStock.TabIndex = 3
        Me.TxtMP_QtyOptimumStock.Text = "0"
        Me.TxtMP_QtyOptimumStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(3, 55)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(65, 13)
        Me.Label14.TabIndex = 59
        Me.Label14.Text = "Pending DO"
        '
        'TxtMP_QtyPendingDO
        '
        Me.TxtMP_QtyPendingDO.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPendingDO.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPendingDO.Location = New System.Drawing.Point(101, 52)
        Me.TxtMP_QtyPendingDO.Name = "TxtMP_QtyPendingDO"
        Me.TxtMP_QtyPendingDO.ReadOnly = True
        Me.TxtMP_QtyPendingDO.Size = New System.Drawing.Size(155, 20)
        Me.TxtMP_QtyPendingDO.TabIndex = 2
        Me.TxtMP_QtyPendingDO.Text = "0"
        Me.TxtMP_QtyPendingDO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(3, 104)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(79, 13)
        Me.Label17.TabIndex = 65
        Me.Label17.Text = "Optimum Stock"
        '
        'TxtMP_QtyFree
        '
        Me.TxtMP_QtyFree.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyFree.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyFree.Location = New System.Drawing.Point(200, 4)
        Me.TxtMP_QtyFree.Name = "TxtMP_QtyFree"
        Me.TxtMP_QtyFree.ReadOnly = True
        Me.TxtMP_QtyFree.Size = New System.Drawing.Size(91, 20)
        Me.TxtMP_QtyFree.TabIndex = 76
        Me.TxtMP_QtyFree.Text = "0"
        Me.TxtMP_QtyFree.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label39
        '
        Me.Label39.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(152, 6)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(47, 13)
        Me.Label39.TabIndex = 77
        Me.Label39.Text = "Qty Free"
        '
        'TxtMP_QtyMOB
        '
        Me.TxtMP_QtyMOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyMOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyMOB.Location = New System.Drawing.Point(107, 75)
        Me.TxtMP_QtyMOB.Name = "TxtMP_QtyMOB"
        Me.TxtMP_QtyMOB.ReadOnly = True
        Me.TxtMP_QtyMOB.Size = New System.Drawing.Size(183, 20)
        Me.TxtMP_QtyMOB.TabIndex = 74
        Me.TxtMP_QtyMOB.Text = "0"
        Me.TxtMP_QtyMOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(3, 78)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(100, 13)
        Me.Label36.TabIndex = 75
        Me.Label36.Text = "Production Balance"
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(3, 102)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(81, 13)
        Me.Label30.TabIndex = 73
        Me.Label30.Text = "Avg. Per Month"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 104)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 13)
        Me.Label6.TabIndex = 71
        Me.Label6.Text = "Order Strategy"
        Me.Label6.Visible = False
        '
        'TxtMP_QtyRequired
        '
        Me.TxtMP_QtyRequired.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRequired.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyRequired.Location = New System.Drawing.Point(107, 126)
        Me.TxtMP_QtyRequired.Name = "TxtMP_QtyRequired"
        Me.TxtMP_QtyRequired.ReadOnly = True
        Me.TxtMP_QtyRequired.Size = New System.Drawing.Size(183, 20)
        Me.TxtMP_QtyRequired.TabIndex = 6
        Me.TxtMP_QtyRequired.Text = "0"
        Me.TxtMP_QtyRequired.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_QtyOnHand
        '
        Me.TxtMP_QtyOnHand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyOnHand.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyOnHand.Location = New System.Drawing.Point(107, 4)
        Me.TxtMP_QtyOnHand.Name = "TxtMP_QtyOnHand"
        Me.TxtMP_QtyOnHand.ReadOnly = True
        Me.TxtMP_QtyOnHand.Size = New System.Drawing.Size(42, 20)
        Me.TxtMP_QtyOnHand.TabIndex = 0
        Me.TxtMP_QtyOnHand.Text = "0"
        Me.TxtMP_QtyOnHand.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(3, 128)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(50, 13)
        Me.Label18.TabIndex = 69
        Me.Label18.Text = "Required"
        '
        'TxtMP_QtyPOB
        '
        Me.TxtMP_QtyPOB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyPOB.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyPOB.Location = New System.Drawing.Point(107, 52)
        Me.TxtMP_QtyPOB.Name = "TxtMP_QtyPOB"
        Me.TxtMP_QtyPOB.ReadOnly = True
        Me.TxtMP_QtyPOB.Size = New System.Drawing.Size(183, 20)
        Me.TxtMP_QtyPOB.TabIndex = 2
        Me.TxtMP_QtyPOB.Text = "0"
        Me.TxtMP_QtyPOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 55)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(94, 13)
        Me.Label10.TabIndex = 53
        Me.Label10.Text = "Purchase Balance"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(3, 6)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(50, 13)
        Me.Label15.TabIndex = 63
        Me.Label15.Text = "On Hand"
        '
        'TxtMP_QtyGIT
        '
        Me.TxtMP_QtyGIT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyGIT.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_QtyGIT.Location = New System.Drawing.Point(107, 28)
        Me.TxtMP_QtyGIT.Name = "TxtMP_QtyGIT"
        Me.TxtMP_QtyGIT.ReadOnly = True
        Me.TxtMP_QtyGIT.Size = New System.Drawing.Size(183, 20)
        Me.TxtMP_QtyGIT.TabIndex = 1
        Me.TxtMP_QtyGIT.Text = "0"
        Me.TxtMP_QtyGIT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 30)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(85, 13)
        Me.Label9.TabIndex = 51
        Me.Label9.Text = "Goods In Transit"
        '
        'TxtMP_AvgPerMonth
        '
        Me.TxtMP_AvgPerMonth.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_AvgPerMonth.Cursor = System.Windows.Forms.Cursors.Hand
        Me.TxtMP_AvgPerMonth.Location = New System.Drawing.Point(107, 100)
        Me.TxtMP_AvgPerMonth.Name = "TxtMP_AvgPerMonth"
        Me.TxtMP_AvgPerMonth.ReadOnly = True
        Me.TxtMP_AvgPerMonth.Size = New System.Drawing.Size(183, 20)
        Me.TxtMP_AvgPerMonth.TabIndex = 3
        Me.TxtMP_AvgPerMonth.Text = "0"
        Me.TxtMP_AvgPerMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboMPOrderStrategy
        '
        Me.cboMPOrderStrategy.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboMPOrderStrategy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMPOrderStrategy.FormattingEnabled = True
        Me.cboMPOrderStrategy.Items.AddRange(New Object() {"Order To Min Qty", "Order To Optimum Qty", "Order To Max Qty", "Order To Safety Stock", "Order By Min Order Qty", "Order By Opt Order Qty", "Order By Max Order Qty"})
        Me.cboMPOrderStrategy.Location = New System.Drawing.Point(107, 100)
        Me.cboMPOrderStrategy.Margin = New System.Windows.Forms.Padding(2)
        Me.cboMPOrderStrategy.Name = "cboMPOrderStrategy"
        Me.cboMPOrderStrategy.Size = New System.Drawing.Size(183, 21)
        Me.cboMPOrderStrategy.TabIndex = 5
        Me.cboMPOrderStrategy.Visible = False
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.SplitContainer3)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage5.Size = New System.Drawing.Size(561, 149)
        Me.TabPage5.TabIndex = 1
        Me.TabPage5.Text = "Detail Purchase"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(2, 2)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Panel2Collapsed = True
        Me.SplitContainer3.Size = New System.Drawing.Size(557, 145)
        Me.SplitContainer3.SplitterDistance = 304
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 72
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_DiscFromPerct)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_DiscAmount)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label32)
        Me.SplitContainer4.Panel1.Controls.Add(Me.LlbMP_UnitPrice)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_CurrencyCode)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_UnitPrice)
        Me.SplitContainer4.Panel1.Controls.Add(Me.btnMP_FindCurrency)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label22)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_PreferedSupplier)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_Discount)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label20)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label19)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_NettPrice)
        Me.SplitContainer4.Panel1.Controls.Add(Me.btnMP_FindSupplier)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_PrimaryCurrency)
        Me.SplitContainer4.Panel1.Controls.Add(Me.TxtMP_PrimarySupplier)
        Me.SplitContainer4.Panel1.Controls.Add(Me.Label13)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.llbMP_Tolerance)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_PreferedSupplierDesc)
        Me.SplitContainer4.Panel2.Controls.Add(Me.llbMP_Ratio)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_QtyRequired2)
        Me.SplitContainer4.Panel2.Controls.Add(Me.LblMP_Unit1)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_QtyOrder)
        Me.SplitContainer4.Panel2.Controls.Add(Me.LblRatio)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label24)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label25)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_SubTotal)
        Me.SplitContainer4.Panel2.Controls.Add(Me.TxtMP_PrimaryOrderUnit)
        Me.SplitContainer4.Panel2.Controls.Add(Me.cboMP_OrderUnit)
        Me.SplitContainer4.Panel2.Controls.Add(Me.Label23)
        Me.SplitContainer4.Size = New System.Drawing.Size(563, 124)
        Me.SplitContainer4.SplitterDistance = 266
        Me.SplitContainer4.SplitterWidth = 3
        Me.SplitContainer4.TabIndex = 83
        Me.SplitContainer4.TabStop = False
        '
        'TxtMP_DiscFromPerct
        '
        Me.TxtMP_DiscFromPerct.Location = New System.Drawing.Point(71, 76)
        Me.TxtMP_DiscFromPerct.Name = "TxtMP_DiscFromPerct"
        Me.TxtMP_DiscFromPerct.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_DiscFromPerct.TabIndex = 74
        Me.TxtMP_DiscFromPerct.TabStop = False
        Me.TxtMP_DiscFromPerct.Text = "1"
        Me.TxtMP_DiscFromPerct.Visible = False
        '
        'TxtMP_DiscAmount
        '
        Me.TxtMP_DiscAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_DiscAmount.Location = New System.Drawing.Point(165, 75)
        Me.TxtMP_DiscAmount.Name = "TxtMP_DiscAmount"
        Me.TxtMP_DiscAmount.Size = New System.Drawing.Size(96, 20)
        Me.TxtMP_DiscAmount.TabIndex = 8
        Me.TxtMP_DiscAmount.Text = "0"
        Me.TxtMP_DiscAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(142, 77)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(15, 13)
        Me.Label32.TabIndex = 73
        Me.Label32.Text = "%"
        '
        'LlbMP_UnitPrice
        '
        Me.LlbMP_UnitPrice.AutoSize = True
        Me.LlbMP_UnitPrice.Location = New System.Drawing.Point(3, 52)
        Me.LlbMP_UnitPrice.Name = "LlbMP_UnitPrice"
        Me.LlbMP_UnitPrice.Size = New System.Drawing.Size(53, 13)
        Me.LlbMP_UnitPrice.TabIndex = 5
        Me.LlbMP_UnitPrice.TabStop = True
        Me.LlbMP_UnitPrice.Text = "Unit Price"
        Me.LlbMP_UnitPrice.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtMP_CurrencyCode
        '
        Me.TxtMP_CurrencyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_CurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMP_CurrencyCode.Location = New System.Drawing.Point(98, 28)
        Me.TxtMP_CurrencyCode.Name = "TxtMP_CurrencyCode"
        Me.TxtMP_CurrencyCode.Size = New System.Drawing.Size(114, 20)
        Me.TxtMP_CurrencyCode.TabIndex = 3
        '
        'TxtMP_UnitPrice
        '
        Me.TxtMP_UnitPrice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_UnitPrice.Location = New System.Drawing.Point(98, 50)
        Me.TxtMP_UnitPrice.Name = "TxtMP_UnitPrice"
        Me.TxtMP_UnitPrice.Size = New System.Drawing.Size(162, 20)
        Me.TxtMP_UnitPrice.TabIndex = 5
        Me.TxtMP_UnitPrice.Text = "0"
        Me.TxtMP_UnitPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnMP_FindCurrency
        '
        Me.btnMP_FindCurrency.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMP_FindCurrency.Location = New System.Drawing.Point(213, 26)
        Me.btnMP_FindCurrency.Name = "btnMP_FindCurrency"
        Me.btnMP_FindCurrency.Size = New System.Drawing.Size(46, 20)
        Me.btnMP_FindCurrency.TabIndex = 3
        Me.btnMP_FindCurrency.TabStop = False
        Me.btnMP_FindCurrency.Text = "Find"
        Me.btnMP_FindCurrency.UseVisualStyleBackColor = True
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(3, 77)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(49, 13)
        Me.Label22.TabIndex = 7
        Me.Label22.Text = "Discount"
        '
        'TxtMP_PreferedSupplier
        '
        Me.TxtMP_PreferedSupplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_PreferedSupplier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMP_PreferedSupplier.Location = New System.Drawing.Point(98, 4)
        Me.TxtMP_PreferedSupplier.Name = "TxtMP_PreferedSupplier"
        Me.TxtMP_PreferedSupplier.Size = New System.Drawing.Size(114, 20)
        Me.TxtMP_PreferedSupplier.TabIndex = 1
        '
        'TxtMP_Discount
        '
        Me.TxtMP_Discount.Location = New System.Drawing.Point(98, 75)
        Me.TxtMP_Discount.Name = "TxtMP_Discount"
        Me.TxtMP_Discount.Size = New System.Drawing.Size(39, 20)
        Me.TxtMP_Discount.TabIndex = 6
        Me.TxtMP_Discount.Text = "0.00"
        Me.TxtMP_Discount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(3, 102)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(54, 13)
        Me.Label20.TabIndex = 67
        Me.Label20.Text = "Nett Price"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(3, 28)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(49, 13)
        Me.Label19.TabIndex = 57
        Me.Label19.Text = "Currency"
        '
        'TxtMP_NettPrice
        '
        Me.TxtMP_NettPrice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_NettPrice.Location = New System.Drawing.Point(98, 99)
        Me.TxtMP_NettPrice.Name = "TxtMP_NettPrice"
        Me.TxtMP_NettPrice.ReadOnly = True
        Me.TxtMP_NettPrice.Size = New System.Drawing.Size(162, 20)
        Me.TxtMP_NettPrice.TabIndex = 9
        Me.TxtMP_NettPrice.TabStop = False
        Me.TxtMP_NettPrice.Text = "0"
        Me.TxtMP_NettPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnMP_FindSupplier
        '
        Me.btnMP_FindSupplier.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMP_FindSupplier.Location = New System.Drawing.Point(213, 6)
        Me.btnMP_FindSupplier.Name = "btnMP_FindSupplier"
        Me.btnMP_FindSupplier.Size = New System.Drawing.Size(46, 20)
        Me.btnMP_FindSupplier.TabIndex = 1
        Me.btnMP_FindSupplier.TabStop = False
        Me.btnMP_FindSupplier.Text = "Find"
        Me.btnMP_FindSupplier.UseVisualStyleBackColor = True
        '
        'TxtMP_PrimaryCurrency
        '
        Me.TxtMP_PrimaryCurrency.Location = New System.Drawing.Point(86, 28)
        Me.TxtMP_PrimaryCurrency.Name = "TxtMP_PrimaryCurrency"
        Me.TxtMP_PrimaryCurrency.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_PrimaryCurrency.TabIndex = 72
        Me.TxtMP_PrimaryCurrency.TabStop = False
        Me.TxtMP_PrimaryCurrency.Text = "-1"
        Me.TxtMP_PrimaryCurrency.Visible = False
        '
        'TxtMP_PrimarySupplier
        '
        Me.TxtMP_PrimarySupplier.Location = New System.Drawing.Point(86, 4)
        Me.TxtMP_PrimarySupplier.Name = "TxtMP_PrimarySupplier"
        Me.TxtMP_PrimarySupplier.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_PrimarySupplier.TabIndex = 69
        Me.TxtMP_PrimarySupplier.TabStop = False
        Me.TxtMP_PrimarySupplier.Text = "-1"
        Me.TxtMP_PrimarySupplier.Visible = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 6)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(88, 13)
        Me.Label13.TabIndex = 55
        Me.Label13.Text = "Prefered Supplier"
        '
        'llbMP_Tolerance
        '
        Me.llbMP_Tolerance.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbMP_Tolerance.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbMP_Tolerance.AutoSize = True
        Me.llbMP_Tolerance.Location = New System.Drawing.Point(228, 77)
        Me.llbMP_Tolerance.Name = "llbMP_Tolerance"
        Me.llbMP_Tolerance.Size = New System.Drawing.Size(55, 13)
        Me.llbMP_Tolerance.TabIndex = 82
        Me.llbMP_Tolerance.TabStop = True
        Me.llbMP_Tolerance.Text = "Tolerance"
        Me.llbMP_Tolerance.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtMP_PreferedSupplierDesc
        '
        Me.TxtMP_PreferedSupplierDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_PreferedSupplierDesc.Location = New System.Drawing.Point(3, 4)
        Me.TxtMP_PreferedSupplierDesc.Name = "TxtMP_PreferedSupplierDesc"
        Me.TxtMP_PreferedSupplierDesc.ReadOnly = True
        Me.TxtMP_PreferedSupplierDesc.Size = New System.Drawing.Size(283, 20)
        Me.TxtMP_PreferedSupplierDesc.TabIndex = 0
        Me.TxtMP_PreferedSupplierDesc.TabStop = False
        '
        'llbMP_Ratio
        '
        Me.llbMP_Ratio.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbMP_Ratio.AutoSize = True
        Me.llbMP_Ratio.Location = New System.Drawing.Point(77, 53)
        Me.llbMP_Ratio.Name = "llbMP_Ratio"
        Me.llbMP_Ratio.Size = New System.Drawing.Size(32, 13)
        Me.llbMP_Ratio.TabIndex = 3
        Me.llbMP_Ratio.TabStop = True
        Me.llbMP_Ratio.Text = "Ratio"
        Me.llbMP_Ratio.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtMP_QtyRequired2
        '
        Me.TxtMP_QtyRequired2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyRequired2.Location = New System.Drawing.Point(77, 30)
        Me.TxtMP_QtyRequired2.Name = "TxtMP_QtyRequired2"
        Me.TxtMP_QtyRequired2.ReadOnly = True
        Me.TxtMP_QtyRequired2.Size = New System.Drawing.Size(145, 20)
        Me.TxtMP_QtyRequired2.TabIndex = 1
        Me.TxtMP_QtyRequired2.TabStop = False
        Me.TxtMP_QtyRequired2.Text = "0"
        Me.TxtMP_QtyRequired2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblMP_Unit1
        '
        Me.LblMP_Unit1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblMP_Unit1.AutoSize = True
        Me.LblMP_Unit1.Location = New System.Drawing.Point(228, 30)
        Me.LblMP_Unit1.Name = "LblMP_Unit1"
        Me.LblMP_Unit1.Size = New System.Drawing.Size(26, 13)
        Me.LblMP_Unit1.TabIndex = 2
        Me.LblMP_Unit1.Text = "Unit"
        '
        'TxtMP_QtyOrder
        '
        Me.TxtMP_QtyOrder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_QtyOrder.Location = New System.Drawing.Point(77, 75)
        Me.TxtMP_QtyOrder.Name = "TxtMP_QtyOrder"
        Me.TxtMP_QtyOrder.Size = New System.Drawing.Size(70, 20)
        Me.TxtMP_QtyOrder.TabIndex = 4
        Me.TxtMP_QtyOrder.Text = "0"
        Me.TxtMP_QtyOrder.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblRatio
        '
        Me.LblRatio.AutoSize = True
        Me.LblRatio.Location = New System.Drawing.Point(3, 52)
        Me.LblRatio.Name = "LblRatio"
        Me.LblRatio.Size = New System.Drawing.Size(32, 13)
        Me.LblRatio.TabIndex = 81
        Me.LblRatio.Text = "Ratio"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(3, 102)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(53, 13)
        Me.Label24.TabIndex = 75
        Me.Label24.Text = "Sub Total"
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(3, 30)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(69, 13)
        Me.Label25.TabIndex = 79
        Me.Label25.Text = "Required Qty"
        '
        'TxtMP_SubTotal
        '
        Me.TxtMP_SubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMP_SubTotal.Location = New System.Drawing.Point(77, 99)
        Me.TxtMP_SubTotal.Name = "TxtMP_SubTotal"
        Me.TxtMP_SubTotal.ReadOnly = True
        Me.TxtMP_SubTotal.Size = New System.Drawing.Size(211, 20)
        Me.TxtMP_SubTotal.TabIndex = 6
        Me.TxtMP_SubTotal.TabStop = False
        Me.TxtMP_SubTotal.Text = "0"
        Me.TxtMP_SubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtMP_PrimaryOrderUnit
        '
        Me.TxtMP_PrimaryOrderUnit.Location = New System.Drawing.Point(50, 76)
        Me.TxtMP_PrimaryOrderUnit.Name = "TxtMP_PrimaryOrderUnit"
        Me.TxtMP_PrimaryOrderUnit.Size = New System.Drawing.Size(22, 20)
        Me.TxtMP_PrimaryOrderUnit.TabIndex = 78
        Me.TxtMP_PrimaryOrderUnit.TabStop = False
        Me.TxtMP_PrimaryOrderUnit.Text = "-1"
        Me.TxtMP_PrimaryOrderUnit.Visible = False
        '
        'cboMP_OrderUnit
        '
        Me.cboMP_OrderUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboMP_OrderUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMP_OrderUnit.FormattingEnabled = True
        Me.cboMP_OrderUnit.Location = New System.Drawing.Point(153, 74)
        Me.cboMP_OrderUnit.Margin = New System.Windows.Forms.Padding(2)
        Me.cboMP_OrderUnit.Name = "cboMP_OrderUnit"
        Me.cboMP_OrderUnit.Size = New System.Drawing.Size(69, 21)
        Me.cboMP_OrderUnit.TabIndex = 5
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(3, 77)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(52, 13)
        Me.Label23.TabIndex = 73
        Me.Label23.Text = "Order Qty"
        '
        'btnMP_Find
        '
        Me.btnMP_Find.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnMP_Find.Location = New System.Drawing.Point(100, 456)
        Me.btnMP_Find.Name = "btnMP_Find"
        Me.btnMP_Find.Size = New System.Drawing.Size(69, 23)
        Me.btnMP_Find.TabIndex = 33
        Me.btnMP_Find.Text = "Find"
        Me.btnMP_Find.UseVisualStyleBackColor = True
        '
        'btnMP_Delete
        '
        Me.btnMP_Delete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnMP_Delete.Location = New System.Drawing.Point(326, 456)
        Me.btnMP_Delete.Name = "btnMP_Delete"
        Me.btnMP_Delete.Size = New System.Drawing.Size(69, 23)
        Me.btnMP_Delete.TabIndex = 35
        Me.btnMP_Delete.Text = "Delete"
        Me.btnMP_Delete.UseVisualStyleBackColor = True
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.btnSF_Edit)
        Me.TabPage3.Controls.Add(Me.btnSF_New)
        Me.TabPage3.Controls.Add(Me.btnSF_Find)
        Me.TabPage3.Controls.Add(Me.btnSF_Print)
        Me.TabPage3.Controls.Add(Me.btnSF_Delete)
        Me.TabPage3.Controls.Add(Me.StatusStrip3)
        Me.TabPage3.Controls.Add(Me.gboSF)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(577, 505)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Shipping Forecast"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'btnSF_Edit
        '
        Me.btnSF_Edit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSF_Edit.Location = New System.Drawing.Point(257, 470)
        Me.btnSF_Edit.Name = "btnSF_Edit"
        Me.btnSF_Edit.Size = New System.Drawing.Size(69, 23)
        Me.btnSF_Edit.TabIndex = 62
        Me.btnSF_Edit.Text = "Edit"
        Me.btnSF_Edit.UseVisualStyleBackColor = True
        '
        'btnSF_New
        '
        Me.btnSF_New.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSF_New.Location = New System.Drawing.Point(182, 470)
        Me.btnSF_New.Name = "btnSF_New"
        Me.btnSF_New.Size = New System.Drawing.Size(69, 23)
        Me.btnSF_New.TabIndex = 61
        Me.btnSF_New.Text = "New"
        Me.btnSF_New.UseVisualStyleBackColor = True
        '
        'btnSF_Find
        '
        Me.btnSF_Find.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSF_Find.Location = New System.Drawing.Point(107, 470)
        Me.btnSF_Find.Name = "btnSF_Find"
        Me.btnSF_Find.Size = New System.Drawing.Size(69, 23)
        Me.btnSF_Find.TabIndex = 59
        Me.btnSF_Find.Text = "Find"
        Me.btnSF_Find.UseVisualStyleBackColor = True
        '
        'btnSF_Print
        '
        Me.btnSF_Print.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSF_Print.Location = New System.Drawing.Point(407, 470)
        Me.btnSF_Print.Name = "btnSF_Print"
        Me.btnSF_Print.Size = New System.Drawing.Size(69, 23)
        Me.btnSF_Print.TabIndex = 58
        Me.btnSF_Print.Text = "Print"
        Me.btnSF_Print.UseVisualStyleBackColor = True
        '
        'btnSF_Delete
        '
        Me.btnSF_Delete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSF_Delete.Location = New System.Drawing.Point(332, 470)
        Me.btnSF_Delete.Name = "btnSF_Delete"
        Me.btnSF_Delete.Size = New System.Drawing.Size(69, 23)
        Me.btnSF_Delete.TabIndex = 60
        Me.btnSF_Delete.Text = "Delete"
        Me.btnSF_Delete.UseVisualStyleBackColor = True
        '
        'StatusStrip3
        '
        Me.StatusStrip3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslSF_CreatedBy, Me.tslSF_CreatedDate, Me.tslSF_ModifiedBy, Me.tslSF_ModifiedDate})
        Me.StatusStrip3.Location = New System.Drawing.Point(2, 475)
        Me.StatusStrip3.Name = "StatusStrip3"
        Me.StatusStrip3.Size = New System.Drawing.Size(551, 22)
        Me.StatusStrip3.TabIndex = 54
        Me.StatusStrip3.Text = "StatusStrip3"
        Me.StatusStrip3.Visible = False
        '
        'tslSF_CreatedBy
        '
        Me.tslSF_CreatedBy.Name = "tslSF_CreatedBy"
        Me.tslSF_CreatedBy.Size = New System.Drawing.Size(134, 17)
        Me.tslSF_CreatedBy.Spring = True
        Me.tslSF_CreatedBy.Text = "Created By:"
        '
        'tslSF_CreatedDate
        '
        Me.tslSF_CreatedDate.Name = "tslSF_CreatedDate"
        Me.tslSF_CreatedDate.Size = New System.Drawing.Size(134, 17)
        Me.tslSF_CreatedDate.Spring = True
        Me.tslSF_CreatedDate.Text = "Created Date:"
        '
        'tslSF_ModifiedBy
        '
        Me.tslSF_ModifiedBy.Name = "tslSF_ModifiedBy"
        Me.tslSF_ModifiedBy.Size = New System.Drawing.Size(134, 17)
        Me.tslSF_ModifiedBy.Spring = True
        Me.tslSF_ModifiedBy.Text = "Modified By:"
        '
        'tslSF_ModifiedDate
        '
        Me.tslSF_ModifiedDate.Name = "tslSF_ModifiedDate"
        Me.tslSF_ModifiedDate.Size = New System.Drawing.Size(134, 17)
        Me.tslSF_ModifiedDate.Spring = True
        Me.tslSF_ModifiedDate.Text = "Modified Date:"
        '
        'gboSF
        '
        Me.gboSF.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboSF.Controls.Add(Me.llbSF_CekAll)
        Me.gboSF.Controls.Add(Me.btnSF_Cancel)
        Me.gboSF.Controls.Add(Me.btnSF_OK)
        Me.gboSF.Controls.Add(Me.llbSF_CreateDO)
        Me.gboSF.Controls.Add(Me.dgvSF)
        Me.gboSF.Controls.Add(Me.SplitContainer6)
        Me.gboSF.Location = New System.Drawing.Point(4, 0)
        Me.gboSF.Margin = New System.Windows.Forms.Padding(2)
        Me.gboSF.Name = "gboSF"
        Me.gboSF.Padding = New System.Windows.Forms.Padding(2)
        Me.gboSF.Size = New System.Drawing.Size(563, 465)
        Me.gboSF.TabIndex = 87
        Me.gboSF.TabStop = False
        '
        'llbSF_CekAll
        '
        Me.llbSF_CekAll.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbSF_CekAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbSF_CekAll.AutoSize = True
        Me.llbSF_CekAll.Location = New System.Drawing.Point(4, 442)
        Me.llbSF_CekAll.Name = "llbSF_CekAll"
        Me.llbSF_CekAll.Size = New System.Drawing.Size(52, 13)
        Me.llbSF_CekAll.TabIndex = 88
        Me.llbSF_CekAll.TabStop = True
        Me.llbSF_CekAll.Text = "Check All"
        Me.llbSF_CekAll.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'btnSF_Cancel
        '
        Me.btnSF_Cancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSF_Cancel.Location = New System.Drawing.Point(284, 436)
        Me.btnSF_Cancel.Name = "btnSF_Cancel"
        Me.btnSF_Cancel.Size = New System.Drawing.Size(69, 23)
        Me.btnSF_Cancel.TabIndex = 40
        Me.btnSF_Cancel.Text = "Cancel"
        Me.btnSF_Cancel.UseVisualStyleBackColor = True
        '
        'btnSF_OK
        '
        Me.btnSF_OK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnSF_OK.Location = New System.Drawing.Point(210, 436)
        Me.btnSF_OK.Name = "btnSF_OK"
        Me.btnSF_OK.Size = New System.Drawing.Size(69, 23)
        Me.btnSF_OK.TabIndex = 39
        Me.btnSF_OK.Text = "OK"
        Me.btnSF_OK.UseVisualStyleBackColor = True
        '
        'llbSF_CreateDO
        '
        Me.llbSF_CreateDO.ActiveLinkColor = System.Drawing.Color.Blue
        Me.llbSF_CreateDO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSF_CreateDO.AutoSize = True
        Me.llbSF_CreateDO.Location = New System.Drawing.Point(430, 442)
        Me.llbSF_CreateDO.Name = "llbSF_CreateDO"
        Me.llbSF_CreateDO.Size = New System.Drawing.Size(129, 13)
        Me.llbSF_CreateDO.TabIndex = 86
        Me.llbSF_CreateDO.TabStop = True
        Me.llbSF_CreateDO.Text = "Create DO/Pick Up Order"
        Me.llbSF_CreateDO.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'dgvSF
        '
        Me.dgvSF.AllowUserToAddRows = False
        Me.dgvSF.AllowUserToDeleteRows = False
        Me.dgvSF.AllowUserToResizeRows = False
        Me.dgvSF.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSF.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSF.ColumnHeadersHeight = 25
        Me.dgvSF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSF.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSF_No, Me.ColSF_Type, Me.ColSF_SONumber, Me.ColSF_CustomerPONumber, Me.ColSF_SODate, Me.ColSF_Customer, Me.ColSF_CodeNo, Me.ColSF_Desc, Me.ColSF_QtySOB, Me.ColSF_QtyAvailable, Me.ColSF_Unit, Me.ColSF_PromiseToDeliver, Me.ColSF_PromiseToDeliver_Str, Me.ColSF_EstimateDelivery, Me.ColSF_EstimateDelivery_Str, Me.ColSF_CreateDO, Me.ColSF_PrimaryCodeNo, Me.ColSF_DeliveryStatus, Me.ColSF_PrimaryUnit, Me.ColSF_PrimarySO, Me.ColSF_Primary, Me.ColSF_PrimaryCust, Me.ColSF_QtyReadyStock, Me.ColSF_TypeInt, Me.ColSF_Remarks})
        Me.dgvSF.Location = New System.Drawing.Point(6, 76)
        Me.dgvSF.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvSF.Name = "dgvSF"
        Me.dgvSF.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvSF.RowTemplate.Height = 24
        Me.dgvSF.Size = New System.Drawing.Size(551, 355)
        Me.dgvSF.TabIndex = 38
        '
        'ColSF_No
        '
        Me.ColSF_No.HeaderText = "No"
        Me.ColSF_No.Name = "ColSF_No"
        Me.ColSF_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSF_No.Width = 27
        '
        'ColSF_Type
        '
        Me.ColSF_Type.HeaderText = "Type"
        Me.ColSF_Type.Name = "ColSF_Type"
        Me.ColSF_Type.Width = 56
        '
        'ColSF_SONumber
        '
        Me.ColSF_SONumber.HeaderText = "Document No"
        Me.ColSF_SONumber.Name = "ColSF_SONumber"
        Me.ColSF_SONumber.Visible = False
        Me.ColSF_SONumber.Width = 98
        '
        'ColSF_CustomerPONumber
        '
        Me.ColSF_CustomerPONumber.HeaderText = "Cust PO Number"
        Me.ColSF_CustomerPONumber.Name = "ColSF_CustomerPONumber"
        Me.ColSF_CustomerPONumber.Visible = False
        Me.ColSF_CustomerPONumber.Width = 111
        '
        'ColSF_SODate
        '
        Me.ColSF_SODate.HeaderText = "Date"
        Me.ColSF_SODate.Name = "ColSF_SODate"
        Me.ColSF_SODate.Width = 55
        '
        'ColSF_Customer
        '
        Me.ColSF_Customer.HeaderText = "Customer/Ref Party"
        Me.ColSF_Customer.Name = "ColSF_Customer"
        Me.ColSF_Customer.Width = 125
        '
        'ColSF_CodeNo
        '
        Me.ColSF_CodeNo.HeaderText = "Code No"
        Me.ColSF_CodeNo.Name = "ColSF_CodeNo"
        Me.ColSF_CodeNo.Width = 74
        '
        'ColSF_Desc
        '
        Me.ColSF_Desc.HeaderText = "Description"
        Me.ColSF_Desc.Name = "ColSF_Desc"
        Me.ColSF_Desc.Visible = False
        Me.ColSF_Desc.Width = 85
        '
        'ColSF_QtySOB
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSF_QtySOB.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColSF_QtySOB.HeaderText = "Order Balance Qty"
        Me.ColSF_QtySOB.Name = "ColSF_QtySOB"
        Me.ColSF_QtySOB.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColSF_QtyAvailable
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSF_QtyAvailable.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColSF_QtyAvailable.HeaderText = "Available Qty"
        Me.ColSF_QtyAvailable.Name = "ColSF_QtyAvailable"
        Me.ColSF_QtyAvailable.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSF_QtyAvailable.Width = 75
        '
        'ColSF_Unit
        '
        Me.ColSF_Unit.HeaderText = "Unit"
        Me.ColSF_Unit.Name = "ColSF_Unit"
        Me.ColSF_Unit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSF_Unit.Width = 32
        '
        'ColSF_PromiseToDeliver
        '
        Me.ColSF_PromiseToDeliver.HeaderText = "Promise To Deliver"
        Me.ColSF_PromiseToDeliver.Name = "ColSF_PromiseToDeliver"
        Me.ColSF_PromiseToDeliver.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSF_PromiseToDeliver.Visible = False
        Me.ColSF_PromiseToDeliver.Width = 102
        '
        'ColSF_PromiseToDeliver_Str
        '
        Me.ColSF_PromiseToDeliver_Str.HeaderText = "Promise To Deliver"
        Me.ColSF_PromiseToDeliver_Str.Name = "ColSF_PromiseToDeliver_Str"
        Me.ColSF_PromiseToDeliver_Str.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSF_PromiseToDeliver_Str.Width = 121
        '
        'ColSF_EstimateDelivery
        '
        Me.ColSF_EstimateDelivery.HeaderText = "Estimate Delivery"
        Me.ColSF_EstimateDelivery.Name = "ColSF_EstimateDelivery"
        Me.ColSF_EstimateDelivery.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColSF_EstimateDelivery.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSF_EstimateDelivery.Visible = False
        Me.ColSF_EstimateDelivery.Width = 94
        '
        'ColSF_EstimateDelivery_Str
        '
        Me.ColSF_EstimateDelivery_Str.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColSF_EstimateDelivery_Str.HeaderText = "Available On"
        Me.ColSF_EstimateDelivery_Str.Name = "ColSF_EstimateDelivery_Str"
        Me.ColSF_EstimateDelivery_Str.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSF_EstimateDelivery_Str.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColSF_EstimateDelivery_Str.Width = 92
        '
        'ColSF_CreateDO
        '
        Me.ColSF_CreateDO.HeaderText = "Follow Up"
        Me.ColSF_CreateDO.Name = "ColSF_CreateDO"
        Me.ColSF_CreateDO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColSF_CreateDO.Width = 79
        '
        'ColSF_PrimaryCodeNo
        '
        Me.ColSF_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColSF_PrimaryCodeNo.Name = "ColSF_PrimaryCodeNo"
        Me.ColSF_PrimaryCodeNo.Visible = False
        Me.ColSF_PrimaryCodeNo.Width = 111
        '
        'ColSF_DeliveryStatus
        '
        Me.ColSF_DeliveryStatus.HeaderText = "Delivery Status"
        Me.ColSF_DeliveryStatus.Name = "ColSF_DeliveryStatus"
        Me.ColSF_DeliveryStatus.Width = 103
        '
        'ColSF_PrimaryUnit
        '
        Me.ColSF_PrimaryUnit.HeaderText = "Primary Unit"
        Me.ColSF_PrimaryUnit.Name = "ColSF_PrimaryUnit"
        Me.ColSF_PrimaryUnit.Visible = False
        Me.ColSF_PrimaryUnit.Width = 88
        '
        'ColSF_PrimarySO
        '
        Me.ColSF_PrimarySO.HeaderText = "Primary SO"
        Me.ColSF_PrimarySO.Name = "ColSF_PrimarySO"
        Me.ColSF_PrimarySO.Visible = False
        Me.ColSF_PrimarySO.Width = 84
        '
        'ColSF_Primary
        '
        Me.ColSF_Primary.HeaderText = "Primary"
        Me.ColSF_Primary.Name = "ColSF_Primary"
        Me.ColSF_Primary.Visible = False
        Me.ColSF_Primary.Width = 66
        '
        'ColSF_PrimaryCust
        '
        Me.ColSF_PrimaryCust.HeaderText = "Primary Cust"
        Me.ColSF_PrimaryCust.Name = "ColSF_PrimaryCust"
        Me.ColSF_PrimaryCust.Visible = False
        Me.ColSF_PrimaryCust.Width = 90
        '
        'ColSF_QtyReadyStock
        '
        Me.ColSF_QtyReadyStock.HeaderText = "Qty Ready Stock"
        Me.ColSF_QtyReadyStock.Name = "ColSF_QtyReadyStock"
        Me.ColSF_QtyReadyStock.Visible = False
        Me.ColSF_QtyReadyStock.Width = 113
        '
        'ColSF_TypeInt
        '
        Me.ColSF_TypeInt.HeaderText = "Type_Int"
        Me.ColSF_TypeInt.Name = "ColSF_TypeInt"
        Me.ColSF_TypeInt.Visible = False
        Me.ColSF_TypeInt.Width = 74
        '
        'ColSF_Remarks
        '
        Me.ColSF_Remarks.HeaderText = "Remarks"
        Me.ColSF_Remarks.Name = "ColSF_Remarks"
        Me.ColSF_Remarks.Width = 74
        '
        'SplitContainer6
        '
        Me.SplitContainer6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer6.IsSplitterFixed = True
        Me.SplitContainer6.Location = New System.Drawing.Point(0, 4)
        Me.SplitContainer6.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer6.Name = "SplitContainer6"
        '
        'SplitContainer6.Panel1
        '
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtSF_StockInPurchaseNo)
        Me.SplitContainer6.Panel1.Controls.Add(Me.LblSF_StockInPurchaseNo)
        Me.SplitContainer6.Panel1.Controls.Add(Me.llbSF_Calculate)
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtSF_Urutan)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer6.Panel1.Controls.Add(Me.dtpSF)
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtSF_Primary)
        Me.SplitContainer6.Panel1.Controls.Add(Me.TxtSF_DocNo)
        Me.SplitContainer6.Panel1.Controls.Add(Me.Label3)
        '
        'SplitContainer6.Panel2
        '
        Me.SplitContainer6.Panel2.Controls.Add(Me.dtpAvailable)
        Me.SplitContainer6.Panel2.Controls.Add(Me.Label4)
        Me.SplitContainer6.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer6.Panel2.Controls.Add(Me.chkSF_LateOnly)
        Me.SplitContainer6.Panel2.Controls.Add(Me.TxtSF_Remarks)
        Me.SplitContainer6.Panel2.Controls.Add(Me.cboAvailability)
        Me.SplitContainer6.Size = New System.Drawing.Size(563, 68)
        Me.SplitContainer6.SplitterDistance = 270
        Me.SplitContainer6.SplitterWidth = 3
        Me.SplitContainer6.TabIndex = 57
        Me.SplitContainer6.TabStop = False
        '
        'TxtSF_StockInPurchaseNo
        '
        Me.TxtSF_StockInPurchaseNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSF_StockInPurchaseNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSF_StockInPurchaseNo.Location = New System.Drawing.Point(96, 48)
        Me.TxtSF_StockInPurchaseNo.Name = "TxtSF_StockInPurchaseNo"
        Me.TxtSF_StockInPurchaseNo.ReadOnly = True
        Me.TxtSF_StockInPurchaseNo.Size = New System.Drawing.Size(172, 20)
        Me.TxtSF_StockInPurchaseNo.TabIndex = 90
        '
        'LblSF_StockInPurchaseNo
        '
        Me.LblSF_StockInPurchaseNo.AutoSize = True
        Me.LblSF_StockInPurchaseNo.Location = New System.Drawing.Point(1, 50)
        Me.LblSF_StockInPurchaseNo.Name = "LblSF_StockInPurchaseNo"
        Me.LblSF_StockInPurchaseNo.Size = New System.Drawing.Size(95, 13)
        Me.LblSF_StockInPurchaseNo.TabIndex = 89
        Me.LblSF_StockInPurchaseNo.Text = "Stock In Purchase"
        '
        'llbSF_Calculate
        '
        Me.llbSF_Calculate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbSF_Calculate.AutoSize = True
        Me.llbSF_Calculate.Location = New System.Drawing.Point(216, 8)
        Me.llbSF_Calculate.Name = "llbSF_Calculate"
        Me.llbSF_Calculate.Size = New System.Drawing.Size(51, 13)
        Me.llbSF_Calculate.TabIndex = 88
        Me.llbSF_Calculate.TabStop = True
        Me.llbSF_Calculate.Text = "Calculate"
        Me.llbSF_Calculate.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'TxtSF_Urutan
        '
        Me.TxtSF_Urutan.Location = New System.Drawing.Point(70, 27)
        Me.TxtSF_Urutan.Name = "TxtSF_Urutan"
        Me.TxtSF_Urutan.Size = New System.Drawing.Size(22, 20)
        Me.TxtSF_Urutan.TabIndex = 73
        Me.TxtSF_Urutan.TabStop = False
        Me.TxtSF_Urutan.Text = "0"
        Me.TxtSF_Urutan.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(1, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 72
        Me.Label2.Text = "Date"
        '
        'dtpSF
        '
        Me.dtpSF.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpSF.CustomFormat = "dd MMM yyyy"
        Me.dtpSF.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpSF.Location = New System.Drawing.Point(96, 27)
        Me.dtpSF.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpSF.Name = "dtpSF"
        Me.dtpSF.Size = New System.Drawing.Size(172, 20)
        Me.dtpSF.TabIndex = 71
        '
        'TxtSF_Primary
        '
        Me.TxtSF_Primary.Location = New System.Drawing.Point(69, 4)
        Me.TxtSF_Primary.Name = "TxtSF_Primary"
        Me.TxtSF_Primary.Size = New System.Drawing.Size(22, 20)
        Me.TxtSF_Primary.TabIndex = 70
        Me.TxtSF_Primary.TabStop = False
        Me.TxtSF_Primary.Text = "-1"
        Me.TxtSF_Primary.Visible = False
        '
        'TxtSF_DocNo
        '
        Me.TxtSF_DocNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSF_DocNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSF_DocNo.Location = New System.Drawing.Point(96, 4)
        Me.TxtSF_DocNo.Name = "TxtSF_DocNo"
        Me.TxtSF_DocNo.Size = New System.Drawing.Size(118, 20)
        Me.TxtSF_DocNo.TabIndex = 52
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Document No"
        '
        'dtpAvailable
        '
        Me.dtpAvailable.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpAvailable.CustomFormat = "dd MMM yyyy"
        Me.dtpAvailable.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpAvailable.Location = New System.Drawing.Point(160, 25)
        Me.dtpAvailable.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpAvailable.Name = "dtpAvailable"
        Me.dtpAvailable.ShowCheckBox = True
        Me.dtpAvailable.Size = New System.Drawing.Size(126, 20)
        Me.dtpAvailable.TabIndex = 89
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 75
        Me.Label4.Text = "Filter"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(1, 6)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 73
        Me.Label5.Text = "Remarks"
        '
        'chkSF_LateOnly
        '
        Me.chkSF_LateOnly.AutoSize = True
        Me.chkSF_LateOnly.Location = New System.Drawing.Point(55, 49)
        Me.chkSF_LateOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.chkSF_LateOnly.Name = "chkSF_LateOnly"
        Me.chkSF_LateOnly.Size = New System.Drawing.Size(112, 17)
        Me.chkSF_LateOnly.TabIndex = 2
        Me.chkSF_LateOnly.Text = "Late Delivery Only"
        Me.chkSF_LateOnly.UseVisualStyleBackColor = True
        '
        'TxtSF_Remarks
        '
        Me.TxtSF_Remarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSF_Remarks.Location = New System.Drawing.Point(55, 4)
        Me.TxtSF_Remarks.Name = "TxtSF_Remarks"
        Me.TxtSF_Remarks.Size = New System.Drawing.Size(230, 20)
        Me.TxtSF_Remarks.TabIndex = 0
        '
        'cboAvailability
        '
        Me.cboAvailability.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAvailability.FormattingEnabled = True
        Me.cboAvailability.Items.AddRange(New Object() {"All", "Available Only", "Unavailable Only"})
        Me.cboAvailability.Location = New System.Drawing.Point(55, 24)
        Me.cboAvailability.Margin = New System.Windows.Forms.Padding(2)
        Me.cboAvailability.Name = "cboAvailability"
        Me.cboAvailability.Size = New System.Drawing.Size(102, 21)
        Me.cboAvailability.TabIndex = 74
        '
        'llbColumnSettings
        '
        Me.llbColumnSettings.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.llbColumnSettings.AutoSize = True
        Me.llbColumnSettings.Location = New System.Drawing.Point(91, 536)
        Me.llbColumnSettings.Name = "llbColumnSettings"
        Me.llbColumnSettings.Size = New System.Drawing.Size(83, 13)
        Me.llbColumnSettings.TabIndex = 89
        Me.llbColumnSettings.TabStop = True
        Me.llbColumnSettings.Text = "Column Settings"
        Me.llbColumnSettings.Visible = False
        Me.llbColumnSettings.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(10, 536)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(75, 13)
        Me.LinkLabel1.TabIndex = 43
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "<<< Hide Filter"
        Me.LinkLabel1.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(91, 536)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(76, 13)
        Me.LinkLabel2.TabIndex = 84
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Edit Parameter"
        Me.LinkLabel2.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerReportsProgress = True
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'LinkLabel3
        '
        Me.LinkLabel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(179, 536)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(83, 13)
        Me.LinkLabel3.TabIndex = 90
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Column Settings"
        Me.LinkLabel3.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'ColMP_No
        '
        Me.ColMP_No.Frozen = True
        Me.ColMP_No.HeaderText = "No"
        Me.ColMP_No.Name = "ColMP_No"
        Me.ColMP_No.Width = 46
        '
        'ColMP_CodeNo
        '
        Me.ColMP_CodeNo.Frozen = True
        Me.ColMP_CodeNo.HeaderText = "Code No"
        Me.ColMP_CodeNo.Name = "ColMP_CodeNo"
        Me.ColMP_CodeNo.Width = 74
        '
        'ColMP_Desc
        '
        Me.ColMP_Desc.HeaderText = "Description"
        Me.ColMP_Desc.Name = "ColMP_Desc"
        Me.ColMP_Desc.Width = 85
        '
        'ColMP_Unit
        '
        Me.ColMP_Unit.HeaderText = "Unit"
        Me.ColMP_Unit.Name = "ColMP_Unit"
        Me.ColMP_Unit.Width = 51
        '
        'ColMP_QtyOnHand
        '
        Me.ColMP_QtyOnHand.HeaderText = "Qty On Hand"
        Me.ColMP_QtyOnHand.Name = "ColMP_QtyOnHand"
        Me.ColMP_QtyOnHand.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOnHand.Visible = False
        Me.ColMP_QtyOnHand.Width = 94
        '
        'ColMP_QtyPOB
        '
        Me.ColMP_QtyPOB.HeaderText = "Qty Purchase Order Balance"
        Me.ColMP_QtyPOB.Name = "ColMP_QtyPOB"
        Me.ColMP_QtyPOB.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyPOB.Width = 167
        '
        'ColMP_QtyGIT
        '
        Me.ColMP_QtyGIT.HeaderText = "Qty Goods In Transit"
        Me.ColMP_QtyGIT.Name = "ColMP_QtyGIT"
        Me.ColMP_QtyGIT.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyGIT.Width = 129
        '
        'ColMP_QtyInquiry
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyInquiry.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColMP_QtyInquiry.HeaderText = "Qty Inquiry"
        Me.ColMP_QtyInquiry.Name = "ColMP_QtyInquiry"
        Me.ColMP_QtyInquiry.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyInquiry.Width = 82
        '
        'ColMP_QtyPendingDO
        '
        Me.ColMP_QtyPendingDO.HeaderText = "Qty Pending DO"
        Me.ColMP_QtyPendingDO.Name = "ColMP_QtyPendingDO"
        Me.ColMP_QtyPendingDO.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyPendingDO.Visible = False
        Me.ColMP_QtyPendingDO.Width = 109
        '
        'ColMP_QtySOB
        '
        Me.ColMP_QtySOB.HeaderText = "Qty Sales Order Balance"
        Me.ColMP_QtySOB.Name = "ColMP_QtySOB"
        Me.ColMP_QtySOB.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtySOB.Width = 148
        '
        'ColMP_QtyPrOB
        '
        Me.ColMP_QtyPrOB.HeaderText = "Qty PrOB"
        Me.ColMP_QtyPrOB.Name = "ColMP_QtyPrOB"
        Me.ColMP_QtyPrOB.Visible = False
        Me.ColMP_QtyPrOB.Width = 76
        '
        'ColMP_QtySafetyStock
        '
        Me.ColMP_QtySafetyStock.HeaderText = "Qty Safety Stock"
        Me.ColMP_QtySafetyStock.Name = "ColMP_QtySafetyStock"
        Me.ColMP_QtySafetyStock.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtySafetyStock.Width = 112
        '
        'ColMP_QtyLeadTime
        '
        Me.ColMP_QtyLeadTime.HeaderText = "Qty Lead Time"
        Me.ColMP_QtyLeadTime.Name = "ColMP_QtyLeadTime"
        Me.ColMP_QtyLeadTime.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyLeadTime.Width = 101
        '
        'ColMP_QtyRequired
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyRequired.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColMP_QtyRequired.HeaderText = "Qty Required"
        Me.ColMP_QtyRequired.Name = "ColMP_QtyRequired"
        Me.ColMP_QtyRequired.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyRequired.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_QtyRequired.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_QtyRequired.Width = 94
        '
        'ColMP_QtyOrderMaximum
        '
        Me.ColMP_QtyOrderMaximum.HeaderText = "Qty Order Maximum"
        Me.ColMP_QtyOrderMaximum.Name = "ColMP_QtyOrderMaximum"
        Me.ColMP_QtyOrderMaximum.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOrderMaximum.Width = 124
        '
        'ColMP_QtyOrderMinimum
        '
        Me.ColMP_QtyOrderMinimum.HeaderText = "Qty Order Minimum"
        Me.ColMP_QtyOrderMinimum.Name = "ColMP_QtyOrderMinimum"
        Me.ColMP_QtyOrderMinimum.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOrderMinimum.Width = 121
        '
        'ColMP_QtyOrderOptimum
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyOrderOptimum.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColMP_QtyOrderOptimum.HeaderText = "Qty Order Optimum"
        Me.ColMP_QtyOrderOptimum.Name = "ColMP_QtyOrderOptimum"
        Me.ColMP_QtyOrderOptimum.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_QtyOrderOptimum.Width = 121
        '
        'ColMP_QtyMax
        '
        Me.ColMP_QtyMax.HeaderText = "Qty Max"
        Me.ColMP_QtyMax.Name = "ColMP_QtyMax"
        Me.ColMP_QtyMax.Visible = False
        Me.ColMP_QtyMax.Width = 71
        '
        'ColMP_QtyMin
        '
        Me.ColMP_QtyMin.HeaderText = "Qty Min"
        Me.ColMP_QtyMin.Name = "ColMP_QtyMin"
        Me.ColMP_QtyMin.Visible = False
        Me.ColMP_QtyMin.Width = 68
        '
        'ColMP_QtyOpt
        '
        Me.ColMP_QtyOpt.HeaderText = "Qty Optimum"
        Me.ColMP_QtyOpt.Name = "ColMP_QtyOpt"
        Me.ColMP_QtyOpt.Visible = False
        Me.ColMP_QtyOpt.Width = 92
        '
        'ColMP_QtyOrder
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMP_QtyOrder.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColMP_QtyOrder.HeaderText = "Qty Order"
        Me.ColMP_QtyOrder.Name = "ColMP_QtyOrder"
        Me.ColMP_QtyOrder.Width = 77
        '
        'ColMP_OrderUnit
        '
        Me.ColMP_OrderUnit.HeaderText = "Unit"
        Me.ColMP_OrderUnit.Name = "ColMP_OrderUnit"
        Me.ColMP_OrderUnit.Visible = False
        Me.ColMP_OrderUnit.Width = 51
        '
        'ColMP_OrderPoint
        '
        Me.ColMP_OrderPoint.HeaderText = "Order Point Date"
        Me.ColMP_OrderPoint.Name = "ColMP_OrderPoint"
        Me.ColMP_OrderPoint.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_OrderPoint.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_OrderPoint.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_OrderPoint.Width = 111
        '
        'ColMP_NextOrderPoint
        '
        Me.ColMP_NextOrderPoint.HeaderText = "Next Order Point Date"
        Me.ColMP_NextOrderPoint.Name = "ColMP_NextOrderPoint"
        Me.ColMP_NextOrderPoint.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColMP_NextOrderPoint.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_NextOrderPoint.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_NextOrderPoint.Width = 136
        '
        'ColMP_ProcurementType
        '
        Me.ColMP_ProcurementType.HeaderText = "Procurement"
        Me.ColMP_ProcurementType.Items.AddRange(New Object() {"Produce", "Buy"})
        Me.ColMP_ProcurementType.Name = "ColMP_ProcurementType"
        Me.ColMP_ProcurementType.Visible = False
        Me.ColMP_ProcurementType.Width = 73
        '
        'ColMP_ProcurementStrategy
        '
        Me.ColMP_ProcurementStrategy.HeaderText = "Procurement Strategy"
        Me.ColMP_ProcurementStrategy.Items.AddRange(New Object() {"Build To Order", "Build To Stock"})
        Me.ColMP_ProcurementStrategy.Name = "ColMP_ProcurementStrategy"
        Me.ColMP_ProcurementStrategy.Width = 115
        '
        'ColMP_OrderStrategy
        '
        Me.ColMP_OrderStrategy.HeaderText = "Order Strategy"
        Me.ColMP_OrderStrategy.Items.AddRange(New Object() {"Order To Optimum Qty", "Order By Optimum Order Qty"})
        Me.ColMP_OrderStrategy.Name = "ColMP_OrderStrategy"
        Me.ColMP_OrderStrategy.Width = 81
        '
        'ColMP_FollowUp
        '
        Me.ColMP_FollowUp.HeaderText = "Follow Up"
        Me.ColMP_FollowUp.Name = "ColMP_FollowUp"
        Me.ColMP_FollowUp.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColMP_FollowUp.Width = 79
        '
        'ColMP_PreferedSupplier
        '
        Me.ColMP_PreferedSupplier.HeaderText = "Prefered Supplier"
        Me.ColMP_PreferedSupplier.Name = "ColMP_PreferedSupplier"
        Me.ColMP_PreferedSupplier.Width = 113
        '
        'ColMP_Currency
        '
        Me.ColMP_Currency.HeaderText = "Currency"
        Me.ColMP_Currency.Name = "ColMP_Currency"
        Me.ColMP_Currency.Width = 74
        '
        'ColMP_UnitPrice
        '
        Me.ColMP_UnitPrice.HeaderText = "Unit Price"
        Me.ColMP_UnitPrice.Name = "ColMP_UnitPrice"
        Me.ColMP_UnitPrice.Width = 78
        '
        'ColMP_Disc
        '
        Me.ColMP_Disc.HeaderText = "Discount"
        Me.ColMP_Disc.Name = "ColMP_Disc"
        Me.ColMP_Disc.Width = 74
        '
        'ColMP_NettPrice
        '
        Me.ColMP_NettPrice.HeaderText = "Nett Price"
        Me.ColMP_NettPrice.Name = "ColMP_NettPrice"
        Me.ColMP_NettPrice.Width = 79
        '
        'ColMP_SubTotal
        '
        Me.ColMP_SubTotal.HeaderText = "Sub Total"
        Me.ColMP_SubTotal.Name = "ColMP_SubTotal"
        Me.ColMP_SubTotal.Width = 78
        '
        'ColMP_PrimaryCurrency
        '
        Me.ColMP_PrimaryCurrency.HeaderText = "Primary Curr"
        Me.ColMP_PrimaryCurrency.Name = "ColMP_PrimaryCurrency"
        Me.ColMP_PrimaryCurrency.Visible = False
        Me.ColMP_PrimaryCurrency.Width = 88
        '
        'ColMP_PrimarySupplier
        '
        Me.ColMP_PrimarySupplier.HeaderText = "Primary Supplier"
        Me.ColMP_PrimarySupplier.Name = "ColMP_PrimarySupplier"
        Me.ColMP_PrimarySupplier.Visible = False
        Me.ColMP_PrimarySupplier.Width = 107
        '
        'ColMP_PrimaryCodeNo
        '
        Me.ColMP_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColMP_PrimaryCodeNo.Name = "ColMP_PrimaryCodeNo"
        Me.ColMP_PrimaryCodeNo.Visible = False
        Me.ColMP_PrimaryCodeNo.Width = 111
        '
        'ColMP_DefPrecCurrency
        '
        Me.ColMP_DefPrecCurrency.HeaderText = "Def Prec Currency"
        Me.ColMP_DefPrecCurrency.Name = "ColMP_DefPrecCurrency"
        Me.ColMP_DefPrecCurrency.Visible = False
        Me.ColMP_DefPrecCurrency.Width = 119
        '
        'ColMP_DefPrecUnit
        '
        Me.ColMP_DefPrecUnit.HeaderText = "DefPrecProduct"
        Me.ColMP_DefPrecUnit.Name = "ColMP_DefPrecUnit"
        Me.ColMP_DefPrecUnit.Visible = False
        Me.ColMP_DefPrecUnit.Width = 108
        '
        'ColMP_PrimaryOrderUnit
        '
        Me.ColMP_PrimaryOrderUnit.HeaderText = "Primary Order Unit"
        Me.ColMP_PrimaryOrderUnit.Name = "ColMP_PrimaryOrderUnit"
        Me.ColMP_PrimaryOrderUnit.Visible = False
        Me.ColMP_PrimaryOrderUnit.Width = 117
        '
        'ColMP_DefPrecOrderUnit
        '
        Me.ColMP_DefPrecOrderUnit.HeaderText = "DefPrec Order Unit"
        Me.ColMP_DefPrecOrderUnit.Name = "ColMP_DefPrecOrderUnit"
        Me.ColMP_DefPrecOrderUnit.Visible = False
        Me.ColMP_DefPrecOrderUnit.Width = 122
        '
        'ColMP_OrderPointDelay
        '
        Me.ColMP_OrderPointDelay.HeaderText = "Order Point Detail"
        Me.ColMP_OrderPointDelay.Name = "ColMP_OrderPointDelay"
        Me.ColMP_OrderPointDelay.Visible = False
        Me.ColMP_OrderPointDelay.Width = 115
        '
        'ColMP_NextOrderPointDelay
        '
        Me.ColMP_NextOrderPointDelay.HeaderText = "Next Order Point Delay"
        Me.ColMP_NextOrderPointDelay.Name = "ColMP_NextOrderPointDelay"
        Me.ColMP_NextOrderPointDelay.Visible = False
        Me.ColMP_NextOrderPointDelay.Width = 140
        '
        'ColMP_LeadTimeDays
        '
        Me.ColMP_LeadTimeDays.HeaderText = "Lead Time (days)"
        Me.ColMP_LeadTimeDays.Name = "ColMP_LeadTimeDays"
        Me.ColMP_LeadTimeDays.Visible = False
        Me.ColMP_LeadTimeDays.Width = 113
        '
        'ColMP_QtyOrderMulti
        '
        Me.ColMP_QtyOrderMulti.HeaderText = "QtyOrderMulti"
        Me.ColMP_QtyOrderMulti.Name = "ColMP_QtyOrderMulti"
        Me.ColMP_QtyOrderMulti.Visible = False
        Me.ColMP_QtyOrderMulti.Width = 96
        '
        'ColMP_PrimaryUnitMulti
        '
        Me.ColMP_PrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColMP_PrimaryUnitMulti.Name = "ColMP_PrimaryUnitMulti"
        Me.ColMP_PrimaryUnitMulti.Visible = False
        Me.ColMP_PrimaryUnitMulti.Width = 113
        '
        'ColMP_DefPrecUnitMulti
        '
        Me.ColMP_DefPrecUnitMulti.HeaderText = "Def Prec Unit Multi"
        Me.ColMP_DefPrecUnitMulti.Name = "ColMP_DefPrecUnitMulti"
        Me.ColMP_DefPrecUnitMulti.Visible = False
        Me.ColMP_DefPrecUnitMulti.Width = 121
        '
        'ColMP_RatioMulti
        '
        Me.ColMP_RatioMulti.HeaderText = "Ratio Multi"
        Me.ColMP_RatioMulti.Name = "ColMP_RatioMulti"
        Me.ColMP_RatioMulti.Visible = False
        Me.ColMP_RatioMulti.Width = 82
        '
        'ColMP_RatioDefault
        '
        Me.ColMP_RatioDefault.HeaderText = "Ratio Default"
        Me.ColMP_RatioDefault.Name = "ColMP_RatioDefault"
        Me.ColMP_RatioDefault.Visible = False
        Me.ColMP_RatioDefault.Width = 94
        '
        'ColMP_DiscAmount
        '
        Me.ColMP_DiscAmount.HeaderText = "Disc Amount"
        Me.ColMP_DiscAmount.Name = "ColMP_DiscAmount"
        Me.ColMP_DiscAmount.Visible = False
        Me.ColMP_DiscAmount.Width = 92
        '
        'ColMP_DiscFromPerct
        '
        Me.ColMP_DiscFromPerct.HeaderText = "DiscFromPerct"
        Me.ColMP_DiscFromPerct.Name = "ColMP_DiscFromPerct"
        Me.ColMP_DiscFromPerct.Visible = False
        Me.ColMP_DiscFromPerct.Width = 101
        '
        'ColMP_DetailFollowUp
        '
        Me.ColMP_DetailFollowUp.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColMP_DetailFollowUp.HeaderText = "Detail Follow Up"
        Me.ColMP_DetailFollowUp.Name = "ColMP_DetailFollowUp"
        Me.ColMP_DetailFollowUp.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColMP_DetailFollowUp.Width = 90
        '
        'FrmTrans_MaterialPlanningDashboard
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(843, 553)
        Me.Controls.Add(Me.llbColumnSettings)
        Me.Controls.Add(Me.LinkLabel3)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.btnClose)
        Me.MinimumSize = New System.Drawing.Size(859, 591)
        Me.Name = "FrmTrans_MaterialPlanningDashboard"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Planning Dashboard"
        CType(Me.dgvMP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GboFilter_PurcReqDate.ResumeLayout(False)
        Me.GboFilter_PurcReqDate.PerformLayout()
        Me.GboFilter_MODate.ResumeLayout(False)
        Me.GboFilter_MODate.PerformLayout()
        Me.GboFilter_SODate.ResumeLayout(False)
        Me.GboFilter_SODate.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.StatusStrip4.ResumeLayout(False)
        Me.StatusStrip4.PerformLayout()
        Me.gboMP.ResumeLayout(False)
        Me.gboMP.PerformLayout()
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel1.PerformLayout()
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        Me.SplitContainer5.Panel2.PerformLayout()
        Me.SplitContainer5.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.TabPage5.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        Me.SplitContainer4.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.StatusStrip3.ResumeLayout(False)
        Me.StatusStrip3.PerformLayout()
        Me.gboSF.ResumeLayout(False)
        Me.gboSF.PerformLayout()
        CType(Me.dgvSF, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer6.Panel1.ResumeLayout(False)
        Me.SplitContainer6.Panel1.PerformLayout()
        Me.SplitContainer6.Panel2.ResumeLayout(False)
        Me.SplitContainer6.Panel2.PerformLayout()
        Me.SplitContainer6.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvMP As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cboFilter_ProductGroupLevel As System.Windows.Forms.ComboBox
    Friend WithEvents LblFilter_ProductGroup As System.Windows.Forms.Label
    Friend WithEvents btnFilter_FindPreferedSupplier As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_PreferedSupplier As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_PreferedSupplier As System.Windows.Forms.Label
    Friend WithEvents chkMP_RequiredOnly As System.Windows.Forms.CheckBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnMP_Print As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents btnMP_Find As System.Windows.Forms.Button
    Friend WithEvents btnMP_OK As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents btnFilter_FindUser As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_UserCreateSO As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_UserCreateSO As System.Windows.Forms.Label
    Friend WithEvents btnFilter_FindCust As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_CustCode As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_Cust As System.Windows.Forms.Label
    Friend WithEvents dgvSF As System.Windows.Forms.DataGridView
    Friend WithEvents chkSF_LateOnly As System.Windows.Forms.CheckBox
    Friend WithEvents StatusStrip3 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslSF_CreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslSF_CreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslSF_ModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslSF_ModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtFilter_PreferedSupplierDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtFilter_CustName As System.Windows.Forms.TextBox
    Friend WithEvents btnMP_Delete As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_Desc As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtFilter_CodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents TxtMP_QtyRequired As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtySafetyStock As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyOptimumStock As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyOnHand As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPendingDO As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtySOB As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyInquiry As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPOB As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyGIT As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtMP_PreferedSupplier As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_NettPrice As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_CurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_Discount As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_UnitPrice As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip4 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslMP_CreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents LblMP_Unit1 As System.Windows.Forms.Label
    Friend WithEvents LblRatio As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyRequired2 As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_PrimaryOrderUnit As System.Windows.Forms.TextBox
    Friend WithEvents cboMP_OrderUnit As System.Windows.Forms.ComboBox
    Friend WithEvents TxtMP_SubTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyOrder As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_PrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents btnMP_FindSupplier As System.Windows.Forms.Button
    Friend WithEvents TxtMP_PreferedSupplierDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_PrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents btnMP_FindCurrency As System.Windows.Forms.Button
    Friend WithEvents llbMP_Ratio As System.Windows.Forms.LinkLabel
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cboMPOrderStrategy As System.Windows.Forms.ComboBox
    Friend WithEvents LlbMP_UnitPrice As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtFilter_ProductGroupLevelDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFilter_FindProductGroupLevel As System.Windows.Forms.Button
    Friend WithEvents TxtFilter_ProductGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_PoductGroupDesc As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents gboMP As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer5 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents dtpMP As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtMP_Primary As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_DocNo As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_Remarks As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents btnMP_Edit As System.Windows.Forms.Button
    Friend WithEvents btnMP_New As System.Windows.Forms.Button
    Friend WithEvents btnMP_Cancel As System.Windows.Forms.Button
    Friend WithEvents llbMP_Calculate As System.Windows.Forms.LinkLabel
    Friend WithEvents llbMP_CreatePO As System.Windows.Forms.LinkLabel
    Friend WithEvents llbMP_Filter As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents dtpMP_RefDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtMP_DaysPerMonth As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dtpMP_Filter As System.Windows.Forms.DateTimePicker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtMP_AvgPerMonth As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents bwRefDate As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtFilter_PrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents TxtFilter_PrimaryProductGroup As System.Windows.Forms.TextBox
    Friend WithEvents TxtFilter_PrimaryPreferedSupplier As System.Windows.Forms.TextBox
    Friend WithEvents llbMP_AddToPO As System.Windows.Forms.LinkLabel
    Friend WithEvents llbMP_CekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents tslMP_CreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslMP_ModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslMP_ModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents llbSF_CreateDO As System.Windows.Forms.LinkLabel
    Friend WithEvents btnSF_Edit As System.Windows.Forms.Button
    Friend WithEvents btnSF_New As System.Windows.Forms.Button
    Friend WithEvents btnSF_Find As System.Windows.Forms.Button
    Friend WithEvents btnSF_Print As System.Windows.Forms.Button
    Friend WithEvents btnSF_Delete As System.Windows.Forms.Button
    Friend WithEvents SplitContainer6 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtSF_Urutan As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpSF As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtSF_Primary As System.Windows.Forms.TextBox
    Friend WithEvents TxtSF_DocNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtSF_Remarks As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents gboSF As System.Windows.Forms.GroupBox
    Friend WithEvents btnSF_Cancel As System.Windows.Forms.Button
    Friend WithEvents btnSF_OK As System.Windows.Forms.Button
    Friend WithEvents llbSF_CekAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbSF_Calculate As System.Windows.Forms.LinkLabel
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents llbColumnSettings As System.Windows.Forms.LinkLabel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cboAvailability As System.Windows.Forms.ComboBox
    Friend WithEvents dtpAvailable As System.Windows.Forms.DateTimePicker
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtFilter_NamaUserCreateSO As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents GboFilter_SODate As System.Windows.Forms.GroupBox
    Friend WithEvents DtpFilter_To As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents DtpFilter_From As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents CboFilter_ProcurementStrategy As System.Windows.Forms.ComboBox
    Friend WithEvents LblFilter_ProcurementStrategy As System.Windows.Forms.Label
    Friend WithEvents chkSODateFilter As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker6 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPrOB As System.Windows.Forms.TextBox
    Friend WithEvents chkFilter_Service As System.Windows.Forms.CheckBox
    Friend WithEvents chkFilter_Goods As System.Windows.Forms.CheckBox
    Friend WithEvents chkFilter_PrOBItem As System.Windows.Forms.CheckBox
    Friend WithEvents chkFilter_SOBItem As System.Windows.Forms.CheckBox
    Friend WithEvents chkFilter_FollowUpOnly As System.Windows.Forms.CheckBox
    Friend WithEvents TxtMP_DiscAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_DiscFromPerct As System.Windows.Forms.TextBox
    Friend WithEvents llbMP_Tolerance As System.Windows.Forms.LinkLabel
    Friend WithEvents btnFilter_FindRefParty As System.Windows.Forms.Button
    Friend WithEvents Txtfilter_PrimaryRefParty As System.Windows.Forms.TextBox
    Friend WithEvents Txtfilter_RefPartyName As System.Windows.Forms.TextBox
    Friend WithEvents Txtfilter_RefPartyCode As System.Windows.Forms.TextBox
    Friend WithEvents LblRefParty As System.Windows.Forms.Label
    Friend WithEvents GboFilter_PurcReqDate As System.Windows.Forms.GroupBox
    Friend WithEvents chkFilter_PurhReqDate As System.Windows.Forms.CheckBox
    Friend WithEvents DtpFilter_PurchReq_To As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents DtpFilter_PurcReq_From As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtFilter_PrimaryPurchReq As System.Windows.Forms.TextBox
    Friend WithEvents TxtFilter_PurcReqNo As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_PurcReq As System.Windows.Forms.Label
    Friend WithEvents btnFilter_FindPurchReq As System.Windows.Forms.Button
    Friend WithEvents llbMP_CopyMP As System.Windows.Forms.LinkLabel
    Friend WithEvents ColSF_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_Type As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_SONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_CustomerPONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_SODate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_Customer As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_QtySOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_QtyAvailable As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_PromiseToDeliver As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_PromiseToDeliver_Str As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_EstimateDelivery As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_EstimateDelivery_Str As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColSF_CreateDO As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColSF_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_DeliveryStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_PrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_PrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_PrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_QtyReadyStock As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_TypeInt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSF_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtSF_StockInPurchaseNo As System.Windows.Forms.TextBox
    Friend WithEvents LblSF_StockInPurchaseNo As System.Windows.Forms.Label
    Friend WithEvents btnFilter_FindCodeNo As System.Windows.Forms.Button
    Friend WithEvents chkFilter_PrOBLargerThanPOB As System.Windows.Forms.CheckBox
    Friend WithEvents llmbMP_SetQtyOrderEqualQtyPrOB As System.Windows.Forms.LinkLabel
    Friend WithEvents llbMP_LocationSetting As System.Windows.Forms.LinkLabel
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyRMReq As System.Windows.Forms.TextBox
    Friend WithEvents TxtMP_QtyMOB As System.Windows.Forms.TextBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents GboFilter_MODate As System.Windows.Forms.GroupBox
    Friend WithEvents chkMODateFilter As System.Windows.Forms.CheckBox
    Friend WithEvents DtpFilter_MO_To As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents DtpFilter_MO_From As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents TxtFilter_PrimaryMO As System.Windows.Forms.TextBox
    Friend WithEvents TxtFilter_MONo As System.Windows.Forms.TextBox
    Friend WithEvents LblFilter_FindMO As System.Windows.Forms.Label
    Friend WithEvents btnFilter_FindMO As System.Windows.Forms.Button
    Friend WithEvents chkClose As System.Windows.Forms.CheckBox
    Friend WithEvents TxtMP_QtyFree As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents llbAuthDetail As System.Windows.Forms.LinkLabel
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyRMReqSpare As System.Windows.Forms.TextBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtySOBSpare As System.Windows.Forms.TextBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents TxtMP_QtyPrOBSpare As System.Windows.Forms.TextBox
    Friend WithEvents ColMP_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOnHand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyPOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyGIT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyInquiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyPendingDO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtySOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyPrOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtySafetyStock As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyLeadTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyRequired As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColMP_QtyOrderMaximum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrderMinimum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrderOptimum As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyMax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyMin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOpt As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_OrderUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_OrderPoint As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColMP_NextOrderPoint As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColMP_ProcurementType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMP_ProcurementStrategy As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMP_OrderStrategy As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColMP_FollowUp As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColMP_PreferedSupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Currency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_UnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_Disc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_NettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_SubTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimarySupplier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryOrderUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecOrderUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_OrderPointDelay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_NextOrderPointDelay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_LeadTimeDays As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_QtyOrderMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_PrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DefPrecUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_RatioMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_RatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMP_DetailFollowUp As System.Windows.Forms.DataGridViewLinkColumn
End Class
