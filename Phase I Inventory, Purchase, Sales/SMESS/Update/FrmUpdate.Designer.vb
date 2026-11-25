<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUpdate
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.ProgPOB = New System.Windows.Forms.ProgressBar
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ProgSOB = New System.Windows.Forms.ProgressBar
        Me.Button2 = New System.Windows.Forms.Button
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.Button3 = New System.Windows.Forms.Button
        Me.ProgUpdateDO1 = New System.Windows.Forms.ProgressBar
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.Prog_SeriFakturPajak = New System.Windows.Forms.ProgressBar
        Me.Button4 = New System.Windows.Forms.Button
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.Prog_UpdateNoPO = New System.Windows.Forms.ProgressBar
        Me.Button5 = New System.Windows.Forms.Button
        Me.Prog_UpdateNoSO = New System.Windows.Forms.ProgressBar
        Me.Button6 = New System.Windows.Forms.Button
        Me.ProgressBar3 = New System.Windows.Forms.ProgressBar
        Me.Button7 = New System.Windows.Forms.Button
        Me.ProgressBar4 = New System.Windows.Forms.ProgressBar
        Me.Button8 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.prgCost = New System.Windows.Forms.ProgressBar
        Me.Button9 = New System.Windows.Forms.Button
        Me.BackgroundWorker5 = New System.ComponentModel.BackgroundWorker
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Button10 = New System.Windows.Forms.Button
        Me.prgItemPR = New System.Windows.Forms.ProgressBar
        Me.btnItemPR = New System.Windows.Forms.Button
        Me.BackgroundWorker6 = New System.ComponentModel.BackgroundWorker
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dtpBalikJurnal = New System.Windows.Forms.DateTimePicker
        Me.lblProgressBalikJurnal = New System.Windows.Forms.Label
        Me.prgBalikJurnal = New System.Windows.Forms.ProgressBar
        Me.btnBalikJurnal = New System.Windows.Forms.Button
        Me.Label62 = New System.Windows.Forms.Label
        Me.LblProgressARPay = New System.Windows.Forms.Label
        Me.dtpARPay = New System.Windows.Forms.DateTimePicker
        Me.prgBayarAR = New System.Windows.Forms.ProgressBar
        Me.btnBayarAR = New System.Windows.Forms.Button
        Me.Label64 = New System.Windows.Forms.Label
        Me.prgRefClosing = New System.Windows.Forms.ProgressBar
        Me.dtpRefClosing = New System.Windows.Forms.DateTimePicker
        Me.btnDeleteClosing = New System.Windows.Forms.Button
        Me.Label59 = New System.Windows.Forms.Label
        Me.prgRefMP = New System.Windows.Forms.ProgressBar
        Me.btnCekUnbalanceAccount = New System.Windows.Forms.Button
        Me.Label56 = New System.Windows.Forms.Label
        Me.dtpRefMP = New System.Windows.Forms.DateTimePicker
        Me.btnUpdatePrimaryAkun = New System.Windows.Forms.Button
        Me.Label54 = New System.Windows.Forms.Label
        Me.btnCekSalaryKomponen = New System.Windows.Forms.Button
        Me.Label53 = New System.Windows.Forms.Label
        Me.prgUpdatePenghasilanBruto = New System.Windows.Forms.ProgressBar
        Me.btnUpdatePenghasilanBruto = New System.Windows.Forms.Button
        Me.Label52 = New System.Windows.Forms.Label
        Me.btnDeleteMPSF = New System.Windows.Forms.Button
        Me.Label48 = New System.Windows.Forms.Label
        Me.TxtAccountDesc_COGS = New System.Windows.Forms.TextBox
        Me.TextBox15 = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccount_COGS = New System.Windows.Forms.TextBox
        Me.btnUpdateAccount_COGS = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.TxtAccount_COGS = New System.Windows.Forms.TextBox
        Me.btnFindAccount_COGS = New System.Windows.Forms.Button
        Me.TxtAccountDesc_SalesReturn = New System.Windows.Forms.TextBox
        Me.TextBox11 = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccount_SalesReturn = New System.Windows.Forms.TextBox
        Me.btnUpdateAccount_SalesReturn = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtAccount_SalesReturn = New System.Windows.Forms.TextBox
        Me.btnFindAccount_SalesReturn = New System.Windows.Forms.Button
        Me.TxtAccountDesc_Sales = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccount_Sales = New System.Windows.Forms.TextBox
        Me.btnUpdateAccount_Sales = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtAccount_Sales = New System.Windows.Forms.TextBox
        Me.btnFindAccount_Sales = New System.Windows.Forms.Button
        Me.TxtAccountDesc_Inventory = New System.Windows.Forms.TextBox
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccount_Inventory = New System.Windows.Forms.TextBox
        Me.btnUpdateAccount_Inventory = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtAccount_Inventory = New System.Windows.Forms.TextBox
        Me.btnFindAccount_Inventory = New System.Windows.Forms.Button
        Me.TxtAccountDesc_Adjustment = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccount_Adjustment = New System.Windows.Forms.TextBox
        Me.btnUpdateAccount_Adjustment = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtAccount_Adjustment = New System.Windows.Forms.TextBox
        Me.btnFindAccount_Adjustment = New System.Windows.Forms.Button
        Me.TxtAccountDesc_StockTaking = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.TxtPrimaryAccount_StockTaking = New System.Windows.Forms.TextBox
        Me.btnUpdateAccount_StockTaking = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtAccount_StockTaking = New System.Windows.Forms.TextBox
        Me.btnFindAccount_StockTaking = New System.Windows.Forms.Button
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Prg_UpdateDiscAmount = New System.Windows.Forms.ProgressBar
        Me.btnUpdateDiscAmount = New System.Windows.Forms.Button
        Me.LblUpdateDiscAmount = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.ProgressBar5 = New System.Windows.Forms.ProgressBar
        Me.btnUpdataHR = New System.Windows.Forms.Button
        Me.lblUpdatev160_HR = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.btnUpdatev160_TBLMasterProduct = New System.Windows.Forms.Button
        Me.LblUpdatev160_TBLMasterProduct = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.prgUpdatev160_TambahTabelBaru = New System.Windows.Forms.ProgressBar
        Me.btnUpdatev160_TambahTabelBaru = New System.Windows.Forms.Button
        Me.lblUpdatev160_TambahTabelBaru = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.prgIdentifierPO = New System.Windows.Forms.ProgressBar
        Me.btnUpdateIdentifierPO = New System.Windows.Forms.Button
        Me.lblProgressPO = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.prgIdentifierSO = New System.Windows.Forms.ProgressBar
        Me.btnUpdate_IdentifierSO = New System.Windows.Forms.Button
        Me.LblIdentifierSO = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.prgUpdateSalesPrice = New System.Windows.Forms.ProgressBar
        Me.btnPindahkanSalesPrice = New System.Windows.Forms.Button
        Me.LblProgres_PindahSalesPrice = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.btnUpdatev160_UpdateNumberingMP = New System.Windows.Forms.Button
        Me.LblProgress_UpdateNumberingMaterialPlanning = New System.Windows.Forms.Label
        Me.LblProgress_UpdateNumbering = New System.Windows.Forms.Label
        Me.btnUpdatev160_UpdateNumberingCL = New System.Windows.Forms.Button
        Me.LblProgress_UpdateNumberingColletionLetter = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnUpdatev160_SC = New System.Windows.Forms.Button
        Me.lblPrgStockCard = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.prgStockCard = New System.Windows.Forms.ProgressBar
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.btnUpdatePresentasiPayroll = New System.Windows.Forms.Button
        Me.Label83 = New System.Windows.Forms.Label
        Me.Label84 = New System.Windows.Forms.Label
        Me.ProgressBar18 = New System.Windows.Forms.ProgressBar
        Me.btnUpdateSalaryOverride = New System.Windows.Forms.Button
        Me.Label85 = New System.Windows.Forms.Label
        Me.Label86 = New System.Windows.Forms.Label
        Me.ProgressBar19 = New System.Windows.Forms.ProgressBar
        Me.btnUpdateGetVersion = New System.Windows.Forms.Button
        Me.btnUpdateSetVersion = New System.Windows.Forms.Button
        Me.Label82 = New System.Windows.Forms.Label
        Me.Button21 = New System.Windows.Forms.Button
        Me.Label78 = New System.Windows.Forms.Label
        Me.Label79 = New System.Windows.Forms.Label
        Me.ProgressBar16 = New System.Windows.Forms.ProgressBar
        Me.Button20 = New System.Windows.Forms.Button
        Me.Label73 = New System.Windows.Forms.Label
        Me.Label74 = New System.Windows.Forms.Label
        Me.ProgressBar15 = New System.Windows.Forms.ProgressBar
        Me.btnUpdate_ChkInv = New System.Windows.Forms.Button
        Me.lblProgress_ChkInv = New System.Windows.Forms.Label
        Me.Label71 = New System.Windows.Forms.Label
        Me.PrgUpdate_ChkInv = New System.Windows.Forms.ProgressBar
        Me.btnUpdate_UpdatePrOB = New System.Windows.Forms.Button
        Me.lblProgress_UpdatePrOB = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.PrgUpdate_PrOB = New System.Windows.Forms.ProgressBar
        Me.btnUpdate_UpdatePOB = New System.Windows.Forms.Button
        Me.lblProgress_UpdatePOB = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.PrgUpdate_UpdatePOB = New System.Windows.Forms.ProgressBar
        Me.btnUpdate_NPWP_AOP = New System.Windows.Forms.Button
        Me.lblProgress_NPWP_AOP = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.PrgUpdate_NPWP_AOP = New System.Windows.Forms.ProgressBar
        Me.btnUpdateEmail = New System.Windows.Forms.Button
        Me.LblUpdateEmail = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.prgUpdateEmail = New System.Windows.Forms.ProgressBar
        Me.btnDeletePackingList = New System.Windows.Forms.Button
        Me.lblPackingList = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.prgPackingList = New System.Windows.Forms.ProgressBar
        Me.btnUpdateSalesPerson = New System.Windows.Forms.Button
        Me.LblUpdateSalesPerson = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.prbUpdateSalesPerson = New System.Windows.Forms.ProgressBar
        Me.TabPage5 = New System.Windows.Forms.TabPage
        Me.prgClearTable = New System.Windows.Forms.ProgressBar
        Me.LblClearTable_Progress = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.llbCheckAll = New System.Windows.Forms.LinkLabel
        Me.llbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.TxtCT_NamaScheme = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.TxtCT_DBName = New System.Windows.Forms.TextBox
        Me.chkCT_New = New System.Windows.Forms.CheckBox
        Me.btnCT_LoadTableName = New System.Windows.Forms.Button
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtCT_ConnectionString = New System.Windows.Forms.TextBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCT_CreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCT_CreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCT_ModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCT_ModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.TxtCT_Primary = New System.Windows.Forms.TextBox
        Me.btnCT_Delete = New System.Windows.Forms.Button
        Me.btnCT_Save = New System.Windows.Forms.Button
        Me.btnCT_Clear = New System.Windows.Forms.Button
        Me.cboCT_SchemeName = New System.Windows.Forms.ComboBox
        Me.dgvCT_DetailTable = New System.Windows.Forms.DataGridView
        Me.ColCT_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_TableName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_RowCount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCT_Clear = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.TabPage6 = New System.Windows.Forms.TabPage
        Me.Label69 = New System.Windows.Forms.Label
        Me.Label68 = New System.Windows.Forms.Label
        Me.TxtLocControl_LocationDesc = New System.Windows.Forms.TextBox
        Me.TxtLocControl_PrimaryLocation = New System.Windows.Forms.TextBox
        Me.BtnLocControl_FindLocation = New System.Windows.Forms.Button
        Me.TxtLocControl_LocationCode = New System.Windows.Forms.TextBox
        Me.TxtLocControl_Desc = New System.Windows.Forms.TextBox
        Me.TxtLocControl_TxtPrimaryProduct = New System.Windows.Forms.TextBox
        Me.BtnLocControl_FindCodeNo = New System.Windows.Forms.Button
        Me.TxtLocControl_CodeNo = New System.Windows.Forms.TextBox
        Me.ChkLocation_tbltrans_stock_taking_tag = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltransaksi_stock_card = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_stock_taking_itemstart = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_stock_taking_itembuffer = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_salesreturn_item_import = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_salesreturn_item_copy = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_salesreturn_item = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_salesinvoice_item_import = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_salesinvoice_item_copy = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_salesinvoice_item = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_purchasereturn_item = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_booking_invitem = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltrans_aop_item = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltransaksi_stocktransfer_item = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltransaksi_stockbundling_item_before = New System.Windows.Forms.CheckBox
        Me.chkLocation_tbltransaksi_stockbundling_item_after = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tbltransaksi_stockinout_item = New System.Windows.Forms.CheckBox
        Me.ChkLocation_tblmaster_productqty = New System.Windows.Forms.CheckBox
        Me.prgLocationCont = New System.Windows.Forms.ProgressBar
        Me.btnLocation_Start = New System.Windows.Forms.Button
        Me.chkLocation_tblmaster_product = New System.Windows.Forms.CheckBox
        Me.TabPage7 = New System.Windows.Forms.TabPage
        Me.btn_v170_quotation = New System.Windows.Forms.Button
        Me.lbl_v170_quotation = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.prg_v170_quotation = New System.Windows.Forms.ProgressBar
        Me.TabPage8 = New System.Windows.Forms.TabPage
        Me.btn_v180_Subtotal = New System.Windows.Forms.Button
        Me.lbl_v180_Subtotal = New System.Windows.Forms.Label
        Me.Label42 = New System.Windows.Forms.Label
        Me.prg_v180_Subtotal = New System.Windows.Forms.ProgressBar
        Me.Button13 = New System.Windows.Forms.Button
        Me.Label37 = New System.Windows.Forms.Label
        Me.Label41 = New System.Windows.Forms.Label
        Me.ProgressBar6 = New System.Windows.Forms.ProgressBar
        Me.btn_v180_DetailItemJurnal = New System.Windows.Forms.Button
        Me.lbl_v180_DetailItemJurnal = New System.Windows.Forms.Label
        Me.Label40 = New System.Windows.Forms.Label
        Me.prg_v180_DetailItemJurnal = New System.Windows.Forms.ProgressBar
        Me.btn_v180_ProsesBundling = New System.Windows.Forms.Button
        Me.lbl_v180_ProsesBundling = New System.Windows.Forms.Label
        Me.Label39 = New System.Windows.Forms.Label
        Me.prg_v180_ProsesBundling = New System.Windows.Forms.ProgressBar
        Me.btn_v180_DetailClosing = New System.Windows.Forms.Button
        Me.lbl_v180_DetailClosing = New System.Windows.Forms.Label
        Me.Label38 = New System.Windows.Forms.Label
        Me.prg_v180_DetailClosing = New System.Windows.Forms.ProgressBar
        Me.btn_v180_Salary = New System.Windows.Forms.Button
        Me.lbl_v180_Salary = New System.Windows.Forms.Label
        Me.Label35 = New System.Windows.Forms.Label
        Me.prg_v180_Salary = New System.Windows.Forms.ProgressBar
        Me.btn_v180_SC = New System.Windows.Forms.Button
        Me.lbl_v180_SC = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.prg_v180_SC = New System.Windows.Forms.ProgressBar
        Me.btn_v180_PR = New System.Windows.Forms.Button
        Me.lbl_v180_PR = New System.Windows.Forms.Label
        Me.Label34 = New System.Windows.Forms.Label
        Me.prg_v180_PR = New System.Windows.Forms.ProgressBar
        Me.btn_v180_multiAcc = New System.Windows.Forms.Button
        Me.lbl_v180_multiAcc = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.prg_v180_multiAcc = New System.Windows.Forms.ProgressBar
        Me.btn_v180_master_cuti = New System.Windows.Forms.Button
        Me.lbl_v180_master_cuti = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.prg_v180_master_cuti = New System.Windows.Forms.ProgressBar
        Me.btn_v180_multiunit = New System.Windows.Forms.Button
        Me.lbl_v180_multiunit = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.prg_v180_multiunit = New System.Windows.Forms.ProgressBar
        Me.TabPage9 = New System.Windows.Forms.TabPage
        Me.btn_v180_UpdateSalary = New System.Windows.Forms.Button
        Me.lbl_v180_UpdateSalary = New System.Windows.Forms.Label
        Me.Label75 = New System.Windows.Forms.Label
        Me.prg_v180_UpdateSalary = New System.Windows.Forms.ProgressBar
        Me.btn_v180_UpdateDiscAmount1 = New System.Windows.Forms.Button
        Me.Label57 = New System.Windows.Forms.Label
        Me.Label58 = New System.Windows.Forms.Label
        Me.ProgressBar10 = New System.Windows.Forms.ProgressBar
        Me.btn_v180_UpdateJurnalItem2 = New System.Windows.Forms.Button
        Me.Label50 = New System.Windows.Forms.Label
        Me.Label51 = New System.Windows.Forms.Label
        Me.ProgressBar9 = New System.Windows.Forms.ProgressBar
        Me.btn_v180_MasterCust_TrimCode = New System.Windows.Forms.Button
        Me.lbl_v180_MasterCust_TrimCode = New System.Windows.Forms.Label
        Me.Label49 = New System.Windows.Forms.Label
        Me.prg_v180_MasterCust_TrimCode = New System.Windows.Forms.ProgressBar
        Me.btn_v180_QuotationLimit = New System.Windows.Forms.Button
        Me.Label46 = New System.Windows.Forms.Label
        Me.Label47 = New System.Windows.Forms.Label
        Me.ProgressBar8 = New System.Windows.Forms.ProgressBar
        Me.btn_v180_JournalPaymentMethod = New System.Windows.Forms.Button
        Me.Label43 = New System.Windows.Forms.Label
        Me.Label45 = New System.Windows.Forms.Label
        Me.ProgressBar7 = New System.Windows.Forms.ProgressBar
        Me.btn_v180_BundlingDate = New System.Windows.Forms.Button
        Me.lbl_v180_BundlingDate = New System.Windows.Forms.Label
        Me.Label44 = New System.Windows.Forms.Label
        Me.prg_v180_BundlingDate = New System.Windows.Forms.ProgressBar
        Me.TabPage10 = New System.Windows.Forms.TabPage
        Me.Button14 = New System.Windows.Forms.Button
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.btnFilter = New System.Windows.Forms.Button
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.btnFindCodeNo = New System.Windows.Forms.Button
        Me.Label55 = New System.Windows.Forms.Label
        Me.TxtCodeNo = New System.Windows.Forms.TextBox
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage11 = New System.Windows.Forms.TabPage
        Me.dgvSC = New System.Windows.Forms.DataGridView
        Me.TabPage12 = New System.Windows.Forms.TabPage
        Me.dgvQty = New System.Windows.Forms.DataGridView
        Me.TabPage13 = New System.Windows.Forms.TabPage
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing = New System.Windows.Forms.Button
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing = New System.Windows.Forms.Label
        Me.Label88 = New System.Windows.Forms.Label
        Me.prg2018_UpdateItemProdQtyMonthlyClosing = New System.Windows.Forms.ProgressBar
        Me.btnUpdate2018_PrimarySalesP = New System.Windows.Forms.Button
        Me.Label80 = New System.Windows.Forms.Label
        Me.Label81 = New System.Windows.Forms.Label
        Me.ProgressBar17 = New System.Windows.Forms.ProgressBar
        Me.btnUpdate2018_AOP = New System.Windows.Forms.Button
        Me.Label70 = New System.Windows.Forms.Label
        Me.Label72 = New System.Windows.Forms.Label
        Me.ProgressBar14 = New System.Windows.Forms.ProgressBar
        Me.Button16 = New System.Windows.Forms.Button
        Me.Label65 = New System.Windows.Forms.Label
        Me.Label67 = New System.Windows.Forms.Label
        Me.ProgressBar13 = New System.Windows.Forms.ProgressBar
        Me.btn2018_UpdateSODO = New System.Windows.Forms.Button
        Me.LblProgress2018_UpdateSODO = New System.Windows.Forms.Label
        Me.Label66 = New System.Windows.Forms.Label
        Me.ProgressBar12 = New System.Windows.Forms.ProgressBar
        Me.btn2018_UpdateMO = New System.Windows.Forms.Button
        Me.LblProgress2018_UpdateMO = New System.Windows.Forms.Label
        Me.Label63 = New System.Windows.Forms.Label
        Me.prg2018_UpdateMO = New System.Windows.Forms.ProgressBar
        Me.Button15 = New System.Windows.Forms.Button
        Me.Label60 = New System.Windows.Forms.Label
        Me.Label61 = New System.Windows.Forms.Label
        Me.ProgressBar11 = New System.Windows.Forms.ProgressBar
        Me.TabPage14 = New System.Windows.Forms.TabPage
        Me.Button19 = New System.Windows.Forms.Button
        Me.Button18 = New System.Windows.Forms.Button
        Me.Button17 = New System.Windows.Forms.Button
        Me.TabPage15 = New System.Windows.Forms.TabPage
        Me.Label76 = New System.Windows.Forms.Label
        Me.dtpCheck_DOItem = New System.Windows.Forms.DateTimePicker
        Me.btnCheckItemDO = New System.Windows.Forms.Button
        Me.lblCheck_ItemDO = New System.Windows.Forms.Label
        Me.Label77 = New System.Windows.Forms.Label
        Me.prgCheck_ItemDO = New System.Windows.Forms.ProgressBar
        Me.BackgroundWorker7 = New System.ComponentModel.BackgroundWorker
        Me.Button11 = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Button12 = New System.Windows.Forms.Button
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.ProgressBar2 = New System.Windows.Forms.ProgressBar
        Me.BackgroundWorker8 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker9 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker10 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker11 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker12 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker13 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker14 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker15 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker16 = New System.ComponentModel.BackgroundWorker
        Me.BgwMultiunit = New System.ComponentModel.BackgroundWorker
        Me.bgwUpdatePOB = New System.ComponentModel.BackgroundWorker
        Me.bgw180_MasterCuti = New System.ComponentModel.BackgroundWorker
        Me.BgwMultiAcc = New System.ComponentModel.BackgroundWorker
        Me.bgwUpdatePrOB = New System.ComponentModel.BackgroundWorker
        Me.BgwPR = New System.ComponentModel.BackgroundWorker
        Me.Bgw_180_SC = New System.ComponentModel.BackgroundWorker
        Me.BgwSalary = New System.ComponentModel.BackgroundWorker
        Me.BgwSubtotal = New System.ComponentModel.BackgroundWorker
        Me.Bgw_MasterCust = New System.ComponentModel.BackgroundWorker
        Me.Bgw2018_UpdateMO = New System.ComponentModel.BackgroundWorker
        Me.BgwBalikJurnal = New System.ComponentModel.BackgroundWorker
        Me.BgwARPay = New System.ComponentModel.BackgroundWorker
        Me.BgwChkInv = New System.ComponentModel.BackgroundWorker
        Me.BgwClearTable = New System.ComponentModel.BackgroundWorker
        Me.Bgw_v180_UpdateSalary = New System.ComponentModel.BackgroundWorker
        Me.Bgw_Check_ItemDO = New System.ComponentModel.BackgroundWorker
        Me.BgwUpdPrimarySalesPerson = New System.ComponentModel.BackgroundWorker
        Me.TabPage16 = New System.Windows.Forms.TabPage
        Me.btnv2021 = New System.Windows.Forms.Button
        Me.btnUpdate2018_ItemKartuAPAR = New System.Windows.Forms.Button
        Me.Label87 = New System.Windows.Forms.Label
        Me.Label89 = New System.Windows.Forms.Label
        Me.ProgressBar20 = New System.Windows.Forms.ProgressBar
        Me.GroupBox1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.TabPage5.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.dgvCT_DetailTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage6.SuspendLayout()
        Me.TabPage7.SuspendLayout()
        Me.TabPage8.SuspendLayout()
        Me.TabPage9.SuspendLayout()
        Me.TabPage10.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage11.SuspendLayout()
        CType(Me.dgvSC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage12.SuspendLayout()
        CType(Me.dgvQty, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage13.SuspendLayout()
        Me.TabPage14.SuspendLayout()
        Me.TabPage15.SuspendLayout()
        Me.TabPage16.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(6, 21)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(261, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Update POB"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ProgPOB
        '
        Me.ProgPOB.Location = New System.Drawing.Point(285, 21)
        Me.ProgPOB.Name = "ProgPOB"
        Me.ProgPOB.Size = New System.Drawing.Size(620, 23)
        Me.ProgPOB.TabIndex = 1
        '
        'BackgroundWorker1
        '
        '
        'ProgSOB
        '
        Me.ProgSOB.Location = New System.Drawing.Point(285, 50)
        Me.ProgSOB.Name = "ProgSOB"
        Me.ProgSOB.Size = New System.Drawing.Size(620, 23)
        Me.ProgSOB.TabIndex = 3
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(6, 50)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(261, 23)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Update SOB"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'BackgroundWorker2
        '
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(6, 79)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(261, 23)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Update Sudah Buat Invoice - Untuk DO"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ProgUpdateDO1
        '
        Me.ProgUpdateDO1.Location = New System.Drawing.Point(285, 79)
        Me.ProgUpdateDO1.Name = "ProgUpdateDO1"
        Me.ProgUpdateDO1.Size = New System.Drawing.Size(620, 23)
        Me.ProgUpdateDO1.TabIndex = 5
        '
        'BackgroundWorker3
        '
        '
        'Prog_SeriFakturPajak
        '
        Me.Prog_SeriFakturPajak.Location = New System.Drawing.Point(285, 108)
        Me.Prog_SeriFakturPajak.Name = "Prog_SeriFakturPajak"
        Me.Prog_SeriFakturPajak.Size = New System.Drawing.Size(620, 23)
        Me.Prog_SeriFakturPajak.TabIndex = 7
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(6, 108)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(261, 23)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "Update No Seri Faktur Pajak"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'BackgroundWorker4
        '
        '
        'Prog_UpdateNoPO
        '
        Me.Prog_UpdateNoPO.Location = New System.Drawing.Point(285, 137)
        Me.Prog_UpdateNoPO.Name = "Prog_UpdateNoPO"
        Me.Prog_UpdateNoPO.Size = New System.Drawing.Size(620, 23)
        Me.Prog_UpdateNoPO.TabIndex = 9
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(6, 137)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(261, 23)
        Me.Button5.TabIndex = 8
        Me.Button5.Text = "Update No PO"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Prog_UpdateNoSO
        '
        Me.Prog_UpdateNoSO.Location = New System.Drawing.Point(285, 166)
        Me.Prog_UpdateNoSO.Name = "Prog_UpdateNoSO"
        Me.Prog_UpdateNoSO.Size = New System.Drawing.Size(620, 23)
        Me.Prog_UpdateNoSO.TabIndex = 11
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(6, 166)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(261, 23)
        Me.Button6.TabIndex = 10
        Me.Button6.Text = "Update No SO"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'ProgressBar3
        '
        Me.ProgressBar3.Location = New System.Drawing.Point(285, 195)
        Me.ProgressBar3.Name = "ProgressBar3"
        Me.ProgressBar3.Size = New System.Drawing.Size(620, 23)
        Me.ProgressBar3.TabIndex = 13
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(6, 195)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(261, 23)
        Me.Button7.TabIndex = 12
        Me.Button7.Text = "Update No DO"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'ProgressBar4
        '
        Me.ProgressBar4.Location = New System.Drawing.Point(285, 224)
        Me.ProgressBar4.Name = "ProgressBar4"
        Me.ProgressBar4.Size = New System.Drawing.Size(620, 23)
        Me.ProgressBar4.TabIndex = 15
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(6, 224)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(261, 23)
        Me.Button8.TabIndex = 14
        Me.Button8.Text = "Update No Invoice"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(6, 340)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(899, 166)
        Me.TextBox1.TabIndex = 16
        '
        'prgCost
        '
        Me.prgCost.Location = New System.Drawing.Point(285, 253)
        Me.prgCost.Name = "prgCost"
        Me.prgCost.Size = New System.Drawing.Size(620, 23)
        Me.prgCost.TabIndex = 18
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(6, 253)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(261, 23)
        Me.Button9.TabIndex = 17
        Me.Button9.Text = "Update Closing Cost"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'BackgroundWorker5
        '
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(285, 282)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(620, 23)
        Me.ProgressBar1.TabIndex = 20
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(6, 282)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(261, 23)
        Me.Button10.TabIndex = 19
        Me.Button10.Text = "Update At Amount Quick AOP"
        Me.Button10.UseVisualStyleBackColor = True
        '
        'prgItemPR
        '
        Me.prgItemPR.Location = New System.Drawing.Point(285, 311)
        Me.prgItemPR.Name = "prgItemPR"
        Me.prgItemPR.Size = New System.Drawing.Size(620, 23)
        Me.prgItemPR.TabIndex = 22
        '
        'btnItemPR
        '
        Me.btnItemPR.Location = New System.Drawing.Point(6, 311)
        Me.btnItemPR.Name = "btnItemPR"
        Me.btnItemPR.Size = New System.Drawing.Size(261, 23)
        Me.btnItemPR.TabIndex = 21
        Me.btnItemPR.Text = "Update Purchase In IDR - PR"
        Me.btnItemPR.UseVisualStyleBackColor = True
        '
        'BackgroundWorker6
        '
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.prgItemPR)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.btnItemPR)
        Me.GroupBox1.Controls.Add(Me.ProgPOB)
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button10)
        Me.GroupBox1.Controls.Add(Me.ProgSOB)
        Me.GroupBox1.Controls.Add(Me.prgCost)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Button9)
        Me.GroupBox1.Controls.Add(Me.ProgUpdateDO1)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.ProgressBar4)
        Me.GroupBox1.Controls.Add(Me.Prog_SeriFakturPajak)
        Me.GroupBox1.Controls.Add(Me.Button8)
        Me.GroupBox1.Controls.Add(Me.Button5)
        Me.GroupBox1.Controls.Add(Me.ProgressBar3)
        Me.GroupBox1.Controls.Add(Me.Prog_UpdateNoPO)
        Me.GroupBox1.Controls.Add(Me.Button7)
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.Prog_UpdateNoSO)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 10)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(911, 512)
        Me.GroupBox1.TabIndex = 23
        Me.GroupBox1.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Controls.Add(Me.TabPage5)
        Me.TabControl1.Controls.Add(Me.TabPage6)
        Me.TabControl1.Controls.Add(Me.TabPage7)
        Me.TabControl1.Controls.Add(Me.TabPage8)
        Me.TabControl1.Controls.Add(Me.TabPage9)
        Me.TabControl1.Controls.Add(Me.TabPage10)
        Me.TabControl1.Controls.Add(Me.TabPage13)
        Me.TabControl1.Controls.Add(Me.TabPage14)
        Me.TabControl1.Controls.Add(Me.TabPage15)
        Me.TabControl1.Controls.Add(Me.TabPage16)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(931, 632)
        Me.TabControl1.TabIndex = 24
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dtpBalikJurnal)
        Me.TabPage3.Controls.Add(Me.lblProgressBalikJurnal)
        Me.TabPage3.Controls.Add(Me.prgBalikJurnal)
        Me.TabPage3.Controls.Add(Me.btnBalikJurnal)
        Me.TabPage3.Controls.Add(Me.Label62)
        Me.TabPage3.Controls.Add(Me.LblProgressARPay)
        Me.TabPage3.Controls.Add(Me.dtpARPay)
        Me.TabPage3.Controls.Add(Me.prgBayarAR)
        Me.TabPage3.Controls.Add(Me.btnBayarAR)
        Me.TabPage3.Controls.Add(Me.Label64)
        Me.TabPage3.Controls.Add(Me.prgRefClosing)
        Me.TabPage3.Controls.Add(Me.dtpRefClosing)
        Me.TabPage3.Controls.Add(Me.btnDeleteClosing)
        Me.TabPage3.Controls.Add(Me.Label59)
        Me.TabPage3.Controls.Add(Me.prgRefMP)
        Me.TabPage3.Controls.Add(Me.btnCekUnbalanceAccount)
        Me.TabPage3.Controls.Add(Me.Label56)
        Me.TabPage3.Controls.Add(Me.dtpRefMP)
        Me.TabPage3.Controls.Add(Me.btnUpdatePrimaryAkun)
        Me.TabPage3.Controls.Add(Me.Label54)
        Me.TabPage3.Controls.Add(Me.btnCekSalaryKomponen)
        Me.TabPage3.Controls.Add(Me.Label53)
        Me.TabPage3.Controls.Add(Me.prgUpdatePenghasilanBruto)
        Me.TabPage3.Controls.Add(Me.btnUpdatePenghasilanBruto)
        Me.TabPage3.Controls.Add(Me.Label52)
        Me.TabPage3.Controls.Add(Me.btnDeleteMPSF)
        Me.TabPage3.Controls.Add(Me.Label48)
        Me.TabPage3.Controls.Add(Me.TxtAccountDesc_COGS)
        Me.TabPage3.Controls.Add(Me.TextBox15)
        Me.TabPage3.Controls.Add(Me.TxtPrimaryAccount_COGS)
        Me.TabPage3.Controls.Add(Me.btnUpdateAccount_COGS)
        Me.TabPage3.Controls.Add(Me.Label7)
        Me.TabPage3.Controls.Add(Me.TxtAccount_COGS)
        Me.TabPage3.Controls.Add(Me.btnFindAccount_COGS)
        Me.TabPage3.Controls.Add(Me.TxtAccountDesc_SalesReturn)
        Me.TabPage3.Controls.Add(Me.TextBox11)
        Me.TabPage3.Controls.Add(Me.TxtPrimaryAccount_SalesReturn)
        Me.TabPage3.Controls.Add(Me.btnUpdateAccount_SalesReturn)
        Me.TabPage3.Controls.Add(Me.Label6)
        Me.TabPage3.Controls.Add(Me.TxtAccount_SalesReturn)
        Me.TabPage3.Controls.Add(Me.btnFindAccount_SalesReturn)
        Me.TabPage3.Controls.Add(Me.TxtAccountDesc_Sales)
        Me.TabPage3.Controls.Add(Me.TextBox4)
        Me.TabPage3.Controls.Add(Me.TxtPrimaryAccount_Sales)
        Me.TabPage3.Controls.Add(Me.btnUpdateAccount_Sales)
        Me.TabPage3.Controls.Add(Me.Label3)
        Me.TabPage3.Controls.Add(Me.TxtAccount_Sales)
        Me.TabPage3.Controls.Add(Me.btnFindAccount_Sales)
        Me.TabPage3.Controls.Add(Me.TxtAccountDesc_Inventory)
        Me.TabPage3.Controls.Add(Me.TextBox8)
        Me.TabPage3.Controls.Add(Me.TxtPrimaryAccount_Inventory)
        Me.TabPage3.Controls.Add(Me.btnUpdateAccount_Inventory)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.TxtAccount_Inventory)
        Me.TabPage3.Controls.Add(Me.btnFindAccount_Inventory)
        Me.TabPage3.Controls.Add(Me.TxtAccountDesc_Adjustment)
        Me.TabPage3.Controls.Add(Me.TextBox3)
        Me.TabPage3.Controls.Add(Me.TxtPrimaryAccount_Adjustment)
        Me.TabPage3.Controls.Add(Me.btnUpdateAccount_Adjustment)
        Me.TabPage3.Controls.Add(Me.Label1)
        Me.TabPage3.Controls.Add(Me.TxtAccount_Adjustment)
        Me.TabPage3.Controls.Add(Me.btnFindAccount_Adjustment)
        Me.TabPage3.Controls.Add(Me.TxtAccountDesc_StockTaking)
        Me.TabPage3.Controls.Add(Me.TxtReadOnly)
        Me.TabPage3.Controls.Add(Me.TxtPrimaryAccount_StockTaking)
        Me.TabPage3.Controls.Add(Me.btnUpdateAccount_StockTaking)
        Me.TabPage3.Controls.Add(Me.Label11)
        Me.TabPage3.Controls.Add(Me.TxtAccount_StockTaking)
        Me.TabPage3.Controls.Add(Me.btnFindAccount_StockTaking)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(923, 603)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "General"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dtpBalikJurnal
        '
        Me.dtpBalikJurnal.Checked = False
        Me.dtpBalikJurnal.CustomFormat = "dd MMM yyyy"
        Me.dtpBalikJurnal.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpBalikJurnal.Location = New System.Drawing.Point(342, 488)
        Me.dtpBalikJurnal.Name = "dtpBalikJurnal"
        Me.dtpBalikJurnal.ShowUpDown = True
        Me.dtpBalikJurnal.Size = New System.Drawing.Size(200, 22)
        Me.dtpBalikJurnal.TabIndex = 93
        '
        'lblProgressBalikJurnal
        '
        Me.lblProgressBalikJurnal.AutoSize = True
        Me.lblProgressBalikJurnal.Location = New System.Drawing.Point(339, 511)
        Me.lblProgressBalikJurnal.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProgressBalikJurnal.Name = "lblProgressBalikJurnal"
        Me.lblProgressBalikJurnal.Size = New System.Drawing.Size(26, 17)
        Me.lblProgressBalikJurnal.TabIndex = 92
        Me.lblProgressBalikJurnal.Text = "xxx"
        '
        'prgBalikJurnal
        '
        Me.prgBalikJurnal.Location = New System.Drawing.Point(548, 485)
        Me.prgBalikJurnal.Name = "prgBalikJurnal"
        Me.prgBalikJurnal.Size = New System.Drawing.Size(278, 23)
        Me.prgBalikJurnal.TabIndex = 91
        '
        'btnBalikJurnal
        '
        Me.btnBalikJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBalikJurnal.Location = New System.Drawing.Point(834, 482)
        Me.btnBalikJurnal.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBalikJurnal.Name = "btnBalikJurnal"
        Me.btnBalikJurnal.Size = New System.Drawing.Size(82, 28)
        Me.btnBalikJurnal.TabIndex = 90
        Me.btnBalikJurnal.TabStop = False
        Me.btnBalikJurnal.Text = "Update"
        Me.btnBalikJurnal.UseVisualStyleBackColor = True
        '
        'Label62
        '
        Me.Label62.AutoSize = True
        Me.Label62.Location = New System.Drawing.Point(9, 488)
        Me.Label62.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(81, 17)
        Me.Label62.TabIndex = 89
        Me.Label62.Text = "Balik Jurnal"
        '
        'LblProgressARPay
        '
        Me.LblProgressARPay.AutoSize = True
        Me.LblProgressARPay.Location = New System.Drawing.Point(339, 465)
        Me.LblProgressARPay.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblProgressARPay.Name = "LblProgressARPay"
        Me.LblProgressARPay.Size = New System.Drawing.Size(26, 17)
        Me.LblProgressARPay.TabIndex = 88
        Me.LblProgressARPay.Text = "xxx"
        '
        'dtpARPay
        '
        Me.dtpARPay.Checked = False
        Me.dtpARPay.CustomFormat = "dd MMM yyyy"
        Me.dtpARPay.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpARPay.Location = New System.Drawing.Point(342, 437)
        Me.dtpARPay.Name = "dtpARPay"
        Me.dtpARPay.ShowUpDown = True
        Me.dtpARPay.Size = New System.Drawing.Size(200, 22)
        Me.dtpARPay.TabIndex = 87
        '
        'prgBayarAR
        '
        Me.prgBayarAR.Location = New System.Drawing.Point(548, 439)
        Me.prgBayarAR.Name = "prgBayarAR"
        Me.prgBayarAR.Size = New System.Drawing.Size(278, 23)
        Me.prgBayarAR.TabIndex = 86
        '
        'btnBayarAR
        '
        Me.btnBayarAR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBayarAR.Location = New System.Drawing.Point(834, 436)
        Me.btnBayarAR.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBayarAR.Name = "btnBayarAR"
        Me.btnBayarAR.Size = New System.Drawing.Size(82, 28)
        Me.btnBayarAR.TabIndex = 85
        Me.btnBayarAR.TabStop = False
        Me.btnBayarAR.Text = "Update"
        Me.btnBayarAR.UseVisualStyleBackColor = True
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Location = New System.Drawing.Point(9, 442)
        Me.Label64.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(116, 17)
        Me.Label64.TabIndex = 84
        Me.Label64.Text = "Bayar Semua AR"
        '
        'prgRefClosing
        '
        Me.prgRefClosing.Location = New System.Drawing.Point(548, 403)
        Me.prgRefClosing.Name = "prgRefClosing"
        Me.prgRefClosing.Size = New System.Drawing.Size(278, 23)
        Me.prgRefClosing.TabIndex = 73
        '
        'dtpRefClosing
        '
        Me.dtpRefClosing.Checked = False
        Me.dtpRefClosing.CustomFormat = "MMM yyyy"
        Me.dtpRefClosing.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpRefClosing.Location = New System.Drawing.Point(342, 404)
        Me.dtpRefClosing.Name = "dtpRefClosing"
        Me.dtpRefClosing.ShowCheckBox = True
        Me.dtpRefClosing.ShowUpDown = True
        Me.dtpRefClosing.Size = New System.Drawing.Size(200, 22)
        Me.dtpRefClosing.TabIndex = 72
        '
        'btnDeleteClosing
        '
        Me.btnDeleteClosing.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeleteClosing.Location = New System.Drawing.Point(834, 400)
        Me.btnDeleteClosing.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDeleteClosing.Name = "btnDeleteClosing"
        Me.btnDeleteClosing.Size = New System.Drawing.Size(82, 28)
        Me.btnDeleteClosing.TabIndex = 71
        Me.btnDeleteClosing.TabStop = False
        Me.btnDeleteClosing.Text = "Update"
        Me.btnDeleteClosing.UseVisualStyleBackColor = True
        '
        'Label59
        '
        Me.Label59.AutoSize = True
        Me.Label59.Location = New System.Drawing.Point(9, 406)
        Me.Label59.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(163, 17)
        Me.Label59.TabIndex = 70
        Me.Label59.Text = "Delete Purchase Closing"
        '
        'prgRefMP
        '
        Me.prgRefMP.Location = New System.Drawing.Point(548, 223)
        Me.prgRefMP.Name = "prgRefMP"
        Me.prgRefMP.Size = New System.Drawing.Size(278, 23)
        Me.prgRefMP.TabIndex = 69
        '
        'btnCekUnbalanceAccount
        '
        Me.btnCekUnbalanceAccount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCekUnbalanceAccount.Location = New System.Drawing.Point(834, 364)
        Me.btnCekUnbalanceAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCekUnbalanceAccount.Name = "btnCekUnbalanceAccount"
        Me.btnCekUnbalanceAccount.Size = New System.Drawing.Size(82, 28)
        Me.btnCekUnbalanceAccount.TabIndex = 68
        Me.btnCekUnbalanceAccount.TabStop = False
        Me.btnCekUnbalanceAccount.Text = "Cek"
        Me.btnCekUnbalanceAccount.UseVisualStyleBackColor = True
        '
        'Label56
        '
        Me.Label56.AutoSize = True
        Me.Label56.Location = New System.Drawing.Point(9, 370)
        Me.Label56.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(157, 17)
        Me.Label56.TabIndex = 67
        Me.Label56.Text = "Cek unbalance Account"
        '
        'dtpRefMP
        '
        Me.dtpRefMP.Checked = False
        Me.dtpRefMP.CustomFormat = "MMM yyyy"
        Me.dtpRefMP.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpRefMP.Location = New System.Drawing.Point(342, 224)
        Me.dtpRefMP.Name = "dtpRefMP"
        Me.dtpRefMP.ShowCheckBox = True
        Me.dtpRefMP.ShowUpDown = True
        Me.dtpRefMP.Size = New System.Drawing.Size(200, 22)
        Me.dtpRefMP.TabIndex = 66
        '
        'btnUpdatePrimaryAkun
        '
        Me.btnUpdatePrimaryAkun.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatePrimaryAkun.Location = New System.Drawing.Point(834, 328)
        Me.btnUpdatePrimaryAkun.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatePrimaryAkun.Name = "btnUpdatePrimaryAkun"
        Me.btnUpdatePrimaryAkun.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatePrimaryAkun.TabIndex = 65
        Me.btnUpdatePrimaryAkun.TabStop = False
        Me.btnUpdatePrimaryAkun.Text = "Update"
        Me.btnUpdatePrimaryAkun.UseVisualStyleBackColor = True
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Location = New System.Drawing.Point(9, 334)
        Me.Label54.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(250, 17)
        Me.Label54.TabIndex = 64
        Me.Label54.Text = "Update Primary Akun pada item salary"
        '
        'btnCekSalaryKomponen
        '
        Me.btnCekSalaryKomponen.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCekSalaryKomponen.Location = New System.Drawing.Point(834, 292)
        Me.btnCekSalaryKomponen.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCekSalaryKomponen.Name = "btnCekSalaryKomponen"
        Me.btnCekSalaryKomponen.Size = New System.Drawing.Size(82, 28)
        Me.btnCekSalaryKomponen.TabIndex = 63
        Me.btnCekSalaryKomponen.TabStop = False
        Me.btnCekSalaryKomponen.Text = "Cek"
        Me.btnCekSalaryKomponen.UseVisualStyleBackColor = True
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Location = New System.Drawing.Point(9, 298)
        Me.Label53.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(452, 17)
        Me.Label53.TabIndex = 62
        Me.Label53.Text = "Cek apakah semua akun salary komponen terdaftar pada master akun"
        '
        'prgUpdatePenghasilanBruto
        '
        Me.prgUpdatePenghasilanBruto.Location = New System.Drawing.Point(338, 259)
        Me.prgUpdatePenghasilanBruto.Name = "prgUpdatePenghasilanBruto"
        Me.prgUpdatePenghasilanBruto.Size = New System.Drawing.Size(488, 23)
        Me.prgUpdatePenghasilanBruto.TabIndex = 61
        '
        'btnUpdatePenghasilanBruto
        '
        Me.btnUpdatePenghasilanBruto.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatePenghasilanBruto.Location = New System.Drawing.Point(834, 256)
        Me.btnUpdatePenghasilanBruto.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatePenghasilanBruto.Name = "btnUpdatePenghasilanBruto"
        Me.btnUpdatePenghasilanBruto.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatePenghasilanBruto.TabIndex = 60
        Me.btnUpdatePenghasilanBruto.TabStop = False
        Me.btnUpdatePenghasilanBruto.Text = "Update"
        Me.btnUpdatePenghasilanBruto.UseVisualStyleBackColor = True
        '
        'Label52
        '
        Me.Label52.AutoSize = True
        Me.Label52.Location = New System.Drawing.Point(9, 262)
        Me.Label52.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(174, 17)
        Me.Label52.TabIndex = 59
        Me.Label52.Text = "Update Penghasilan Bruto"
        '
        'btnDeleteMPSF
        '
        Me.btnDeleteMPSF.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeleteMPSF.Location = New System.Drawing.Point(834, 220)
        Me.btnDeleteMPSF.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDeleteMPSF.Name = "btnDeleteMPSF"
        Me.btnDeleteMPSF.Size = New System.Drawing.Size(82, 28)
        Me.btnDeleteMPSF.TabIndex = 58
        Me.btnDeleteMPSF.TabStop = False
        Me.btnDeleteMPSF.Text = "Update"
        Me.btnDeleteMPSF.UseVisualStyleBackColor = True
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.Location = New System.Drawing.Point(9, 226)
        Me.Label48.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(141, 17)
        Me.Label48.TabIndex = 57
        Me.Label48.Text = "Delete MP && SF Data"
        '
        'TxtAccountDesc_COGS
        '
        Me.TxtAccountDesc_COGS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountDesc_COGS.Location = New System.Drawing.Point(424, 187)
        Me.TxtAccountDesc_COGS.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccountDesc_COGS.Name = "TxtAccountDesc_COGS"
        Me.TxtAccountDesc_COGS.ReadOnly = True
        Me.TxtAccountDesc_COGS.Size = New System.Drawing.Size(402, 22)
        Me.TxtAccountDesc_COGS.TabIndex = 49
        Me.TxtAccountDesc_COGS.TabStop = False
        '
        'TextBox15
        '
        Me.TextBox15.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox15.Location = New System.Drawing.Point(424, 187)
        Me.TextBox15.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox15.Name = "TextBox15"
        Me.TextBox15.ReadOnly = True
        Me.TextBox15.Size = New System.Drawing.Size(94, 22)
        Me.TextBox15.TabIndex = 53
        Me.TextBox15.TabStop = False
        '
        'TxtPrimaryAccount_COGS
        '
        Me.TxtPrimaryAccount_COGS.Location = New System.Drawing.Point(127, 187)
        Me.TxtPrimaryAccount_COGS.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryAccount_COGS.Name = "TxtPrimaryAccount_COGS"
        Me.TxtPrimaryAccount_COGS.Size = New System.Drawing.Size(49, 22)
        Me.TxtPrimaryAccount_COGS.TabIndex = 52
        Me.TxtPrimaryAccount_COGS.TabStop = False
        Me.TxtPrimaryAccount_COGS.Text = "-1"
        Me.TxtPrimaryAccount_COGS.Visible = False
        '
        'btnUpdateAccount_COGS
        '
        Me.btnUpdateAccount_COGS.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateAccount_COGS.Location = New System.Drawing.Point(834, 184)
        Me.btnUpdateAccount_COGS.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateAccount_COGS.Name = "btnUpdateAccount_COGS"
        Me.btnUpdateAccount_COGS.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateAccount_COGS.TabIndex = 51
        Me.btnUpdateAccount_COGS.TabStop = False
        Me.btnUpdateAccount_COGS.Text = "Update"
        Me.btnUpdateAccount_COGS.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(9, 190)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(103, 17)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "COGS Account"
        '
        'TxtAccount_COGS
        '
        Me.TxtAccount_COGS.Location = New System.Drawing.Point(184, 187)
        Me.TxtAccount_COGS.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccount_COGS.Name = "TxtAccount_COGS"
        Me.TxtAccount_COGS.Size = New System.Drawing.Size(146, 22)
        Me.TxtAccount_COGS.TabIndex = 47
        '
        'btnFindAccount_COGS
        '
        Me.btnFindAccount_COGS.Location = New System.Drawing.Point(338, 184)
        Me.btnFindAccount_COGS.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindAccount_COGS.Name = "btnFindAccount_COGS"
        Me.btnFindAccount_COGS.Size = New System.Drawing.Size(78, 28)
        Me.btnFindAccount_COGS.TabIndex = 48
        Me.btnFindAccount_COGS.TabStop = False
        Me.btnFindAccount_COGS.Text = "Find"
        Me.btnFindAccount_COGS.UseVisualStyleBackColor = True
        '
        'TxtAccountDesc_SalesReturn
        '
        Me.TxtAccountDesc_SalesReturn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountDesc_SalesReturn.Location = New System.Drawing.Point(424, 151)
        Me.TxtAccountDesc_SalesReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccountDesc_SalesReturn.Name = "TxtAccountDesc_SalesReturn"
        Me.TxtAccountDesc_SalesReturn.ReadOnly = True
        Me.TxtAccountDesc_SalesReturn.Size = New System.Drawing.Size(402, 22)
        Me.TxtAccountDesc_SalesReturn.TabIndex = 42
        Me.TxtAccountDesc_SalesReturn.TabStop = False
        '
        'TextBox11
        '
        Me.TextBox11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox11.Location = New System.Drawing.Point(424, 151)
        Me.TextBox11.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.ReadOnly = True
        Me.TextBox11.Size = New System.Drawing.Size(94, 22)
        Me.TextBox11.TabIndex = 46
        Me.TextBox11.TabStop = False
        '
        'TxtPrimaryAccount_SalesReturn
        '
        Me.TxtPrimaryAccount_SalesReturn.Location = New System.Drawing.Point(127, 151)
        Me.TxtPrimaryAccount_SalesReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryAccount_SalesReturn.Name = "TxtPrimaryAccount_SalesReturn"
        Me.TxtPrimaryAccount_SalesReturn.Size = New System.Drawing.Size(49, 22)
        Me.TxtPrimaryAccount_SalesReturn.TabIndex = 45
        Me.TxtPrimaryAccount_SalesReturn.TabStop = False
        Me.TxtPrimaryAccount_SalesReturn.Text = "-1"
        Me.TxtPrimaryAccount_SalesReturn.Visible = False
        '
        'btnUpdateAccount_SalesReturn
        '
        Me.btnUpdateAccount_SalesReturn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateAccount_SalesReturn.Location = New System.Drawing.Point(834, 148)
        Me.btnUpdateAccount_SalesReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateAccount_SalesReturn.Name = "btnUpdateAccount_SalesReturn"
        Me.btnUpdateAccount_SalesReturn.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateAccount_SalesReturn.TabIndex = 44
        Me.btnUpdateAccount_SalesReturn.TabStop = False
        Me.btnUpdateAccount_SalesReturn.Text = "Update"
        Me.btnUpdateAccount_SalesReturn.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(9, 154)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(145, 17)
        Me.Label6.TabIndex = 43
        Me.Label6.Text = "Sales Return Account"
        '
        'TxtAccount_SalesReturn
        '
        Me.TxtAccount_SalesReturn.Location = New System.Drawing.Point(184, 151)
        Me.TxtAccount_SalesReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccount_SalesReturn.Name = "TxtAccount_SalesReturn"
        Me.TxtAccount_SalesReturn.Size = New System.Drawing.Size(146, 22)
        Me.TxtAccount_SalesReturn.TabIndex = 40
        '
        'btnFindAccount_SalesReturn
        '
        Me.btnFindAccount_SalesReturn.Location = New System.Drawing.Point(338, 148)
        Me.btnFindAccount_SalesReturn.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindAccount_SalesReturn.Name = "btnFindAccount_SalesReturn"
        Me.btnFindAccount_SalesReturn.Size = New System.Drawing.Size(78, 28)
        Me.btnFindAccount_SalesReturn.TabIndex = 41
        Me.btnFindAccount_SalesReturn.TabStop = False
        Me.btnFindAccount_SalesReturn.Text = "Find"
        Me.btnFindAccount_SalesReturn.UseVisualStyleBackColor = True
        '
        'TxtAccountDesc_Sales
        '
        Me.TxtAccountDesc_Sales.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountDesc_Sales.Location = New System.Drawing.Point(424, 115)
        Me.TxtAccountDesc_Sales.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccountDesc_Sales.Name = "TxtAccountDesc_Sales"
        Me.TxtAccountDesc_Sales.ReadOnly = True
        Me.TxtAccountDesc_Sales.Size = New System.Drawing.Size(402, 22)
        Me.TxtAccountDesc_Sales.TabIndex = 35
        Me.TxtAccountDesc_Sales.TabStop = False
        '
        'TextBox4
        '
        Me.TextBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox4.Location = New System.Drawing.Point(424, 115)
        Me.TextBox4.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ReadOnly = True
        Me.TextBox4.Size = New System.Drawing.Size(94, 22)
        Me.TextBox4.TabIndex = 39
        Me.TextBox4.TabStop = False
        '
        'TxtPrimaryAccount_Sales
        '
        Me.TxtPrimaryAccount_Sales.Location = New System.Drawing.Point(127, 115)
        Me.TxtPrimaryAccount_Sales.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryAccount_Sales.Name = "TxtPrimaryAccount_Sales"
        Me.TxtPrimaryAccount_Sales.Size = New System.Drawing.Size(49, 22)
        Me.TxtPrimaryAccount_Sales.TabIndex = 38
        Me.TxtPrimaryAccount_Sales.TabStop = False
        Me.TxtPrimaryAccount_Sales.Text = "-1"
        Me.TxtPrimaryAccount_Sales.Visible = False
        '
        'btnUpdateAccount_Sales
        '
        Me.btnUpdateAccount_Sales.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateAccount_Sales.Location = New System.Drawing.Point(834, 112)
        Me.btnUpdateAccount_Sales.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateAccount_Sales.Name = "btnUpdateAccount_Sales"
        Me.btnUpdateAccount_Sales.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateAccount_Sales.TabIndex = 37
        Me.btnUpdateAccount_Sales.TabStop = False
        Me.btnUpdateAccount_Sales.Text = "Update"
        Me.btnUpdateAccount_Sales.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 118)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(98, 17)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "Sales Account"
        '
        'TxtAccount_Sales
        '
        Me.TxtAccount_Sales.Location = New System.Drawing.Point(184, 115)
        Me.TxtAccount_Sales.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccount_Sales.Name = "TxtAccount_Sales"
        Me.TxtAccount_Sales.Size = New System.Drawing.Size(146, 22)
        Me.TxtAccount_Sales.TabIndex = 33
        '
        'btnFindAccount_Sales
        '
        Me.btnFindAccount_Sales.Location = New System.Drawing.Point(338, 112)
        Me.btnFindAccount_Sales.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindAccount_Sales.Name = "btnFindAccount_Sales"
        Me.btnFindAccount_Sales.Size = New System.Drawing.Size(78, 28)
        Me.btnFindAccount_Sales.TabIndex = 34
        Me.btnFindAccount_Sales.TabStop = False
        Me.btnFindAccount_Sales.Text = "Find"
        Me.btnFindAccount_Sales.UseVisualStyleBackColor = True
        '
        'TxtAccountDesc_Inventory
        '
        Me.TxtAccountDesc_Inventory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountDesc_Inventory.Location = New System.Drawing.Point(424, 79)
        Me.TxtAccountDesc_Inventory.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccountDesc_Inventory.Name = "TxtAccountDesc_Inventory"
        Me.TxtAccountDesc_Inventory.ReadOnly = True
        Me.TxtAccountDesc_Inventory.Size = New System.Drawing.Size(402, 22)
        Me.TxtAccountDesc_Inventory.TabIndex = 28
        Me.TxtAccountDesc_Inventory.TabStop = False
        '
        'TextBox8
        '
        Me.TextBox8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox8.Location = New System.Drawing.Point(424, 79)
        Me.TextBox8.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.ReadOnly = True
        Me.TextBox8.Size = New System.Drawing.Size(94, 22)
        Me.TextBox8.TabIndex = 32
        Me.TextBox8.TabStop = False
        '
        'TxtPrimaryAccount_Inventory
        '
        Me.TxtPrimaryAccount_Inventory.Location = New System.Drawing.Point(127, 79)
        Me.TxtPrimaryAccount_Inventory.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryAccount_Inventory.Name = "TxtPrimaryAccount_Inventory"
        Me.TxtPrimaryAccount_Inventory.Size = New System.Drawing.Size(49, 22)
        Me.TxtPrimaryAccount_Inventory.TabIndex = 31
        Me.TxtPrimaryAccount_Inventory.TabStop = False
        Me.TxtPrimaryAccount_Inventory.Text = "-1"
        Me.TxtPrimaryAccount_Inventory.Visible = False
        '
        'btnUpdateAccount_Inventory
        '
        Me.btnUpdateAccount_Inventory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateAccount_Inventory.Location = New System.Drawing.Point(834, 76)
        Me.btnUpdateAccount_Inventory.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateAccount_Inventory.Name = "btnUpdateAccount_Inventory"
        Me.btnUpdateAccount_Inventory.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateAccount_Inventory.TabIndex = 30
        Me.btnUpdateAccount_Inventory.TabStop = False
        Me.btnUpdateAccount_Inventory.Text = "Update"
        Me.btnUpdateAccount_Inventory.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 82)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(121, 17)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Inventory Account"
        '
        'TxtAccount_Inventory
        '
        Me.TxtAccount_Inventory.Location = New System.Drawing.Point(184, 79)
        Me.TxtAccount_Inventory.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccount_Inventory.Name = "TxtAccount_Inventory"
        Me.TxtAccount_Inventory.Size = New System.Drawing.Size(146, 22)
        Me.TxtAccount_Inventory.TabIndex = 26
        '
        'btnFindAccount_Inventory
        '
        Me.btnFindAccount_Inventory.Location = New System.Drawing.Point(338, 76)
        Me.btnFindAccount_Inventory.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindAccount_Inventory.Name = "btnFindAccount_Inventory"
        Me.btnFindAccount_Inventory.Size = New System.Drawing.Size(78, 28)
        Me.btnFindAccount_Inventory.TabIndex = 27
        Me.btnFindAccount_Inventory.TabStop = False
        Me.btnFindAccount_Inventory.Text = "Find"
        Me.btnFindAccount_Inventory.UseVisualStyleBackColor = True
        '
        'TxtAccountDesc_Adjustment
        '
        Me.TxtAccountDesc_Adjustment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountDesc_Adjustment.Location = New System.Drawing.Point(424, 43)
        Me.TxtAccountDesc_Adjustment.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccountDesc_Adjustment.Name = "TxtAccountDesc_Adjustment"
        Me.TxtAccountDesc_Adjustment.ReadOnly = True
        Me.TxtAccountDesc_Adjustment.Size = New System.Drawing.Size(402, 22)
        Me.TxtAccountDesc_Adjustment.TabIndex = 21
        Me.TxtAccountDesc_Adjustment.TabStop = False
        '
        'TextBox3
        '
        Me.TextBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox3.Location = New System.Drawing.Point(424, 43)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(94, 22)
        Me.TextBox3.TabIndex = 25
        Me.TextBox3.TabStop = False
        '
        'TxtPrimaryAccount_Adjustment
        '
        Me.TxtPrimaryAccount_Adjustment.Location = New System.Drawing.Point(127, 43)
        Me.TxtPrimaryAccount_Adjustment.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryAccount_Adjustment.Name = "TxtPrimaryAccount_Adjustment"
        Me.TxtPrimaryAccount_Adjustment.Size = New System.Drawing.Size(49, 22)
        Me.TxtPrimaryAccount_Adjustment.TabIndex = 24
        Me.TxtPrimaryAccount_Adjustment.TabStop = False
        Me.TxtPrimaryAccount_Adjustment.Text = "-1"
        Me.TxtPrimaryAccount_Adjustment.Visible = False
        '
        'btnUpdateAccount_Adjustment
        '
        Me.btnUpdateAccount_Adjustment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateAccount_Adjustment.Location = New System.Drawing.Point(834, 40)
        Me.btnUpdateAccount_Adjustment.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateAccount_Adjustment.Name = "btnUpdateAccount_Adjustment"
        Me.btnUpdateAccount_Adjustment.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateAccount_Adjustment.TabIndex = 23
        Me.btnUpdateAccount_Adjustment.TabStop = False
        Me.btnUpdateAccount_Adjustment.Text = "Update"
        Me.btnUpdateAccount_Adjustment.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 46)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(133, 17)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Adjustment Account"
        '
        'TxtAccount_Adjustment
        '
        Me.TxtAccount_Adjustment.Location = New System.Drawing.Point(184, 43)
        Me.TxtAccount_Adjustment.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccount_Adjustment.Name = "TxtAccount_Adjustment"
        Me.TxtAccount_Adjustment.Size = New System.Drawing.Size(146, 22)
        Me.TxtAccount_Adjustment.TabIndex = 19
        '
        'btnFindAccount_Adjustment
        '
        Me.btnFindAccount_Adjustment.Location = New System.Drawing.Point(338, 40)
        Me.btnFindAccount_Adjustment.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindAccount_Adjustment.Name = "btnFindAccount_Adjustment"
        Me.btnFindAccount_Adjustment.Size = New System.Drawing.Size(78, 28)
        Me.btnFindAccount_Adjustment.TabIndex = 20
        Me.btnFindAccount_Adjustment.TabStop = False
        Me.btnFindAccount_Adjustment.Text = "Find"
        Me.btnFindAccount_Adjustment.UseVisualStyleBackColor = True
        '
        'TxtAccountDesc_StockTaking
        '
        Me.TxtAccountDesc_StockTaking.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAccountDesc_StockTaking.Location = New System.Drawing.Point(424, 7)
        Me.TxtAccountDesc_StockTaking.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccountDesc_StockTaking.Name = "TxtAccountDesc_StockTaking"
        Me.TxtAccountDesc_StockTaking.ReadOnly = True
        Me.TxtAccountDesc_StockTaking.Size = New System.Drawing.Size(402, 22)
        Me.TxtAccountDesc_StockTaking.TabIndex = 14
        Me.TxtAccountDesc_StockTaking.TabStop = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(424, 7)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(94, 22)
        Me.TxtReadOnly.TabIndex = 18
        Me.TxtReadOnly.TabStop = False
        '
        'TxtPrimaryAccount_StockTaking
        '
        Me.TxtPrimaryAccount_StockTaking.Location = New System.Drawing.Point(127, 7)
        Me.TxtPrimaryAccount_StockTaking.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryAccount_StockTaking.Name = "TxtPrimaryAccount_StockTaking"
        Me.TxtPrimaryAccount_StockTaking.Size = New System.Drawing.Size(49, 22)
        Me.TxtPrimaryAccount_StockTaking.TabIndex = 17
        Me.TxtPrimaryAccount_StockTaking.TabStop = False
        Me.TxtPrimaryAccount_StockTaking.Text = "-1"
        Me.TxtPrimaryAccount_StockTaking.Visible = False
        '
        'btnUpdateAccount_StockTaking
        '
        Me.btnUpdateAccount_StockTaking.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateAccount_StockTaking.Location = New System.Drawing.Point(834, 4)
        Me.btnUpdateAccount_StockTaking.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateAccount_StockTaking.Name = "btnUpdateAccount_StockTaking"
        Me.btnUpdateAccount_StockTaking.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateAccount_StockTaking.TabIndex = 16
        Me.btnUpdateAccount_StockTaking.TabStop = False
        Me.btnUpdateAccount_StockTaking.Text = "Update"
        Me.btnUpdateAccount_StockTaking.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(9, 10)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(145, 17)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "Stock Taking Account"
        '
        'TxtAccount_StockTaking
        '
        Me.TxtAccount_StockTaking.Location = New System.Drawing.Point(184, 7)
        Me.TxtAccount_StockTaking.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAccount_StockTaking.Name = "TxtAccount_StockTaking"
        Me.TxtAccount_StockTaking.Size = New System.Drawing.Size(146, 22)
        Me.TxtAccount_StockTaking.TabIndex = 12
        '
        'btnFindAccount_StockTaking
        '
        Me.btnFindAccount_StockTaking.Location = New System.Drawing.Point(338, 4)
        Me.btnFindAccount_StockTaking.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindAccount_StockTaking.Name = "btnFindAccount_StockTaking"
        Me.btnFindAccount_StockTaking.Size = New System.Drawing.Size(78, 28)
        Me.btnFindAccount_StockTaking.TabIndex = 13
        Me.btnFindAccount_StockTaking.TabStop = False
        Me.btnFindAccount_StockTaking.Text = "Find"
        Me.btnFindAccount_StockTaking.UseVisualStyleBackColor = True
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(923, 603)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "v.1.5.29"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Prg_UpdateDiscAmount)
        Me.TabPage2.Controls.Add(Me.btnUpdateDiscAmount)
        Me.TabPage2.Controls.Add(Me.LblUpdateDiscAmount)
        Me.TabPage2.Controls.Add(Me.Label24)
        Me.TabPage2.Controls.Add(Me.ProgressBar5)
        Me.TabPage2.Controls.Add(Me.btnUpdataHR)
        Me.TabPage2.Controls.Add(Me.lblUpdatev160_HR)
        Me.TabPage2.Controls.Add(Me.Label23)
        Me.TabPage2.Controls.Add(Me.btnUpdatev160_TBLMasterProduct)
        Me.TabPage2.Controls.Add(Me.LblUpdatev160_TBLMasterProduct)
        Me.TabPage2.Controls.Add(Me.Label19)
        Me.TabPage2.Controls.Add(Me.prgUpdatev160_TambahTabelBaru)
        Me.TabPage2.Controls.Add(Me.btnUpdatev160_TambahTabelBaru)
        Me.TabPage2.Controls.Add(Me.lblUpdatev160_TambahTabelBaru)
        Me.TabPage2.Controls.Add(Me.Label18)
        Me.TabPage2.Controls.Add(Me.prgIdentifierPO)
        Me.TabPage2.Controls.Add(Me.btnUpdateIdentifierPO)
        Me.TabPage2.Controls.Add(Me.lblProgressPO)
        Me.TabPage2.Controls.Add(Me.Label17)
        Me.TabPage2.Controls.Add(Me.prgIdentifierSO)
        Me.TabPage2.Controls.Add(Me.btnUpdate_IdentifierSO)
        Me.TabPage2.Controls.Add(Me.LblIdentifierSO)
        Me.TabPage2.Controls.Add(Me.Label16)
        Me.TabPage2.Controls.Add(Me.prgUpdateSalesPrice)
        Me.TabPage2.Controls.Add(Me.btnPindahkanSalesPrice)
        Me.TabPage2.Controls.Add(Me.LblProgres_PindahSalesPrice)
        Me.TabPage2.Controls.Add(Me.Label15)
        Me.TabPage2.Controls.Add(Me.Label13)
        Me.TabPage2.Controls.Add(Me.btnUpdatev160_UpdateNumberingMP)
        Me.TabPage2.Controls.Add(Me.LblProgress_UpdateNumberingMaterialPlanning)
        Me.TabPage2.Controls.Add(Me.LblProgress_UpdateNumbering)
        Me.TabPage2.Controls.Add(Me.btnUpdatev160_UpdateNumberingCL)
        Me.TabPage2.Controls.Add(Me.LblProgress_UpdateNumberingColletionLetter)
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.btnUpdatev160_SC)
        Me.TabPage2.Controls.Add(Me.lblPrgStockCard)
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.prgStockCard)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(923, 603)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "v.1.6.0"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Prg_UpdateDiscAmount
        '
        Me.Prg_UpdateDiscAmount.Location = New System.Drawing.Point(245, 422)
        Me.Prg_UpdateDiscAmount.Name = "Prg_UpdateDiscAmount"
        Me.Prg_UpdateDiscAmount.Size = New System.Drawing.Size(582, 23)
        Me.Prg_UpdateDiscAmount.TabIndex = 52
        '
        'btnUpdateDiscAmount
        '
        Me.btnUpdateDiscAmount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateDiscAmount.Location = New System.Drawing.Point(834, 420)
        Me.btnUpdateDiscAmount.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateDiscAmount.Name = "btnUpdateDiscAmount"
        Me.btnUpdateDiscAmount.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateDiscAmount.TabIndex = 51
        Me.btnUpdateDiscAmount.TabStop = False
        Me.btnUpdateDiscAmount.Text = "Update"
        Me.btnUpdateDiscAmount.UseVisualStyleBackColor = True
        '
        'LblUpdateDiscAmount
        '
        Me.LblUpdateDiscAmount.AutoSize = True
        Me.LblUpdateDiscAmount.Location = New System.Drawing.Point(245, 451)
        Me.LblUpdateDiscAmount.Name = "LblUpdateDiscAmount"
        Me.LblUpdateDiscAmount.Size = New System.Drawing.Size(81, 17)
        Me.LblUpdateDiscAmount.TabIndex = 50
        Me.LblUpdateDiscAmount.Text = "Progress...."
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(6, 428)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(137, 17)
        Me.Label24.TabIndex = 49
        Me.Label24.Text = "Update Disc Amount"
        '
        'ProgressBar5
        '
        Me.ProgressBar5.Location = New System.Drawing.Point(245, 373)
        Me.ProgressBar5.Name = "ProgressBar5"
        Me.ProgressBar5.Size = New System.Drawing.Size(582, 23)
        Me.ProgressBar5.TabIndex = 48
        '
        'btnUpdataHR
        '
        Me.btnUpdataHR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdataHR.Location = New System.Drawing.Point(834, 371)
        Me.btnUpdataHR.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdataHR.Name = "btnUpdataHR"
        Me.btnUpdataHR.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdataHR.TabIndex = 47
        Me.btnUpdataHR.TabStop = False
        Me.btnUpdataHR.Text = "Update"
        Me.btnUpdataHR.UseVisualStyleBackColor = True
        '
        'lblUpdatev160_HR
        '
        Me.lblUpdatev160_HR.AutoSize = True
        Me.lblUpdatev160_HR.Location = New System.Drawing.Point(245, 402)
        Me.lblUpdatev160_HR.Name = "lblUpdatev160_HR"
        Me.lblUpdatev160_HR.Size = New System.Drawing.Size(81, 17)
        Me.lblUpdatev160_HR.TabIndex = 46
        Me.lblUpdatev160_HR.Text = "Progress...."
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(6, 379)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(144, 17)
        Me.Label23.TabIndex = 45
        Me.Label23.Text = "Berkaitan dengan HR"
        '
        'btnUpdatev160_TBLMasterProduct
        '
        Me.btnUpdatev160_TBLMasterProduct.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatev160_TBLMasterProduct.Location = New System.Drawing.Point(834, 122)
        Me.btnUpdatev160_TBLMasterProduct.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatev160_TBLMasterProduct.Name = "btnUpdatev160_TBLMasterProduct"
        Me.btnUpdatev160_TBLMasterProduct.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatev160_TBLMasterProduct.TabIndex = 44
        Me.btnUpdatev160_TBLMasterProduct.TabStop = False
        Me.btnUpdatev160_TBLMasterProduct.Text = "Update"
        Me.btnUpdatev160_TBLMasterProduct.UseVisualStyleBackColor = True
        '
        'LblUpdatev160_TBLMasterProduct
        '
        Me.LblUpdatev160_TBLMasterProduct.AutoSize = True
        Me.LblUpdatev160_TBLMasterProduct.Location = New System.Drawing.Point(245, 128)
        Me.LblUpdatev160_TBLMasterProduct.Name = "LblUpdatev160_TBLMasterProduct"
        Me.LblUpdatev160_TBLMasterProduct.Size = New System.Drawing.Size(81, 17)
        Me.LblUpdatev160_TBLMasterProduct.TabIndex = 43
        Me.LblUpdatev160_TBLMasterProduct.Text = "Progress...."
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(6, 130)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(153, 17)
        Me.Label19.TabIndex = 42
        Me.Label19.Text = "Update master product"
        '
        'prgUpdatev160_TambahTabelBaru
        '
        Me.prgUpdatev160_TambahTabelBaru.Location = New System.Drawing.Point(245, 177)
        Me.prgUpdatev160_TambahTabelBaru.Name = "prgUpdatev160_TambahTabelBaru"
        Me.prgUpdatev160_TambahTabelBaru.Size = New System.Drawing.Size(582, 23)
        Me.prgUpdatev160_TambahTabelBaru.TabIndex = 41
        '
        'btnUpdatev160_TambahTabelBaru
        '
        Me.btnUpdatev160_TambahTabelBaru.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatev160_TambahTabelBaru.Location = New System.Drawing.Point(834, 175)
        Me.btnUpdatev160_TambahTabelBaru.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatev160_TambahTabelBaru.Name = "btnUpdatev160_TambahTabelBaru"
        Me.btnUpdatev160_TambahTabelBaru.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatev160_TambahTabelBaru.TabIndex = 40
        Me.btnUpdatev160_TambahTabelBaru.TabStop = False
        Me.btnUpdatev160_TambahTabelBaru.Text = "Update"
        Me.btnUpdatev160_TambahTabelBaru.UseVisualStyleBackColor = True
        '
        'lblUpdatev160_TambahTabelBaru
        '
        Me.lblUpdatev160_TambahTabelBaru.AutoSize = True
        Me.lblUpdatev160_TambahTabelBaru.Location = New System.Drawing.Point(245, 206)
        Me.lblUpdatev160_TambahTabelBaru.Name = "lblUpdatev160_TambahTabelBaru"
        Me.lblUpdatev160_TambahTabelBaru.Size = New System.Drawing.Size(81, 17)
        Me.lblUpdatev160_TambahTabelBaru.TabIndex = 39
        Me.lblUpdatev160_TambahTabelBaru.Text = "Progress...."
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(6, 183)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(151, 17)
        Me.Label18.TabIndex = 38
        Me.Label18.Text = "Tambahkan tabel baru"
        '
        'prgIdentifierPO
        '
        Me.prgIdentifierPO.Location = New System.Drawing.Point(245, 324)
        Me.prgIdentifierPO.Name = "prgIdentifierPO"
        Me.prgIdentifierPO.Size = New System.Drawing.Size(582, 23)
        Me.prgIdentifierPO.TabIndex = 37
        '
        'btnUpdateIdentifierPO
        '
        Me.btnUpdateIdentifierPO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateIdentifierPO.Location = New System.Drawing.Point(834, 322)
        Me.btnUpdateIdentifierPO.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateIdentifierPO.Name = "btnUpdateIdentifierPO"
        Me.btnUpdateIdentifierPO.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateIdentifierPO.TabIndex = 36
        Me.btnUpdateIdentifierPO.TabStop = False
        Me.btnUpdateIdentifierPO.Text = "Update"
        Me.btnUpdateIdentifierPO.UseVisualStyleBackColor = True
        '
        'lblProgressPO
        '
        Me.lblProgressPO.AutoSize = True
        Me.lblProgressPO.Location = New System.Drawing.Point(245, 353)
        Me.lblProgressPO.Name = "lblProgressPO"
        Me.lblProgressPO.Size = New System.Drawing.Size(81, 17)
        Me.lblProgressPO.TabIndex = 35
        Me.lblProgressPO.Text = "Progress...."
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(6, 330)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(136, 17)
        Me.Label17.TabIndex = 34
        Me.Label17.Text = "Update Identifier PO"
        '
        'prgIdentifierSO
        '
        Me.prgIdentifierSO.Location = New System.Drawing.Point(245, 275)
        Me.prgIdentifierSO.Name = "prgIdentifierSO"
        Me.prgIdentifierSO.Size = New System.Drawing.Size(582, 23)
        Me.prgIdentifierSO.TabIndex = 33
        '
        'btnUpdate_IdentifierSO
        '
        Me.btnUpdate_IdentifierSO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate_IdentifierSO.Location = New System.Drawing.Point(834, 273)
        Me.btnUpdate_IdentifierSO.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate_IdentifierSO.Name = "btnUpdate_IdentifierSO"
        Me.btnUpdate_IdentifierSO.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdate_IdentifierSO.TabIndex = 32
        Me.btnUpdate_IdentifierSO.TabStop = False
        Me.btnUpdate_IdentifierSO.Text = "Update"
        Me.btnUpdate_IdentifierSO.UseVisualStyleBackColor = True
        '
        'LblIdentifierSO
        '
        Me.LblIdentifierSO.AutoSize = True
        Me.LblIdentifierSO.Location = New System.Drawing.Point(245, 304)
        Me.LblIdentifierSO.Name = "LblIdentifierSO"
        Me.LblIdentifierSO.Size = New System.Drawing.Size(81, 17)
        Me.LblIdentifierSO.TabIndex = 31
        Me.LblIdentifierSO.Text = "Progress...."
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(6, 281)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(136, 17)
        Me.Label16.TabIndex = 30
        Me.Label16.Text = "Update Identifier SO"
        '
        'prgUpdateSalesPrice
        '
        Me.prgUpdateSalesPrice.Location = New System.Drawing.Point(245, 226)
        Me.prgUpdateSalesPrice.Name = "prgUpdateSalesPrice"
        Me.prgUpdateSalesPrice.Size = New System.Drawing.Size(582, 23)
        Me.prgUpdateSalesPrice.TabIndex = 29
        '
        'btnPindahkanSalesPrice
        '
        Me.btnPindahkanSalesPrice.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPindahkanSalesPrice.Location = New System.Drawing.Point(834, 224)
        Me.btnPindahkanSalesPrice.Margin = New System.Windows.Forms.Padding(4)
        Me.btnPindahkanSalesPrice.Name = "btnPindahkanSalesPrice"
        Me.btnPindahkanSalesPrice.Size = New System.Drawing.Size(82, 28)
        Me.btnPindahkanSalesPrice.TabIndex = 28
        Me.btnPindahkanSalesPrice.TabStop = False
        Me.btnPindahkanSalesPrice.Text = "Update"
        Me.btnPindahkanSalesPrice.UseVisualStyleBackColor = True
        '
        'LblProgres_PindahSalesPrice
        '
        Me.LblProgres_PindahSalesPrice.AutoSize = True
        Me.LblProgres_PindahSalesPrice.Location = New System.Drawing.Point(245, 255)
        Me.LblProgres_PindahSalesPrice.Name = "LblProgres_PindahSalesPrice"
        Me.LblProgres_PindahSalesPrice.Size = New System.Drawing.Size(81, 17)
        Me.LblProgres_PindahSalesPrice.TabIndex = 27
        Me.LblProgres_PindahSalesPrice.Text = "Progress...."
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(6, 232)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(179, 17)
        Me.Label15.TabIndex = 26
        Me.Label15.Text = "Pindahkan data sales price"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(6, 154)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(885, 17)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "_________________________________________________________________________________" & _
            "__________________  Sudah CMSI"
        '
        'btnUpdatev160_UpdateNumberingMP
        '
        Me.btnUpdatev160_UpdateNumberingMP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatev160_UpdateNumberingMP.Location = New System.Drawing.Point(834, 86)
        Me.btnUpdatev160_UpdateNumberingMP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatev160_UpdateNumberingMP.Name = "btnUpdatev160_UpdateNumberingMP"
        Me.btnUpdatev160_UpdateNumberingMP.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatev160_UpdateNumberingMP.TabIndex = 24
        Me.btnUpdatev160_UpdateNumberingMP.TabStop = False
        Me.btnUpdatev160_UpdateNumberingMP.Text = "Update"
        Me.btnUpdatev160_UpdateNumberingMP.UseVisualStyleBackColor = True
        '
        'LblProgress_UpdateNumberingMaterialPlanning
        '
        Me.LblProgress_UpdateNumberingMaterialPlanning.AutoSize = True
        Me.LblProgress_UpdateNumberingMaterialPlanning.Location = New System.Drawing.Point(245, 94)
        Me.LblProgress_UpdateNumberingMaterialPlanning.Name = "LblProgress_UpdateNumberingMaterialPlanning"
        Me.LblProgress_UpdateNumberingMaterialPlanning.Size = New System.Drawing.Size(81, 17)
        Me.LblProgress_UpdateNumberingMaterialPlanning.TabIndex = 23
        Me.LblProgress_UpdateNumberingMaterialPlanning.Text = "Progress...."
        '
        'LblProgress_UpdateNumbering
        '
        Me.LblProgress_UpdateNumbering.AutoSize = True
        Me.LblProgress_UpdateNumbering.Location = New System.Drawing.Point(6, 94)
        Me.LblProgress_UpdateNumbering.Name = "LblProgress_UpdateNumbering"
        Me.LblProgress_UpdateNumbering.Size = New System.Drawing.Size(240, 17)
        Me.LblProgress_UpdateNumbering.TabIndex = 22
        Me.LblProgress_UpdateNumbering.Text = "Update Numbering Material Planning"
        '
        'btnUpdatev160_UpdateNumberingCL
        '
        Me.btnUpdatev160_UpdateNumberingCL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatev160_UpdateNumberingCL.Location = New System.Drawing.Point(834, 50)
        Me.btnUpdatev160_UpdateNumberingCL.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatev160_UpdateNumberingCL.Name = "btnUpdatev160_UpdateNumberingCL"
        Me.btnUpdatev160_UpdateNumberingCL.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatev160_UpdateNumberingCL.TabIndex = 21
        Me.btnUpdatev160_UpdateNumberingCL.TabStop = False
        Me.btnUpdatev160_UpdateNumberingCL.Text = "Update"
        Me.btnUpdatev160_UpdateNumberingCL.UseVisualStyleBackColor = True
        '
        'LblProgress_UpdateNumberingColletionLetter
        '
        Me.LblProgress_UpdateNumberingColletionLetter.AutoSize = True
        Me.LblProgress_UpdateNumberingColletionLetter.Location = New System.Drawing.Point(245, 58)
        Me.LblProgress_UpdateNumberingColletionLetter.Name = "LblProgress_UpdateNumberingColletionLetter"
        Me.LblProgress_UpdateNumberingColletionLetter.Size = New System.Drawing.Size(81, 17)
        Me.LblProgress_UpdateNumberingColletionLetter.TabIndex = 20
        Me.LblProgress_UpdateNumberingColletionLetter.Text = "Progress...."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 58)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(233, 17)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Update Numbering Collection Letter"
        '
        'btnUpdatev160_SC
        '
        Me.btnUpdatev160_SC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatev160_SC.Location = New System.Drawing.Point(834, 4)
        Me.btnUpdatev160_SC.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatev160_SC.Name = "btnUpdatev160_SC"
        Me.btnUpdatev160_SC.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatev160_SC.TabIndex = 17
        Me.btnUpdatev160_SC.TabStop = False
        Me.btnUpdatev160_SC.Text = "Update"
        Me.btnUpdatev160_SC.UseVisualStyleBackColor = True
        '
        'lblPrgStockCard
        '
        Me.lblPrgStockCard.AutoSize = True
        Me.lblPrgStockCard.Location = New System.Drawing.Point(245, 32)
        Me.lblPrgStockCard.Name = "lblPrgStockCard"
        Me.lblPrgStockCard.Size = New System.Drawing.Size(81, 17)
        Me.lblPrgStockCard.TabIndex = 2
        Me.lblPrgStockCard.Text = "Progress...."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(161, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Update Stock Card Data"
        '
        'prgStockCard
        '
        Me.prgStockCard.Location = New System.Drawing.Point(245, 6)
        Me.prgStockCard.Name = "prgStockCard"
        Me.prgStockCard.Size = New System.Drawing.Size(582, 23)
        Me.prgStockCard.TabIndex = 0
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.btnUpdatePresentasiPayroll)
        Me.TabPage4.Controls.Add(Me.Label83)
        Me.TabPage4.Controls.Add(Me.Label84)
        Me.TabPage4.Controls.Add(Me.ProgressBar18)
        Me.TabPage4.Controls.Add(Me.btnUpdateSalaryOverride)
        Me.TabPage4.Controls.Add(Me.Label85)
        Me.TabPage4.Controls.Add(Me.Label86)
        Me.TabPage4.Controls.Add(Me.ProgressBar19)
        Me.TabPage4.Controls.Add(Me.btnUpdateGetVersion)
        Me.TabPage4.Controls.Add(Me.btnUpdateSetVersion)
        Me.TabPage4.Controls.Add(Me.Label82)
        Me.TabPage4.Controls.Add(Me.Button21)
        Me.TabPage4.Controls.Add(Me.Label78)
        Me.TabPage4.Controls.Add(Me.Label79)
        Me.TabPage4.Controls.Add(Me.ProgressBar16)
        Me.TabPage4.Controls.Add(Me.Button20)
        Me.TabPage4.Controls.Add(Me.Label73)
        Me.TabPage4.Controls.Add(Me.Label74)
        Me.TabPage4.Controls.Add(Me.ProgressBar15)
        Me.TabPage4.Controls.Add(Me.btnUpdate_ChkInv)
        Me.TabPage4.Controls.Add(Me.lblProgress_ChkInv)
        Me.TabPage4.Controls.Add(Me.Label71)
        Me.TabPage4.Controls.Add(Me.PrgUpdate_ChkInv)
        Me.TabPage4.Controls.Add(Me.btnUpdate_UpdatePrOB)
        Me.TabPage4.Controls.Add(Me.lblProgress_UpdatePrOB)
        Me.TabPage4.Controls.Add(Me.Label27)
        Me.TabPage4.Controls.Add(Me.PrgUpdate_PrOB)
        Me.TabPage4.Controls.Add(Me.btnUpdate_UpdatePOB)
        Me.TabPage4.Controls.Add(Me.lblProgress_UpdatePOB)
        Me.TabPage4.Controls.Add(Me.Label30)
        Me.TabPage4.Controls.Add(Me.PrgUpdate_UpdatePOB)
        Me.TabPage4.Controls.Add(Me.btnUpdate_NPWP_AOP)
        Me.TabPage4.Controls.Add(Me.lblProgress_NPWP_AOP)
        Me.TabPage4.Controls.Add(Me.Label25)
        Me.TabPage4.Controls.Add(Me.PrgUpdate_NPWP_AOP)
        Me.TabPage4.Controls.Add(Me.btnUpdateEmail)
        Me.TabPage4.Controls.Add(Me.LblUpdateEmail)
        Me.TabPage4.Controls.Add(Me.Label22)
        Me.TabPage4.Controls.Add(Me.prgUpdateEmail)
        Me.TabPage4.Controls.Add(Me.btnDeletePackingList)
        Me.TabPage4.Controls.Add(Me.lblPackingList)
        Me.TabPage4.Controls.Add(Me.Label21)
        Me.TabPage4.Controls.Add(Me.prgPackingList)
        Me.TabPage4.Controls.Add(Me.btnUpdateSalesPerson)
        Me.TabPage4.Controls.Add(Me.LblUpdateSalesPerson)
        Me.TabPage4.Controls.Add(Me.Label20)
        Me.TabPage4.Controls.Add(Me.prbUpdateSalesPerson)
        Me.TabPage4.Location = New System.Drawing.Point(4, 25)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(923, 603)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Others"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'btnUpdatePresentasiPayroll
        '
        Me.btnUpdatePresentasiPayroll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdatePresentasiPayroll.Location = New System.Drawing.Point(833, 500)
        Me.btnUpdatePresentasiPayroll.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdatePresentasiPayroll.Name = "btnUpdatePresentasiPayroll"
        Me.btnUpdatePresentasiPayroll.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdatePresentasiPayroll.TabIndex = 69
        Me.btnUpdatePresentasiPayroll.TabStop = False
        Me.btnUpdatePresentasiPayroll.Text = "Update"
        Me.btnUpdatePresentasiPayroll.UseVisualStyleBackColor = True
        '
        'Label83
        '
        Me.Label83.AutoSize = True
        Me.Label83.Location = New System.Drawing.Point(258, 528)
        Me.Label83.Name = "Label83"
        Me.Label83.Size = New System.Drawing.Size(81, 17)
        Me.Label83.TabIndex = 68
        Me.Label83.Text = "Progress...."
        '
        'Label84
        '
        Me.Label84.AutoSize = True
        Me.Label84.Location = New System.Drawing.Point(19, 508)
        Me.Label84.Name = "Label84"
        Me.Label84.Size = New System.Drawing.Size(223, 17)
        Me.Label84.TabIndex = 67
        Me.Label84.Text = "Update Presentasi Payroll SMESS"
        '
        'ProgressBar18
        '
        Me.ProgressBar18.Location = New System.Drawing.Point(258, 502)
        Me.ProgressBar18.Name = "ProgressBar18"
        Me.ProgressBar18.Size = New System.Drawing.Size(568, 23)
        Me.ProgressBar18.TabIndex = 66
        '
        'btnUpdateSalaryOverride
        '
        Me.btnUpdateSalaryOverride.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateSalaryOverride.Location = New System.Drawing.Point(833, 454)
        Me.btnUpdateSalaryOverride.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateSalaryOverride.Name = "btnUpdateSalaryOverride"
        Me.btnUpdateSalaryOverride.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateSalaryOverride.TabIndex = 65
        Me.btnUpdateSalaryOverride.TabStop = False
        Me.btnUpdateSalaryOverride.Text = "Update"
        Me.btnUpdateSalaryOverride.UseVisualStyleBackColor = True
        '
        'Label85
        '
        Me.Label85.AutoSize = True
        Me.Label85.Location = New System.Drawing.Point(258, 482)
        Me.Label85.Name = "Label85"
        Me.Label85.Size = New System.Drawing.Size(81, 17)
        Me.Label85.TabIndex = 64
        Me.Label85.Text = "Progress...."
        '
        'Label86
        '
        Me.Label86.AutoSize = True
        Me.Label86.Location = New System.Drawing.Point(19, 462)
        Me.Label86.Name = "Label86"
        Me.Label86.Size = New System.Drawing.Size(212, 17)
        Me.Label86.TabIndex = 63
        Me.Label86.Text = "Update Override/Other Payment"
        '
        'ProgressBar19
        '
        Me.ProgressBar19.Location = New System.Drawing.Point(258, 456)
        Me.ProgressBar19.Name = "ProgressBar19"
        Me.ProgressBar19.Size = New System.Drawing.Size(568, 23)
        Me.ProgressBar19.TabIndex = 62
        '
        'btnUpdateGetVersion
        '
        Me.btnUpdateGetVersion.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateGetVersion.Location = New System.Drawing.Point(454, 421)
        Me.btnUpdateGetVersion.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateGetVersion.Name = "btnUpdateGetVersion"
        Me.btnUpdateGetVersion.Size = New System.Drawing.Size(188, 28)
        Me.btnUpdateGetVersion.TabIndex = 61
        Me.btnUpdateGetVersion.TabStop = False
        Me.btnUpdateGetVersion.Text = "Copy From My.Settings"
        Me.btnUpdateGetVersion.UseVisualStyleBackColor = True
        '
        'btnUpdateSetVersion
        '
        Me.btnUpdateSetVersion.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateSetVersion.Location = New System.Drawing.Point(258, 421)
        Me.btnUpdateSetVersion.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateSetVersion.Name = "btnUpdateSetVersion"
        Me.btnUpdateSetVersion.Size = New System.Drawing.Size(188, 28)
        Me.btnUpdateSetVersion.TabIndex = 60
        Me.btnUpdateSetVersion.TabStop = False
        Me.btnUpdateSetVersion.Text = "Copy To My.Settings"
        Me.btnUpdateSetVersion.UseVisualStyleBackColor = True
        '
        'Label82
        '
        Me.Label82.AutoSize = True
        Me.Label82.Location = New System.Drawing.Point(19, 429)
        Me.Label82.Name = "Label82"
        Me.Label82.Size = New System.Drawing.Size(173, 17)
        Me.Label82.TabIndex = 59
        Me.Label82.Text = "Set Version + Update Link"
        '
        'Button21
        '
        Me.Button21.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button21.Location = New System.Drawing.Point(833, 372)
        Me.Button21.Margin = New System.Windows.Forms.Padding(4)
        Me.Button21.Name = "Button21"
        Me.Button21.Size = New System.Drawing.Size(82, 28)
        Me.Button21.TabIndex = 57
        Me.Button21.TabStop = False
        Me.Button21.Text = "Update"
        Me.Button21.UseVisualStyleBackColor = True
        '
        'Label78
        '
        Me.Label78.AutoSize = True
        Me.Label78.Location = New System.Drawing.Point(258, 400)
        Me.Label78.Name = "Label78"
        Me.Label78.Size = New System.Drawing.Size(81, 17)
        Me.Label78.TabIndex = 56
        Me.Label78.Text = "Progress...."
        '
        'Label79
        '
        Me.Label79.AutoSize = True
        Me.Label79.Location = New System.Drawing.Point(5, 380)
        Me.Label79.Name = "Label79"
        Me.Label79.Size = New System.Drawing.Size(180, 17)
        Me.Label79.TabIndex = 55
        Me.Label79.Text = "Cek APAR Card vs Balance"
        '
        'ProgressBar16
        '
        Me.ProgressBar16.Location = New System.Drawing.Point(258, 374)
        Me.ProgressBar16.Name = "ProgressBar16"
        Me.ProgressBar16.Size = New System.Drawing.Size(568, 23)
        Me.ProgressBar16.TabIndex = 54
        '
        'Button20
        '
        Me.Button20.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button20.Location = New System.Drawing.Point(833, 326)
        Me.Button20.Margin = New System.Windows.Forms.Padding(4)
        Me.Button20.Name = "Button20"
        Me.Button20.Size = New System.Drawing.Size(82, 28)
        Me.Button20.TabIndex = 53
        Me.Button20.TabStop = False
        Me.Button20.Text = "Update"
        Me.Button20.UseVisualStyleBackColor = True
        '
        'Label73
        '
        Me.Label73.AutoSize = True
        Me.Label73.Location = New System.Drawing.Point(258, 354)
        Me.Label73.Name = "Label73"
        Me.Label73.Size = New System.Drawing.Size(81, 17)
        Me.Label73.TabIndex = 52
        Me.Label73.Text = "Progress...."
        '
        'Label74
        '
        Me.Label74.AutoSize = True
        Me.Label74.Location = New System.Drawing.Point(5, 334)
        Me.Label74.Name = "Label74"
        Me.Label74.Size = New System.Drawing.Size(163, 17)
        Me.Label74.TabIndex = 51
        Me.Label74.Text = "Re-Order Schedule LMB"
        '
        'ProgressBar15
        '
        Me.ProgressBar15.Location = New System.Drawing.Point(258, 328)
        Me.ProgressBar15.Name = "ProgressBar15"
        Me.ProgressBar15.Size = New System.Drawing.Size(568, 23)
        Me.ProgressBar15.TabIndex = 50
        '
        'btnUpdate_ChkInv
        '
        Me.btnUpdate_ChkInv.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate_ChkInv.Location = New System.Drawing.Point(833, 280)
        Me.btnUpdate_ChkInv.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate_ChkInv.Name = "btnUpdate_ChkInv"
        Me.btnUpdate_ChkInv.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdate_ChkInv.TabIndex = 49
        Me.btnUpdate_ChkInv.TabStop = False
        Me.btnUpdate_ChkInv.Text = "Update"
        Me.btnUpdate_ChkInv.UseVisualStyleBackColor = True
        '
        'lblProgress_ChkInv
        '
        Me.lblProgress_ChkInv.AutoSize = True
        Me.lblProgress_ChkInv.Location = New System.Drawing.Point(258, 308)
        Me.lblProgress_ChkInv.Name = "lblProgress_ChkInv"
        Me.lblProgress_ChkInv.Size = New System.Drawing.Size(81, 17)
        Me.lblProgress_ChkInv.TabIndex = 48
        Me.lblProgress_ChkInv.Text = "Progress...."
        '
        'Label71
        '
        Me.Label71.AutoSize = True
        Me.Label71.Location = New System.Drawing.Point(5, 288)
        Me.Label71.Name = "Label71"
        Me.Label71.Size = New System.Drawing.Size(247, 17)
        Me.Label71.TabIndex = 47
        Me.Label71.Text = "Check Invoice vs Balance vs Kartu AR"
        '
        'PrgUpdate_ChkInv
        '
        Me.PrgUpdate_ChkInv.Location = New System.Drawing.Point(258, 282)
        Me.PrgUpdate_ChkInv.Name = "PrgUpdate_ChkInv"
        Me.PrgUpdate_ChkInv.Size = New System.Drawing.Size(568, 23)
        Me.PrgUpdate_ChkInv.TabIndex = 46
        '
        'btnUpdate_UpdatePrOB
        '
        Me.btnUpdate_UpdatePrOB.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate_UpdatePrOB.Location = New System.Drawing.Point(833, 234)
        Me.btnUpdate_UpdatePrOB.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate_UpdatePrOB.Name = "btnUpdate_UpdatePrOB"
        Me.btnUpdate_UpdatePrOB.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdate_UpdatePrOB.TabIndex = 45
        Me.btnUpdate_UpdatePrOB.TabStop = False
        Me.btnUpdate_UpdatePrOB.Text = "Update"
        Me.btnUpdate_UpdatePrOB.UseVisualStyleBackColor = True
        '
        'lblProgress_UpdatePrOB
        '
        Me.lblProgress_UpdatePrOB.AutoSize = True
        Me.lblProgress_UpdatePrOB.Location = New System.Drawing.Point(258, 262)
        Me.lblProgress_UpdatePrOB.Name = "lblProgress_UpdatePrOB"
        Me.lblProgress_UpdatePrOB.Size = New System.Drawing.Size(81, 17)
        Me.lblProgress_UpdatePrOB.TabIndex = 44
        Me.lblProgress_UpdatePrOB.Text = "Progress...."
        '
        'Label27
        '
        Me.Label27.AutoSize = True
        Me.Label27.Location = New System.Drawing.Point(5, 242)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(92, 17)
        Me.Label27.TabIndex = 43
        Me.Label27.Text = "Update PrOB"
        '
        'PrgUpdate_PrOB
        '
        Me.PrgUpdate_PrOB.Location = New System.Drawing.Point(258, 236)
        Me.PrgUpdate_PrOB.Name = "PrgUpdate_PrOB"
        Me.PrgUpdate_PrOB.Size = New System.Drawing.Size(568, 23)
        Me.PrgUpdate_PrOB.TabIndex = 42
        '
        'btnUpdate_UpdatePOB
        '
        Me.btnUpdate_UpdatePOB.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate_UpdatePOB.Location = New System.Drawing.Point(833, 188)
        Me.btnUpdate_UpdatePOB.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate_UpdatePOB.Name = "btnUpdate_UpdatePOB"
        Me.btnUpdate_UpdatePOB.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdate_UpdatePOB.TabIndex = 37
        Me.btnUpdate_UpdatePOB.TabStop = False
        Me.btnUpdate_UpdatePOB.Text = "Update"
        Me.btnUpdate_UpdatePOB.UseVisualStyleBackColor = True
        '
        'lblProgress_UpdatePOB
        '
        Me.lblProgress_UpdatePOB.AutoSize = True
        Me.lblProgress_UpdatePOB.Location = New System.Drawing.Point(258, 216)
        Me.lblProgress_UpdatePOB.Name = "lblProgress_UpdatePOB"
        Me.lblProgress_UpdatePOB.Size = New System.Drawing.Size(81, 17)
        Me.lblProgress_UpdatePOB.TabIndex = 36
        Me.lblProgress_UpdatePOB.Text = "Progress...."
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.Location = New System.Drawing.Point(5, 196)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(87, 17)
        Me.Label30.TabIndex = 35
        Me.Label30.Text = "Update POB"
        '
        'PrgUpdate_UpdatePOB
        '
        Me.PrgUpdate_UpdatePOB.Location = New System.Drawing.Point(258, 190)
        Me.PrgUpdate_UpdatePOB.Name = "PrgUpdate_UpdatePOB"
        Me.PrgUpdate_UpdatePOB.Size = New System.Drawing.Size(568, 23)
        Me.PrgUpdate_UpdatePOB.TabIndex = 34
        '
        'btnUpdate_NPWP_AOP
        '
        Me.btnUpdate_NPWP_AOP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate_NPWP_AOP.Location = New System.Drawing.Point(833, 142)
        Me.btnUpdate_NPWP_AOP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate_NPWP_AOP.Name = "btnUpdate_NPWP_AOP"
        Me.btnUpdate_NPWP_AOP.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdate_NPWP_AOP.TabIndex = 33
        Me.btnUpdate_NPWP_AOP.TabStop = False
        Me.btnUpdate_NPWP_AOP.Text = "Update"
        Me.btnUpdate_NPWP_AOP.UseVisualStyleBackColor = True
        '
        'lblProgress_NPWP_AOP
        '
        Me.lblProgress_NPWP_AOP.AutoSize = True
        Me.lblProgress_NPWP_AOP.Location = New System.Drawing.Point(258, 170)
        Me.lblProgress_NPWP_AOP.Name = "lblProgress_NPWP_AOP"
        Me.lblProgress_NPWP_AOP.Size = New System.Drawing.Size(81, 17)
        Me.lblProgress_NPWP_AOP.TabIndex = 32
        Me.lblProgress_NPWP_AOP.Text = "Progress...."
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(5, 150)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(132, 17)
        Me.Label25.TabIndex = 31
        Me.Label25.Text = "Update NPWP AOP"
        '
        'PrgUpdate_NPWP_AOP
        '
        Me.PrgUpdate_NPWP_AOP.Location = New System.Drawing.Point(258, 144)
        Me.PrgUpdate_NPWP_AOP.Name = "PrgUpdate_NPWP_AOP"
        Me.PrgUpdate_NPWP_AOP.Size = New System.Drawing.Size(568, 23)
        Me.PrgUpdate_NPWP_AOP.TabIndex = 30
        '
        'btnUpdateEmail
        '
        Me.btnUpdateEmail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateEmail.Location = New System.Drawing.Point(833, 96)
        Me.btnUpdateEmail.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateEmail.Name = "btnUpdateEmail"
        Me.btnUpdateEmail.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateEmail.TabIndex = 29
        Me.btnUpdateEmail.TabStop = False
        Me.btnUpdateEmail.Text = "Update"
        Me.btnUpdateEmail.UseVisualStyleBackColor = True
        '
        'LblUpdateEmail
        '
        Me.LblUpdateEmail.AutoSize = True
        Me.LblUpdateEmail.Location = New System.Drawing.Point(258, 124)
        Me.LblUpdateEmail.Name = "LblUpdateEmail"
        Me.LblUpdateEmail.Size = New System.Drawing.Size(81, 17)
        Me.LblUpdateEmail.TabIndex = 28
        Me.LblUpdateEmail.Text = "Progress...."
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(5, 104)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(149, 17)
        Me.Label22.TabIndex = 27
        Me.Label22.Text = "Random Update Email"
        '
        'prgUpdateEmail
        '
        Me.prgUpdateEmail.Location = New System.Drawing.Point(258, 98)
        Me.prgUpdateEmail.Name = "prgUpdateEmail"
        Me.prgUpdateEmail.Size = New System.Drawing.Size(568, 23)
        Me.prgUpdateEmail.TabIndex = 26
        '
        'btnDeletePackingList
        '
        Me.btnDeletePackingList.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDeletePackingList.Location = New System.Drawing.Point(833, 50)
        Me.btnDeletePackingList.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDeletePackingList.Name = "btnDeletePackingList"
        Me.btnDeletePackingList.Size = New System.Drawing.Size(82, 28)
        Me.btnDeletePackingList.TabIndex = 25
        Me.btnDeletePackingList.TabStop = False
        Me.btnDeletePackingList.Text = "Update"
        Me.btnDeletePackingList.UseVisualStyleBackColor = True
        '
        'lblPackingList
        '
        Me.lblPackingList.AutoSize = True
        Me.lblPackingList.Location = New System.Drawing.Point(258, 78)
        Me.lblPackingList.Name = "lblPackingList"
        Me.lblPackingList.Size = New System.Drawing.Size(81, 17)
        Me.lblPackingList.TabIndex = 24
        Me.lblPackingList.Text = "Progress...."
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(5, 58)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(129, 17)
        Me.Label21.TabIndex = 23
        Me.Label21.Text = "Delete Packing List"
        '
        'prgPackingList
        '
        Me.prgPackingList.Location = New System.Drawing.Point(258, 52)
        Me.prgPackingList.Name = "prgPackingList"
        Me.prgPackingList.Size = New System.Drawing.Size(568, 23)
        Me.prgPackingList.TabIndex = 22
        '
        'btnUpdateSalesPerson
        '
        Me.btnUpdateSalesPerson.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateSalesPerson.Location = New System.Drawing.Point(833, 4)
        Me.btnUpdateSalesPerson.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdateSalesPerson.Name = "btnUpdateSalesPerson"
        Me.btnUpdateSalesPerson.Size = New System.Drawing.Size(82, 28)
        Me.btnUpdateSalesPerson.TabIndex = 21
        Me.btnUpdateSalesPerson.TabStop = False
        Me.btnUpdateSalesPerson.Text = "Update"
        Me.btnUpdateSalesPerson.UseVisualStyleBackColor = True
        '
        'LblUpdateSalesPerson
        '
        Me.LblUpdateSalesPerson.AutoSize = True
        Me.LblUpdateSalesPerson.Location = New System.Drawing.Point(258, 32)
        Me.LblUpdateSalesPerson.Name = "LblUpdateSalesPerson"
        Me.LblUpdateSalesPerson.Size = New System.Drawing.Size(81, 17)
        Me.LblUpdateSalesPerson.TabIndex = 20
        Me.LblUpdateSalesPerson.Text = "Progress...."
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(5, 12)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(199, 17)
        Me.Label20.TabIndex = 19
        Me.Label20.Text = "Random Update Sales Person"
        '
        'prbUpdateSalesPerson
        '
        Me.prbUpdateSalesPerson.Location = New System.Drawing.Point(258, 6)
        Me.prbUpdateSalesPerson.Name = "prbUpdateSalesPerson"
        Me.prbUpdateSalesPerson.Size = New System.Drawing.Size(568, 23)
        Me.prbUpdateSalesPerson.TabIndex = 18
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.prgClearTable)
        Me.TabPage5.Controls.Add(Me.LblClearTable_Progress)
        Me.TabPage5.Controls.Add(Me.Label33)
        Me.TabPage5.Controls.Add(Me.llbCheckAll)
        Me.TabPage5.Controls.Add(Me.llbUncheckAll)
        Me.TabPage5.Controls.Add(Me.TxtCT_NamaScheme)
        Me.TabPage5.Controls.Add(Me.Label26)
        Me.TabPage5.Controls.Add(Me.TxtCT_DBName)
        Me.TabPage5.Controls.Add(Me.chkCT_New)
        Me.TabPage5.Controls.Add(Me.btnCT_LoadTableName)
        Me.TabPage5.Controls.Add(Me.Label14)
        Me.TabPage5.Controls.Add(Me.TxtCT_ConnectionString)
        Me.TabPage5.Controls.Add(Me.StatusStrip1)
        Me.TabPage5.Controls.Add(Me.TxtCT_Primary)
        Me.TabPage5.Controls.Add(Me.btnCT_Delete)
        Me.TabPage5.Controls.Add(Me.btnCT_Save)
        Me.TabPage5.Controls.Add(Me.btnCT_Clear)
        Me.TabPage5.Controls.Add(Me.cboCT_SchemeName)
        Me.TabPage5.Controls.Add(Me.dgvCT_DetailTable)
        Me.TabPage5.Location = New System.Drawing.Point(4, 25)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(923, 603)
        Me.TabPage5.TabIndex = 4
        Me.TabPage5.Text = "Clear Tabel"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'prgClearTable
        '
        Me.prgClearTable.Location = New System.Drawing.Point(6, 512)
        Me.prgClearTable.Name = "prgClearTable"
        Me.prgClearTable.Size = New System.Drawing.Size(346, 25)
        Me.prgClearTable.TabIndex = 53
        '
        'LblClearTable_Progress
        '
        Me.LblClearTable_Progress.AutoSize = True
        Me.LblClearTable_Progress.Location = New System.Drawing.Point(355, 516)
        Me.LblClearTable_Progress.Name = "LblClearTable_Progress"
        Me.LblClearTable_Progress.Size = New System.Drawing.Size(59, 17)
        Me.LblClearTable_Progress.TabIndex = 52
        Me.LblClearTable_Progress.Text = "Label75"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(753, 36)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(59, 17)
        Me.Label33.TabIndex = 51
        Me.Label33.Text = "Label33"
        '
        'llbCheckAll
        '
        Me.llbCheckAll.AutoSize = True
        Me.llbCheckAll.Location = New System.Drawing.Point(760, 69)
        Me.llbCheckAll.Name = "llbCheckAll"
        Me.llbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.llbCheckAll.TabIndex = 50
        Me.llbCheckAll.TabStop = True
        Me.llbCheckAll.Text = "Check All"
        '
        'llbUncheckAll
        '
        Me.llbUncheckAll.AutoSize = True
        Me.llbUncheckAll.Location = New System.Drawing.Point(832, 69)
        Me.llbUncheckAll.Name = "llbUncheckAll"
        Me.llbUncheckAll.Size = New System.Drawing.Size(82, 17)
        Me.llbUncheckAll.TabIndex = 49
        Me.llbUncheckAll.TabStop = True
        Me.llbUncheckAll.Text = "Uncheck All"
        '
        'TxtCT_NamaScheme
        '
        Me.TxtCT_NamaScheme.Location = New System.Drawing.Point(601, 64)
        Me.TxtCT_NamaScheme.Name = "TxtCT_NamaScheme"
        Me.TxtCT_NamaScheme.Size = New System.Drawing.Size(157, 22)
        Me.TxtCT_NamaScheme.TabIndex = 48
        Me.TxtCT_NamaScheme.TabStop = False
        Me.TxtCT_NamaScheme.Visible = False
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(6, 67)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(110, 17)
        Me.Label26.TabIndex = 47
        Me.Label26.Text = "Database Name"
        '
        'TxtCT_DBName
        '
        Me.TxtCT_DBName.Location = New System.Drawing.Point(132, 64)
        Me.TxtCT_DBName.Name = "TxtCT_DBName"
        Me.TxtCT_DBName.Size = New System.Drawing.Size(463, 22)
        Me.TxtCT_DBName.TabIndex = 46
        '
        'chkCT_New
        '
        Me.chkCT_New.AutoSize = True
        Me.chkCT_New.Location = New System.Drawing.Point(439, 8)
        Me.chkCT_New.Name = "chkCT_New"
        Me.chkCT_New.Size = New System.Drawing.Size(57, 21)
        Me.chkCT_New.TabIndex = 45
        Me.chkCT_New.Text = "New"
        Me.chkCT_New.UseVisualStyleBackColor = True
        '
        'btnCT_LoadTableName
        '
        Me.btnCT_LoadTableName.Location = New System.Drawing.Point(601, 33)
        Me.btnCT_LoadTableName.Name = "btnCT_LoadTableName"
        Me.btnCT_LoadTableName.Size = New System.Drawing.Size(146, 23)
        Me.btnCT_LoadTableName.TabIndex = 44
        Me.btnCT_LoadTableName.Text = "Load Table Name"
        Me.btnCT_LoadTableName.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(6, 39)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(120, 17)
        Me.Label14.TabIndex = 43
        Me.Label14.Text = "Connection String"
        '
        'TxtCT_ConnectionString
        '
        Me.TxtCT_ConnectionString.Location = New System.Drawing.Point(132, 36)
        Me.TxtCT_ConnectionString.Name = "TxtCT_ConnectionString"
        Me.TxtCT_ConnectionString.Size = New System.Drawing.Size(463, 22)
        Me.TxtCT_ConnectionString.TabIndex = 42
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCT_CreatedBy, Me.tslCT_CreatedDate, Me.tslCT_ModifiedBy, Me.tslCT_ModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(3, 575)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(917, 25)
        Me.StatusStrip1.TabIndex = 40
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCT_CreatedBy
        '
        Me.tslCT_CreatedBy.Name = "tslCT_CreatedBy"
        Me.tslCT_CreatedBy.Size = New System.Drawing.Size(224, 20)
        Me.tslCT_CreatedBy.Spring = True
        Me.tslCT_CreatedBy.Text = "Created By: "
        '
        'tslCT_CreatedDate
        '
        Me.tslCT_CreatedDate.Name = "tslCT_CreatedDate"
        Me.tslCT_CreatedDate.Size = New System.Drawing.Size(224, 20)
        Me.tslCT_CreatedDate.Spring = True
        Me.tslCT_CreatedDate.Text = "Created Date:"
        '
        'tslCT_ModifiedBy
        '
        Me.tslCT_ModifiedBy.Name = "tslCT_ModifiedBy"
        Me.tslCT_ModifiedBy.Size = New System.Drawing.Size(224, 20)
        Me.tslCT_ModifiedBy.Spring = True
        Me.tslCT_ModifiedBy.Text = "Modified By: "
        '
        'tslCT_ModifiedDate
        '
        Me.tslCT_ModifiedDate.Name = "tslCT_ModifiedDate"
        Me.tslCT_ModifiedDate.Size = New System.Drawing.Size(224, 20)
        Me.tslCT_ModifiedDate.Spring = True
        Me.tslCT_ModifiedDate.Text = "Modified Date: "
        '
        'TxtCT_Primary
        '
        Me.TxtCT_Primary.Location = New System.Drawing.Point(99, 8)
        Me.TxtCT_Primary.Name = "TxtCT_Primary"
        Me.TxtCT_Primary.Size = New System.Drawing.Size(48, 22)
        Me.TxtCT_Primary.TabIndex = 5
        Me.TxtCT_Primary.TabStop = False
        Me.TxtCT_Primary.Text = "-1"
        Me.TxtCT_Primary.Visible = False
        '
        'btnCT_Delete
        '
        Me.btnCT_Delete.Location = New System.Drawing.Point(502, 7)
        Me.btnCT_Delete.Name = "btnCT_Delete"
        Me.btnCT_Delete.Size = New System.Drawing.Size(75, 23)
        Me.btnCT_Delete.TabIndex = 4
        Me.btnCT_Delete.Text = "Delete"
        Me.btnCT_Delete.UseVisualStyleBackColor = True
        '
        'btnCT_Save
        '
        Me.btnCT_Save.Location = New System.Drawing.Point(358, 7)
        Me.btnCT_Save.Name = "btnCT_Save"
        Me.btnCT_Save.Size = New System.Drawing.Size(75, 23)
        Me.btnCT_Save.TabIndex = 3
        Me.btnCT_Save.Text = "Save"
        Me.btnCT_Save.UseVisualStyleBackColor = True
        '
        'btnCT_Clear
        '
        Me.btnCT_Clear.Location = New System.Drawing.Point(251, 7)
        Me.btnCT_Clear.Name = "btnCT_Clear"
        Me.btnCT_Clear.Size = New System.Drawing.Size(101, 23)
        Me.btnCT_Clear.TabIndex = 2
        Me.btnCT_Clear.Text = "Clear Table"
        Me.btnCT_Clear.UseVisualStyleBackColor = True
        '
        'cboCT_SchemeName
        '
        Me.cboCT_SchemeName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCT_SchemeName.FormattingEnabled = True
        Me.cboCT_SchemeName.Location = New System.Drawing.Point(6, 6)
        Me.cboCT_SchemeName.Name = "cboCT_SchemeName"
        Me.cboCT_SchemeName.Size = New System.Drawing.Size(239, 24)
        Me.cboCT_SchemeName.TabIndex = 1
        '
        'dgvCT_DetailTable
        '
        Me.dgvCT_DetailTable.AllowUserToAddRows = False
        Me.dgvCT_DetailTable.AllowUserToDeleteRows = False
        Me.dgvCT_DetailTable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCT_DetailTable.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCT_DetailTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCT_DetailTable.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCT_No, Me.ColCT_TableName, Me.ColCT_RowCount, Me.ColCT_Clear})
        Me.dgvCT_DetailTable.Location = New System.Drawing.Point(6, 92)
        Me.dgvCT_DetailTable.Name = "dgvCT_DetailTable"
        Me.dgvCT_DetailTable.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCT_DetailTable.RowTemplate.Height = 24
        Me.dgvCT_DetailTable.Size = New System.Drawing.Size(908, 414)
        Me.dgvCT_DetailTable.TabIndex = 0
        '
        'ColCT_No
        '
        Me.ColCT_No.HeaderText = "No"
        Me.ColCT_No.Name = "ColCT_No"
        Me.ColCT_No.Width = 51
        '
        'ColCT_TableName
        '
        Me.ColCT_TableName.HeaderText = "Table's Name"
        Me.ColCT_TableName.Name = "ColCT_TableName"
        Me.ColCT_TableName.Width = 120
        '
        'ColCT_RowCount
        '
        Me.ColCT_RowCount.HeaderText = "Row Count"
        Me.ColCT_RowCount.Name = "ColCT_RowCount"
        Me.ColCT_RowCount.Width = 101
        '
        'ColCT_Clear
        '
        Me.ColCT_Clear.HeaderText = "Clear"
        Me.ColCT_Clear.Name = "ColCT_Clear"
        Me.ColCT_Clear.Width = 47
        '
        'TabPage6
        '
        Me.TabPage6.Controls.Add(Me.Label69)
        Me.TabPage6.Controls.Add(Me.Label68)
        Me.TabPage6.Controls.Add(Me.TxtLocControl_LocationDesc)
        Me.TabPage6.Controls.Add(Me.TxtLocControl_PrimaryLocation)
        Me.TabPage6.Controls.Add(Me.BtnLocControl_FindLocation)
        Me.TabPage6.Controls.Add(Me.TxtLocControl_LocationCode)
        Me.TabPage6.Controls.Add(Me.TxtLocControl_Desc)
        Me.TabPage6.Controls.Add(Me.TxtLocControl_TxtPrimaryProduct)
        Me.TabPage6.Controls.Add(Me.BtnLocControl_FindCodeNo)
        Me.TabPage6.Controls.Add(Me.TxtLocControl_CodeNo)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_stock_taking_tag)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltransaksi_stock_card)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_stock_taking_itemstart)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_stock_taking_itembuffer)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_salesreturn_item_import)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_salesreturn_item_copy)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_salesreturn_item)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_salesinvoice_item_import)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_salesinvoice_item_copy)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_salesinvoice_item)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_purchasereturn_item)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_booking_invitem)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltrans_aop_item)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltransaksi_stocktransfer_item)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltransaksi_stockbundling_item_before)
        Me.TabPage6.Controls.Add(Me.chkLocation_tbltransaksi_stockbundling_item_after)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tbltransaksi_stockinout_item)
        Me.TabPage6.Controls.Add(Me.ChkLocation_tblmaster_productqty)
        Me.TabPage6.Controls.Add(Me.prgLocationCont)
        Me.TabPage6.Controls.Add(Me.btnLocation_Start)
        Me.TabPage6.Controls.Add(Me.chkLocation_tblmaster_product)
        Me.TabPage6.Location = New System.Drawing.Point(4, 25)
        Me.TabPage6.Name = "TabPage6"
        Me.TabPage6.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage6.Size = New System.Drawing.Size(923, 603)
        Me.TabPage6.TabIndex = 5
        Me.TabPage6.Text = "Change To Location Control"
        Me.TabPage6.UseVisualStyleBackColor = True
        '
        'Label69
        '
        Me.Label69.AutoSize = True
        Me.Label69.Location = New System.Drawing.Point(6, 65)
        Me.Label69.Name = "Label69"
        Me.Label69.Size = New System.Drawing.Size(62, 17)
        Me.Label69.TabIndex = 32
        Me.Label69.Text = "Location"
        '
        'Label68
        '
        Me.Label68.AutoSize = True
        Me.Label68.Location = New System.Drawing.Point(6, 38)
        Me.Label68.Name = "Label68"
        Me.Label68.Size = New System.Drawing.Size(63, 17)
        Me.Label68.TabIndex = 31
        Me.Label68.Text = "Code No"
        '
        'TxtLocControl_LocationDesc
        '
        Me.TxtLocControl_LocationDesc.Location = New System.Drawing.Point(358, 62)
        Me.TxtLocControl_LocationDesc.Name = "TxtLocControl_LocationDesc"
        Me.TxtLocControl_LocationDesc.ReadOnly = True
        Me.TxtLocControl_LocationDesc.Size = New System.Drawing.Size(559, 22)
        Me.TxtLocControl_LocationDesc.TabIndex = 30
        '
        'TxtLocControl_PrimaryLocation
        '
        Me.TxtLocControl_PrimaryLocation.Location = New System.Drawing.Point(37, 64)
        Me.TxtLocControl_PrimaryLocation.Name = "TxtLocControl_PrimaryLocation"
        Me.TxtLocControl_PrimaryLocation.Size = New System.Drawing.Size(44, 22)
        Me.TxtLocControl_PrimaryLocation.TabIndex = 29
        Me.TxtLocControl_PrimaryLocation.TabStop = False
        Me.TxtLocControl_PrimaryLocation.Text = "-1"
        Me.TxtLocControl_PrimaryLocation.Visible = False
        '
        'BtnLocControl_FindLocation
        '
        Me.BtnLocControl_FindLocation.Location = New System.Drawing.Point(277, 62)
        Me.BtnLocControl_FindLocation.Name = "BtnLocControl_FindLocation"
        Me.BtnLocControl_FindLocation.Size = New System.Drawing.Size(75, 23)
        Me.BtnLocControl_FindLocation.TabIndex = 28
        Me.BtnLocControl_FindLocation.Text = "Find"
        Me.BtnLocControl_FindLocation.UseVisualStyleBackColor = True
        '
        'TxtLocControl_LocationCode
        '
        Me.TxtLocControl_LocationCode.Location = New System.Drawing.Point(87, 63)
        Me.TxtLocControl_LocationCode.Name = "TxtLocControl_LocationCode"
        Me.TxtLocControl_LocationCode.Size = New System.Drawing.Size(187, 22)
        Me.TxtLocControl_LocationCode.TabIndex = 27
        '
        'TxtLocControl_Desc
        '
        Me.TxtLocControl_Desc.Location = New System.Drawing.Point(358, 34)
        Me.TxtLocControl_Desc.Name = "TxtLocControl_Desc"
        Me.TxtLocControl_Desc.ReadOnly = True
        Me.TxtLocControl_Desc.Size = New System.Drawing.Size(559, 22)
        Me.TxtLocControl_Desc.TabIndex = 26
        '
        'TxtLocControl_TxtPrimaryProduct
        '
        Me.TxtLocControl_TxtPrimaryProduct.Location = New System.Drawing.Point(37, 36)
        Me.TxtLocControl_TxtPrimaryProduct.Name = "TxtLocControl_TxtPrimaryProduct"
        Me.TxtLocControl_TxtPrimaryProduct.Size = New System.Drawing.Size(44, 22)
        Me.TxtLocControl_TxtPrimaryProduct.TabIndex = 25
        Me.TxtLocControl_TxtPrimaryProduct.TabStop = False
        Me.TxtLocControl_TxtPrimaryProduct.Text = "-1"
        Me.TxtLocControl_TxtPrimaryProduct.Visible = False
        '
        'BtnLocControl_FindCodeNo
        '
        Me.BtnLocControl_FindCodeNo.Location = New System.Drawing.Point(277, 34)
        Me.BtnLocControl_FindCodeNo.Name = "BtnLocControl_FindCodeNo"
        Me.BtnLocControl_FindCodeNo.Size = New System.Drawing.Size(75, 23)
        Me.BtnLocControl_FindCodeNo.TabIndex = 23
        Me.BtnLocControl_FindCodeNo.Text = "Find"
        Me.BtnLocControl_FindCodeNo.UseVisualStyleBackColor = True
        '
        'TxtLocControl_CodeNo
        '
        Me.TxtLocControl_CodeNo.Location = New System.Drawing.Point(87, 35)
        Me.TxtLocControl_CodeNo.Name = "TxtLocControl_CodeNo"
        Me.TxtLocControl_CodeNo.Size = New System.Drawing.Size(187, 22)
        Me.TxtLocControl_CodeNo.TabIndex = 22
        '
        'ChkLocation_tbltrans_stock_taking_tag
        '
        Me.ChkLocation_tbltrans_stock_taking_tag.AutoSize = True
        Me.ChkLocation_tbltrans_stock_taking_tag.Location = New System.Drawing.Point(414, 129)
        Me.ChkLocation_tbltrans_stock_taking_tag.Name = "ChkLocation_tbltrans_stock_taking_tag"
        Me.ChkLocation_tbltrans_stock_taking_tag.Size = New System.Drawing.Size(192, 21)
        Me.ChkLocation_tbltrans_stock_taking_tag.TabIndex = 21
        Me.ChkLocation_tbltrans_stock_taking_tag.Text = "tbltrans_stock_taking_tag"
        Me.ChkLocation_tbltrans_stock_taking_tag.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltransaksi_stock_card
        '
        Me.ChkLocation_tbltransaksi_stock_card.AutoSize = True
        Me.ChkLocation_tbltransaksi_stock_card.Location = New System.Drawing.Point(414, 102)
        Me.ChkLocation_tbltransaksi_stock_card.Name = "ChkLocation_tbltransaksi_stock_card"
        Me.ChkLocation_tbltransaksi_stock_card.Size = New System.Drawing.Size(179, 21)
        Me.ChkLocation_tbltransaksi_stock_card.TabIndex = 20
        Me.ChkLocation_tbltransaksi_stock_card.Text = "tbltransaksi_stock_card"
        Me.ChkLocation_tbltransaksi_stock_card.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_stock_taking_itemstart
        '
        Me.ChkLocation_tbltrans_stock_taking_itemstart.AutoSize = True
        Me.ChkLocation_tbltrans_stock_taking_itemstart.Location = New System.Drawing.Point(414, 156)
        Me.ChkLocation_tbltrans_stock_taking_itemstart.Name = "ChkLocation_tbltrans_stock_taking_itemstart"
        Me.ChkLocation_tbltrans_stock_taking_itemstart.Size = New System.Drawing.Size(226, 21)
        Me.ChkLocation_tbltrans_stock_taking_itemstart.TabIndex = 19
        Me.ChkLocation_tbltrans_stock_taking_itemstart.Text = "tbltrans_stock_taking_itemstart"
        Me.ChkLocation_tbltrans_stock_taking_itemstart.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_stock_taking_itempenyesuaian
        '
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.AutoSize = True
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.Location = New System.Drawing.Point(6, 534)
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.Name = "ChkLocation_tbltrans_stock_taking_itempenyesuaian"
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.Size = New System.Drawing.Size(279, 21)
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.TabIndex = 18
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.Text = "tbltrans_stock_taking_itempenyesuaian"
        Me.ChkLocation_tbltrans_stock_taking_itempenyesuaian.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_stock_taking_itembuffer
        '
        Me.ChkLocation_tbltrans_stock_taking_itembuffer.AutoSize = True
        Me.ChkLocation_tbltrans_stock_taking_itembuffer.Location = New System.Drawing.Point(6, 507)
        Me.ChkLocation_tbltrans_stock_taking_itembuffer.Name = "ChkLocation_tbltrans_stock_taking_itembuffer"
        Me.ChkLocation_tbltrans_stock_taking_itembuffer.Size = New System.Drawing.Size(235, 21)
        Me.ChkLocation_tbltrans_stock_taking_itembuffer.TabIndex = 17
        Me.ChkLocation_tbltrans_stock_taking_itembuffer.Text = "tbltrans_stock_taking_itembuffer"
        Me.ChkLocation_tbltrans_stock_taking_itembuffer.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_salesreturn_item_import
        '
        Me.ChkLocation_tbltrans_salesreturn_item_import.AutoSize = True
        Me.ChkLocation_tbltrans_salesreturn_item_import.Location = New System.Drawing.Point(6, 480)
        Me.ChkLocation_tbltrans_salesreturn_item_import.Name = "ChkLocation_tbltrans_salesreturn_item_import"
        Me.ChkLocation_tbltrans_salesreturn_item_import.Size = New System.Drawing.Size(237, 21)
        Me.ChkLocation_tbltrans_salesreturn_item_import.TabIndex = 16
        Me.ChkLocation_tbltrans_salesreturn_item_import.Text = "tbltrans_salesreturn_item_import"
        Me.ChkLocation_tbltrans_salesreturn_item_import.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_salesreturn_item_copy
        '
        Me.ChkLocation_tbltrans_salesreturn_item_copy.AutoSize = True
        Me.ChkLocation_tbltrans_salesreturn_item_copy.Location = New System.Drawing.Point(6, 453)
        Me.ChkLocation_tbltrans_salesreturn_item_copy.Name = "ChkLocation_tbltrans_salesreturn_item_copy"
        Me.ChkLocation_tbltrans_salesreturn_item_copy.Size = New System.Drawing.Size(228, 21)
        Me.ChkLocation_tbltrans_salesreturn_item_copy.TabIndex = 15
        Me.ChkLocation_tbltrans_salesreturn_item_copy.Text = "tbltrans_salesreturn_item_copy"
        Me.ChkLocation_tbltrans_salesreturn_item_copy.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_salesreturn_item
        '
        Me.ChkLocation_tbltrans_salesreturn_item.AutoSize = True
        Me.ChkLocation_tbltrans_salesreturn_item.Location = New System.Drawing.Point(6, 426)
        Me.ChkLocation_tbltrans_salesreturn_item.Name = "ChkLocation_tbltrans_salesreturn_item"
        Me.ChkLocation_tbltrans_salesreturn_item.Size = New System.Drawing.Size(190, 21)
        Me.ChkLocation_tbltrans_salesreturn_item.TabIndex = 14
        Me.ChkLocation_tbltrans_salesreturn_item.Text = "tbltrans_salesreturn_item"
        Me.ChkLocation_tbltrans_salesreturn_item.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_salesinvoice_item_import
        '
        Me.ChkLocation_tbltrans_salesinvoice_item_import.AutoSize = True
        Me.ChkLocation_tbltrans_salesinvoice_item_import.Location = New System.Drawing.Point(6, 399)
        Me.ChkLocation_tbltrans_salesinvoice_item_import.Name = "ChkLocation_tbltrans_salesinvoice_item_import"
        Me.ChkLocation_tbltrans_salesinvoice_item_import.Size = New System.Drawing.Size(243, 21)
        Me.ChkLocation_tbltrans_salesinvoice_item_import.TabIndex = 13
        Me.ChkLocation_tbltrans_salesinvoice_item_import.Text = "tbltrans_salesinvoice_item_import"
        Me.ChkLocation_tbltrans_salesinvoice_item_import.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_salesinvoice_item_copy
        '
        Me.ChkLocation_tbltrans_salesinvoice_item_copy.AutoSize = True
        Me.ChkLocation_tbltrans_salesinvoice_item_copy.Location = New System.Drawing.Point(6, 372)
        Me.ChkLocation_tbltrans_salesinvoice_item_copy.Name = "ChkLocation_tbltrans_salesinvoice_item_copy"
        Me.ChkLocation_tbltrans_salesinvoice_item_copy.Size = New System.Drawing.Size(234, 21)
        Me.ChkLocation_tbltrans_salesinvoice_item_copy.TabIndex = 12
        Me.ChkLocation_tbltrans_salesinvoice_item_copy.Text = "tbltrans_salesinvoice_item_copy"
        Me.ChkLocation_tbltrans_salesinvoice_item_copy.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_salesinvoice_item
        '
        Me.ChkLocation_tbltrans_salesinvoice_item.AutoSize = True
        Me.ChkLocation_tbltrans_salesinvoice_item.Location = New System.Drawing.Point(6, 345)
        Me.ChkLocation_tbltrans_salesinvoice_item.Name = "ChkLocation_tbltrans_salesinvoice_item"
        Me.ChkLocation_tbltrans_salesinvoice_item.Size = New System.Drawing.Size(196, 21)
        Me.ChkLocation_tbltrans_salesinvoice_item.TabIndex = 11
        Me.ChkLocation_tbltrans_salesinvoice_item.Text = "tbltrans_salesinvoice_item"
        Me.ChkLocation_tbltrans_salesinvoice_item.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_purchasereturn_item
        '
        Me.ChkLocation_tbltrans_purchasereturn_item.AutoSize = True
        Me.ChkLocation_tbltrans_purchasereturn_item.Location = New System.Drawing.Point(6, 318)
        Me.ChkLocation_tbltrans_purchasereturn_item.Name = "ChkLocation_tbltrans_purchasereturn_item"
        Me.ChkLocation_tbltrans_purchasereturn_item.Size = New System.Drawing.Size(216, 21)
        Me.ChkLocation_tbltrans_purchasereturn_item.TabIndex = 10
        Me.ChkLocation_tbltrans_purchasereturn_item.Text = "tbltrans_purchasereturn_item"
        Me.ChkLocation_tbltrans_purchasereturn_item.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_booking_invitem
        '
        Me.ChkLocation_tbltrans_booking_invitem.AutoSize = True
        Me.ChkLocation_tbltrans_booking_invitem.Location = New System.Drawing.Point(6, 291)
        Me.ChkLocation_tbltrans_booking_invitem.Name = "ChkLocation_tbltrans_booking_invitem"
        Me.ChkLocation_tbltrans_booking_invitem.Size = New System.Drawing.Size(187, 21)
        Me.ChkLocation_tbltrans_booking_invitem.TabIndex = 9
        Me.ChkLocation_tbltrans_booking_invitem.Text = "tbltrans_booking_invitem"
        Me.ChkLocation_tbltrans_booking_invitem.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltrans_aop_item
        '
        Me.ChkLocation_tbltrans_aop_item.AutoSize = True
        Me.ChkLocation_tbltrans_aop_item.Location = New System.Drawing.Point(6, 264)
        Me.ChkLocation_tbltrans_aop_item.Name = "ChkLocation_tbltrans_aop_item"
        Me.ChkLocation_tbltrans_aop_item.Size = New System.Drawing.Size(143, 21)
        Me.ChkLocation_tbltrans_aop_item.TabIndex = 8
        Me.ChkLocation_tbltrans_aop_item.Text = "tbltrans_aop_item"
        Me.ChkLocation_tbltrans_aop_item.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltransaksi_stocktransfer_item
        '
        Me.ChkLocation_tbltransaksi_stocktransfer_item.AutoSize = True
        Me.ChkLocation_tbltransaksi_stocktransfer_item.Location = New System.Drawing.Point(6, 237)
        Me.ChkLocation_tbltransaksi_stocktransfer_item.Name = "ChkLocation_tbltransaksi_stocktransfer_item"
        Me.ChkLocation_tbltransaksi_stocktransfer_item.Size = New System.Drawing.Size(226, 21)
        Me.ChkLocation_tbltransaksi_stocktransfer_item.TabIndex = 7
        Me.ChkLocation_tbltransaksi_stocktransfer_item.Text = "tbltransaksi_stocktransfer_item"
        Me.ChkLocation_tbltransaksi_stocktransfer_item.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltransaksi_stockbundling_item_before
        '
        Me.ChkLocation_tbltransaksi_stockbundling_item_before.AutoSize = True
        Me.ChkLocation_tbltransaksi_stockbundling_item_before.Location = New System.Drawing.Point(6, 210)
        Me.ChkLocation_tbltransaksi_stockbundling_item_before.Name = "ChkLocation_tbltransaksi_stockbundling_item_before"
        Me.ChkLocation_tbltransaksi_stockbundling_item_before.Size = New System.Drawing.Size(280, 21)
        Me.ChkLocation_tbltransaksi_stockbundling_item_before.TabIndex = 6
        Me.ChkLocation_tbltransaksi_stockbundling_item_before.Text = "tbltransaksi_stockbundling_item_before"
        Me.ChkLocation_tbltransaksi_stockbundling_item_before.UseVisualStyleBackColor = True
        '
        'chkLocation_tbltransaksi_stockbundling_item_after
        '
        Me.chkLocation_tbltransaksi_stockbundling_item_after.AutoSize = True
        Me.chkLocation_tbltransaksi_stockbundling_item_after.Location = New System.Drawing.Point(6, 183)
        Me.chkLocation_tbltransaksi_stockbundling_item_after.Name = "chkLocation_tbltransaksi_stockbundling_item_after"
        Me.chkLocation_tbltransaksi_stockbundling_item_after.Size = New System.Drawing.Size(268, 21)
        Me.chkLocation_tbltransaksi_stockbundling_item_after.TabIndex = 5
        Me.chkLocation_tbltransaksi_stockbundling_item_after.Text = "tbltransaksi_stockbundling_item_after"
        Me.chkLocation_tbltransaksi_stockbundling_item_after.UseVisualStyleBackColor = True
        '
        'ChkLocation_tbltransaksi_stockinout_item
        '
        Me.ChkLocation_tbltransaksi_stockinout_item.AutoSize = True
        Me.ChkLocation_tbltransaksi_stockinout_item.Location = New System.Drawing.Point(6, 156)
        Me.ChkLocation_tbltransaksi_stockinout_item.Name = "ChkLocation_tbltransaksi_stockinout_item"
        Me.ChkLocation_tbltransaksi_stockinout_item.Size = New System.Drawing.Size(208, 21)
        Me.ChkLocation_tbltransaksi_stockinout_item.TabIndex = 4
        Me.ChkLocation_tbltransaksi_stockinout_item.Text = "tbltransaksi_stockinout_item"
        Me.ChkLocation_tbltransaksi_stockinout_item.UseVisualStyleBackColor = True
        '
        'ChkLocation_tblmaster_productqty
        '
        Me.ChkLocation_tblmaster_productqty.AutoSize = True
        Me.ChkLocation_tblmaster_productqty.Location = New System.Drawing.Point(6, 129)
        Me.ChkLocation_tblmaster_productqty.Name = "ChkLocation_tblmaster_productqty"
        Me.ChkLocation_tblmaster_productqty.Size = New System.Drawing.Size(163, 21)
        Me.ChkLocation_tblmaster_productqty.TabIndex = 3
        Me.ChkLocation_tblmaster_productqty.Text = "tblmaster_productqty"
        Me.ChkLocation_tblmaster_productqty.UseVisualStyleBackColor = True
        '
        'prgLocationCont
        '
        Me.prgLocationCont.Location = New System.Drawing.Point(87, 6)
        Me.prgLocationCont.Name = "prgLocationCont"
        Me.prgLocationCont.Size = New System.Drawing.Size(830, 23)
        Me.prgLocationCont.TabIndex = 2
        '
        'btnLocation_Start
        '
        Me.btnLocation_Start.Location = New System.Drawing.Point(6, 6)
        Me.btnLocation_Start.Name = "btnLocation_Start"
        Me.btnLocation_Start.Size = New System.Drawing.Size(75, 23)
        Me.btnLocation_Start.TabIndex = 1
        Me.btnLocation_Start.Text = "Start"
        Me.btnLocation_Start.UseVisualStyleBackColor = True
        '
        'chkLocation_tblmaster_product
        '
        Me.chkLocation_tblmaster_product.AutoSize = True
        Me.chkLocation_tblmaster_product.Location = New System.Drawing.Point(6, 102)
        Me.chkLocation_tblmaster_product.Name = "chkLocation_tblmaster_product"
        Me.chkLocation_tblmaster_product.Size = New System.Drawing.Size(144, 21)
        Me.chkLocation_tblmaster_product.TabIndex = 0
        Me.chkLocation_tblmaster_product.Text = "tblmaster_product"
        Me.chkLocation_tblmaster_product.UseVisualStyleBackColor = True
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.btn_v170_quotation)
        Me.TabPage7.Controls.Add(Me.lbl_v170_quotation)
        Me.TabPage7.Controls.Add(Me.Label28)
        Me.TabPage7.Controls.Add(Me.prg_v170_quotation)
        Me.TabPage7.Location = New System.Drawing.Point(4, 25)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage7.Size = New System.Drawing.Size(923, 603)
        Me.TabPage7.TabIndex = 6
        Me.TabPage7.Text = "v.1.7.0"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'btn_v170_quotation
        '
        Me.btn_v170_quotation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v170_quotation.Location = New System.Drawing.Point(834, 4)
        Me.btn_v170_quotation.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v170_quotation.Name = "btn_v170_quotation"
        Me.btn_v170_quotation.Size = New System.Drawing.Size(82, 28)
        Me.btn_v170_quotation.TabIndex = 21
        Me.btn_v170_quotation.TabStop = False
        Me.btn_v170_quotation.Text = "Update"
        Me.btn_v170_quotation.UseVisualStyleBackColor = True
        '
        'lbl_v170_quotation
        '
        Me.lbl_v170_quotation.AutoSize = True
        Me.lbl_v170_quotation.Location = New System.Drawing.Point(245, 32)
        Me.lbl_v170_quotation.Name = "lbl_v170_quotation"
        Me.lbl_v170_quotation.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v170_quotation.TabIndex = 20
        Me.lbl_v170_quotation.Text = "Progress...."
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.Location = New System.Drawing.Point(6, 12)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(120, 17)
        Me.Label28.TabIndex = 19
        Me.Label28.Text = "Update Quotation"
        '
        'prg_v170_quotation
        '
        Me.prg_v170_quotation.Location = New System.Drawing.Point(245, 6)
        Me.prg_v170_quotation.Name = "prg_v170_quotation"
        Me.prg_v170_quotation.Size = New System.Drawing.Size(582, 23)
        Me.prg_v170_quotation.TabIndex = 18
        '
        'TabPage8
        '
        Me.TabPage8.Controls.Add(Me.btn_v180_Subtotal)
        Me.TabPage8.Controls.Add(Me.lbl_v180_Subtotal)
        Me.TabPage8.Controls.Add(Me.Label42)
        Me.TabPage8.Controls.Add(Me.prg_v180_Subtotal)
        Me.TabPage8.Controls.Add(Me.Button13)
        Me.TabPage8.Controls.Add(Me.Label37)
        Me.TabPage8.Controls.Add(Me.Label41)
        Me.TabPage8.Controls.Add(Me.ProgressBar6)
        Me.TabPage8.Controls.Add(Me.btn_v180_DetailItemJurnal)
        Me.TabPage8.Controls.Add(Me.lbl_v180_DetailItemJurnal)
        Me.TabPage8.Controls.Add(Me.Label40)
        Me.TabPage8.Controls.Add(Me.prg_v180_DetailItemJurnal)
        Me.TabPage8.Controls.Add(Me.btn_v180_ProsesBundling)
        Me.TabPage8.Controls.Add(Me.lbl_v180_ProsesBundling)
        Me.TabPage8.Controls.Add(Me.Label39)
        Me.TabPage8.Controls.Add(Me.prg_v180_ProsesBundling)
        Me.TabPage8.Controls.Add(Me.btn_v180_DetailClosing)
        Me.TabPage8.Controls.Add(Me.lbl_v180_DetailClosing)
        Me.TabPage8.Controls.Add(Me.Label38)
        Me.TabPage8.Controls.Add(Me.prg_v180_DetailClosing)
        Me.TabPage8.Controls.Add(Me.btn_v180_Salary)
        Me.TabPage8.Controls.Add(Me.lbl_v180_Salary)
        Me.TabPage8.Controls.Add(Me.Label35)
        Me.TabPage8.Controls.Add(Me.prg_v180_Salary)
        Me.TabPage8.Controls.Add(Me.btn_v180_SC)
        Me.TabPage8.Controls.Add(Me.lbl_v180_SC)
        Me.TabPage8.Controls.Add(Me.Label36)
        Me.TabPage8.Controls.Add(Me.prg_v180_SC)
        Me.TabPage8.Controls.Add(Me.btn_v180_PR)
        Me.TabPage8.Controls.Add(Me.lbl_v180_PR)
        Me.TabPage8.Controls.Add(Me.Label34)
        Me.TabPage8.Controls.Add(Me.prg_v180_PR)
        Me.TabPage8.Controls.Add(Me.btn_v180_multiAcc)
        Me.TabPage8.Controls.Add(Me.lbl_v180_multiAcc)
        Me.TabPage8.Controls.Add(Me.Label32)
        Me.TabPage8.Controls.Add(Me.prg_v180_multiAcc)
        Me.TabPage8.Controls.Add(Me.btn_v180_master_cuti)
        Me.TabPage8.Controls.Add(Me.lbl_v180_master_cuti)
        Me.TabPage8.Controls.Add(Me.Label31)
        Me.TabPage8.Controls.Add(Me.prg_v180_master_cuti)
        Me.TabPage8.Controls.Add(Me.btn_v180_multiunit)
        Me.TabPage8.Controls.Add(Me.lbl_v180_multiunit)
        Me.TabPage8.Controls.Add(Me.Label29)
        Me.TabPage8.Controls.Add(Me.prg_v180_multiunit)
        Me.TabPage8.Location = New System.Drawing.Point(4, 25)
        Me.TabPage8.Name = "TabPage8"
        Me.TabPage8.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage8.Size = New System.Drawing.Size(923, 603)
        Me.TabPage8.TabIndex = 7
        Me.TabPage8.Text = "v.1.8.0"
        Me.TabPage8.UseVisualStyleBackColor = True
        '
        'btn_v180_Subtotal
        '
        Me.btn_v180_Subtotal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_Subtotal.Location = New System.Drawing.Point(835, 452)
        Me.btn_v180_Subtotal.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_Subtotal.Name = "btn_v180_Subtotal"
        Me.btn_v180_Subtotal.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_Subtotal.TabIndex = 89
        Me.btn_v180_Subtotal.TabStop = False
        Me.btn_v180_Subtotal.Text = "Update"
        Me.btn_v180_Subtotal.UseVisualStyleBackColor = True
        '
        'lbl_v180_Subtotal
        '
        Me.lbl_v180_Subtotal.AutoSize = True
        Me.lbl_v180_Subtotal.Location = New System.Drawing.Point(267, 481)
        Me.lbl_v180_Subtotal.Name = "lbl_v180_Subtotal"
        Me.lbl_v180_Subtotal.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_Subtotal.TabIndex = 88
        Me.lbl_v180_Subtotal.Text = "Progress...."
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(6, 462)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(230, 17)
        Me.Label42.TabIndex = 87
        Me.Label42.Text = "Update Perhitungan Subtotal Sales"
        '
        'prg_v180_Subtotal
        '
        Me.prg_v180_Subtotal.Location = New System.Drawing.Point(267, 456)
        Me.prg_v180_Subtotal.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_Subtotal.Name = "prg_v180_Subtotal"
        Me.prg_v180_Subtotal.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_Subtotal.TabIndex = 86
        '
        'Button13
        '
        Me.Button13.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button13.Location = New System.Drawing.Point(835, 408)
        Me.Button13.Margin = New System.Windows.Forms.Padding(4)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(83, 28)
        Me.Button13.TabIndex = 85
        Me.Button13.TabStop = False
        Me.Button13.Text = "Update"
        Me.Button13.UseVisualStyleBackColor = True
        '
        'Label37
        '
        Me.Label37.AutoSize = True
        Me.Label37.Location = New System.Drawing.Point(267, 437)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(81, 17)
        Me.Label37.TabIndex = 84
        Me.Label37.Text = "Progress...."
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(6, 417)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(223, 17)
        Me.Label41.TabIndex = 83
        Me.Label41.Text = "Test Buat DP Purchase dan Sales"
        '
        'ProgressBar6
        '
        Me.ProgressBar6.Location = New System.Drawing.Point(267, 411)
        Me.ProgressBar6.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar6.Name = "ProgressBar6"
        Me.ProgressBar6.Size = New System.Drawing.Size(560, 23)
        Me.ProgressBar6.TabIndex = 82
        '
        'btn_v180_DetailItemJurnal
        '
        Me.btn_v180_DetailItemJurnal.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_DetailItemJurnal.Location = New System.Drawing.Point(835, 363)
        Me.btn_v180_DetailItemJurnal.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_DetailItemJurnal.Name = "btn_v180_DetailItemJurnal"
        Me.btn_v180_DetailItemJurnal.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_DetailItemJurnal.TabIndex = 81
        Me.btn_v180_DetailItemJurnal.TabStop = False
        Me.btn_v180_DetailItemJurnal.Text = "Update"
        Me.btn_v180_DetailItemJurnal.UseVisualStyleBackColor = True
        '
        'lbl_v180_DetailItemJurnal
        '
        Me.lbl_v180_DetailItemJurnal.AutoSize = True
        Me.lbl_v180_DetailItemJurnal.Location = New System.Drawing.Point(267, 392)
        Me.lbl_v180_DetailItemJurnal.Name = "lbl_v180_DetailItemJurnal"
        Me.lbl_v180_DetailItemJurnal.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_DetailItemJurnal.TabIndex = 80
        Me.lbl_v180_DetailItemJurnal.Text = "Progress...."
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(6, 372)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(229, 17)
        Me.Label40.TabIndex = 79
        Me.Label40.Text = "Update Tipe Item Detail Item Jurnal"
        '
        'prg_v180_DetailItemJurnal
        '
        Me.prg_v180_DetailItemJurnal.Location = New System.Drawing.Point(267, 366)
        Me.prg_v180_DetailItemJurnal.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_DetailItemJurnal.Name = "prg_v180_DetailItemJurnal"
        Me.prg_v180_DetailItemJurnal.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_DetailItemJurnal.TabIndex = 78
        '
        'btn_v180_ProsesBundling
        '
        Me.btn_v180_ProsesBundling.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_ProsesBundling.Location = New System.Drawing.Point(835, 318)
        Me.btn_v180_ProsesBundling.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_ProsesBundling.Name = "btn_v180_ProsesBundling"
        Me.btn_v180_ProsesBundling.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_ProsesBundling.TabIndex = 77
        Me.btn_v180_ProsesBundling.TabStop = False
        Me.btn_v180_ProsesBundling.Text = "Update"
        Me.btn_v180_ProsesBundling.UseVisualStyleBackColor = True
        '
        'lbl_v180_ProsesBundling
        '
        Me.lbl_v180_ProsesBundling.AutoSize = True
        Me.lbl_v180_ProsesBundling.Location = New System.Drawing.Point(267, 347)
        Me.lbl_v180_ProsesBundling.Name = "lbl_v180_ProsesBundling"
        Me.lbl_v180_ProsesBundling.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_ProsesBundling.TabIndex = 76
        Me.lbl_v180_ProsesBundling.Text = "Progress...."
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(6, 327)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(181, 17)
        Me.Label39.TabIndex = 75
        Me.Label39.Text = "Update Proses pd Bundling"
        '
        'prg_v180_ProsesBundling
        '
        Me.prg_v180_ProsesBundling.Location = New System.Drawing.Point(267, 321)
        Me.prg_v180_ProsesBundling.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_ProsesBundling.Name = "prg_v180_ProsesBundling"
        Me.prg_v180_ProsesBundling.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_ProsesBundling.TabIndex = 74
        '
        'btn_v180_DetailClosing
        '
        Me.btn_v180_DetailClosing.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_DetailClosing.Location = New System.Drawing.Point(835, 273)
        Me.btn_v180_DetailClosing.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_DetailClosing.Name = "btn_v180_DetailClosing"
        Me.btn_v180_DetailClosing.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_DetailClosing.TabIndex = 73
        Me.btn_v180_DetailClosing.TabStop = False
        Me.btn_v180_DetailClosing.Text = "Update"
        Me.btn_v180_DetailClosing.UseVisualStyleBackColor = True
        '
        'lbl_v180_DetailClosing
        '
        Me.lbl_v180_DetailClosing.AutoSize = True
        Me.lbl_v180_DetailClosing.Location = New System.Drawing.Point(267, 302)
        Me.lbl_v180_DetailClosing.Name = "lbl_v180_DetailClosing"
        Me.lbl_v180_DetailClosing.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_DetailClosing.TabIndex = 72
        Me.lbl_v180_DetailClosing.Text = "Progress...."
        '
        'Label38
        '
        Me.Label38.AutoSize = True
        Me.Label38.Location = New System.Drawing.Point(6, 282)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(206, 17)
        Me.Label38.TabIndex = 71
        Me.Label38.Text = "Update Tipe Item Detail Closing"
        '
        'prg_v180_DetailClosing
        '
        Me.prg_v180_DetailClosing.Location = New System.Drawing.Point(267, 276)
        Me.prg_v180_DetailClosing.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_DetailClosing.Name = "prg_v180_DetailClosing"
        Me.prg_v180_DetailClosing.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_DetailClosing.TabIndex = 70
        '
        'btn_v180_Salary
        '
        Me.btn_v180_Salary.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_Salary.Location = New System.Drawing.Point(835, 228)
        Me.btn_v180_Salary.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_Salary.Name = "btn_v180_Salary"
        Me.btn_v180_Salary.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_Salary.TabIndex = 69
        Me.btn_v180_Salary.TabStop = False
        Me.btn_v180_Salary.Text = "Update"
        Me.btn_v180_Salary.UseVisualStyleBackColor = True
        '
        'lbl_v180_Salary
        '
        Me.lbl_v180_Salary.AutoSize = True
        Me.lbl_v180_Salary.Location = New System.Drawing.Point(267, 257)
        Me.lbl_v180_Salary.Name = "lbl_v180_Salary"
        Me.lbl_v180_Salary.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_Salary.TabIndex = 68
        Me.lbl_v180_Salary.Text = "Progress...."
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(6, 237)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(174, 17)
        Me.Label35.TabIndex = 67
        Me.Label35.Text = "Update Salary Component"
        '
        'prg_v180_Salary
        '
        Me.prg_v180_Salary.Location = New System.Drawing.Point(267, 231)
        Me.prg_v180_Salary.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_Salary.Name = "prg_v180_Salary"
        Me.prg_v180_Salary.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_Salary.TabIndex = 66
        '
        'btn_v180_SC
        '
        Me.btn_v180_SC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_SC.Location = New System.Drawing.Point(835, 184)
        Me.btn_v180_SC.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_SC.Name = "btn_v180_SC"
        Me.btn_v180_SC.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_SC.TabIndex = 65
        Me.btn_v180_SC.TabStop = False
        Me.btn_v180_SC.Text = "Update"
        Me.btn_v180_SC.UseVisualStyleBackColor = True
        '
        'lbl_v180_SC
        '
        Me.lbl_v180_SC.AutoSize = True
        Me.lbl_v180_SC.Location = New System.Drawing.Point(267, 212)
        Me.lbl_v180_SC.Name = "lbl_v180_SC"
        Me.lbl_v180_SC.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_SC.TabIndex = 64
        Me.lbl_v180_SC.Text = "Progress...."
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Location = New System.Drawing.Point(6, 192)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(161, 17)
        Me.Label36.TabIndex = 63
        Me.Label36.Text = "Update Stock Card Data"
        '
        'prg_v180_SC
        '
        Me.prg_v180_SC.Location = New System.Drawing.Point(267, 186)
        Me.prg_v180_SC.Name = "prg_v180_SC"
        Me.prg_v180_SC.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_SC.TabIndex = 62
        '
        'btn_v180_PR
        '
        Me.btn_v180_PR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_PR.Location = New System.Drawing.Point(835, 137)
        Me.btn_v180_PR.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_PR.Name = "btn_v180_PR"
        Me.btn_v180_PR.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_PR.TabIndex = 61
        Me.btn_v180_PR.TabStop = False
        Me.btn_v180_PR.Text = "Update"
        Me.btn_v180_PR.UseVisualStyleBackColor = True
        '
        'lbl_v180_PR
        '
        Me.lbl_v180_PR.AutoSize = True
        Me.lbl_v180_PR.Location = New System.Drawing.Point(267, 166)
        Me.lbl_v180_PR.Name = "lbl_v180_PR"
        Me.lbl_v180_PR.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_PR.TabIndex = 60
        Me.lbl_v180_PR.Text = "Progress...."
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(6, 146)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(175, 17)
        Me.Label34.TabIndex = 59
        Me.Label34.Text = "Update Purchase Request"
        '
        'prg_v180_PR
        '
        Me.prg_v180_PR.Location = New System.Drawing.Point(267, 140)
        Me.prg_v180_PR.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_PR.Name = "prg_v180_PR"
        Me.prg_v180_PR.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_PR.TabIndex = 58
        '
        'btn_v180_multiAcc
        '
        Me.btn_v180_multiAcc.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_multiAcc.Location = New System.Drawing.Point(835, 93)
        Me.btn_v180_multiAcc.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_multiAcc.Name = "btn_v180_multiAcc"
        Me.btn_v180_multiAcc.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_multiAcc.TabIndex = 37
        Me.btn_v180_multiAcc.TabStop = False
        Me.btn_v180_multiAcc.Text = "Update"
        Me.btn_v180_multiAcc.UseVisualStyleBackColor = True
        '
        'lbl_v180_multiAcc
        '
        Me.lbl_v180_multiAcc.AutoSize = True
        Me.lbl_v180_multiAcc.Location = New System.Drawing.Point(267, 121)
        Me.lbl_v180_multiAcc.Name = "lbl_v180_multiAcc"
        Me.lbl_v180_multiAcc.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_multiAcc.TabIndex = 36
        Me.lbl_v180_multiAcc.Text = "Progress...."
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(6, 101)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(255, 17)
        Me.Label32.TabIndex = 35
        Me.Label32.Text = "Update Multi Account + Sales Multi Unit"
        '
        'prg_v180_multiAcc
        '
        Me.prg_v180_multiAcc.Location = New System.Drawing.Point(267, 95)
        Me.prg_v180_multiAcc.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_multiAcc.Name = "prg_v180_multiAcc"
        Me.prg_v180_multiAcc.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_multiAcc.TabIndex = 34
        '
        'btn_v180_master_cuti
        '
        Me.btn_v180_master_cuti.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_master_cuti.Location = New System.Drawing.Point(835, 48)
        Me.btn_v180_master_cuti.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_master_cuti.Name = "btn_v180_master_cuti"
        Me.btn_v180_master_cuti.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_master_cuti.TabIndex = 33
        Me.btn_v180_master_cuti.TabStop = False
        Me.btn_v180_master_cuti.Text = "Update"
        Me.btn_v180_master_cuti.UseVisualStyleBackColor = True
        '
        'lbl_v180_master_cuti
        '
        Me.lbl_v180_master_cuti.AutoSize = True
        Me.lbl_v180_master_cuti.Location = New System.Drawing.Point(267, 76)
        Me.lbl_v180_master_cuti.Name = "lbl_v180_master_cuti"
        Me.lbl_v180_master_cuti.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_master_cuti.TabIndex = 32
        Me.lbl_v180_master_cuti.Text = "Progress...."
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(6, 56)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(187, 17)
        Me.Label31.TabIndex = 31
        Me.Label31.Text = "Update Identifier Master Cuti"
        '
        'prg_v180_master_cuti
        '
        Me.prg_v180_master_cuti.Location = New System.Drawing.Point(267, 50)
        Me.prg_v180_master_cuti.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_master_cuti.Name = "prg_v180_master_cuti"
        Me.prg_v180_master_cuti.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_master_cuti.TabIndex = 30
        '
        'btn_v180_multiunit
        '
        Me.btn_v180_multiunit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_multiunit.Location = New System.Drawing.Point(835, 3)
        Me.btn_v180_multiunit.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_multiunit.Name = "btn_v180_multiunit"
        Me.btn_v180_multiunit.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_multiunit.TabIndex = 29
        Me.btn_v180_multiunit.TabStop = False
        Me.btn_v180_multiunit.Text = "Update"
        Me.btn_v180_multiunit.UseVisualStyleBackColor = True
        '
        'lbl_v180_multiunit
        '
        Me.lbl_v180_multiunit.AutoSize = True
        Me.lbl_v180_multiunit.Location = New System.Drawing.Point(267, 31)
        Me.lbl_v180_multiunit.Name = "lbl_v180_multiunit"
        Me.lbl_v180_multiunit.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_multiunit.TabIndex = 28
        Me.lbl_v180_multiunit.Text = "Progress...."
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(6, 11)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(192, 17)
        Me.Label29.TabIndex = 27
        Me.Label29.Text = "Update Multi Unit Component"
        '
        'prg_v180_multiunit
        '
        Me.prg_v180_multiunit.Location = New System.Drawing.Point(267, 5)
        Me.prg_v180_multiunit.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_multiunit.Name = "prg_v180_multiunit"
        Me.prg_v180_multiunit.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_multiunit.TabIndex = 26
        '
        'TabPage9
        '
        Me.TabPage9.Controls.Add(Me.btn_v180_UpdateSalary)
        Me.TabPage9.Controls.Add(Me.lbl_v180_UpdateSalary)
        Me.TabPage9.Controls.Add(Me.Label75)
        Me.TabPage9.Controls.Add(Me.prg_v180_UpdateSalary)
        Me.TabPage9.Controls.Add(Me.btn_v180_UpdateDiscAmount1)
        Me.TabPage9.Controls.Add(Me.Label57)
        Me.TabPage9.Controls.Add(Me.Label58)
        Me.TabPage9.Controls.Add(Me.ProgressBar10)
        Me.TabPage9.Controls.Add(Me.btn_v180_UpdateJurnalItem2)
        Me.TabPage9.Controls.Add(Me.Label50)
        Me.TabPage9.Controls.Add(Me.Label51)
        Me.TabPage9.Controls.Add(Me.ProgressBar9)
        Me.TabPage9.Controls.Add(Me.btn_v180_MasterCust_TrimCode)
        Me.TabPage9.Controls.Add(Me.lbl_v180_MasterCust_TrimCode)
        Me.TabPage9.Controls.Add(Me.Label49)
        Me.TabPage9.Controls.Add(Me.prg_v180_MasterCust_TrimCode)
        Me.TabPage9.Controls.Add(Me.btn_v180_QuotationLimit)
        Me.TabPage9.Controls.Add(Me.Label46)
        Me.TabPage9.Controls.Add(Me.Label47)
        Me.TabPage9.Controls.Add(Me.ProgressBar8)
        Me.TabPage9.Controls.Add(Me.btn_v180_JournalPaymentMethod)
        Me.TabPage9.Controls.Add(Me.Label43)
        Me.TabPage9.Controls.Add(Me.Label45)
        Me.TabPage9.Controls.Add(Me.ProgressBar7)
        Me.TabPage9.Controls.Add(Me.btn_v180_BundlingDate)
        Me.TabPage9.Controls.Add(Me.lbl_v180_BundlingDate)
        Me.TabPage9.Controls.Add(Me.Label44)
        Me.TabPage9.Controls.Add(Me.prg_v180_BundlingDate)
        Me.TabPage9.Location = New System.Drawing.Point(4, 25)
        Me.TabPage9.Name = "TabPage9"
        Me.TabPage9.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage9.Size = New System.Drawing.Size(923, 603)
        Me.TabPage9.TabIndex = 8
        Me.TabPage9.Text = "v.1.8.0 - cont"
        Me.TabPage9.UseVisualStyleBackColor = True
        '
        'btn_v180_UpdateSalary
        '
        Me.btn_v180_UpdateSalary.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_UpdateSalary.Location = New System.Drawing.Point(833, 274)
        Me.btn_v180_UpdateSalary.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_UpdateSalary.Name = "btn_v180_UpdateSalary"
        Me.btn_v180_UpdateSalary.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_UpdateSalary.TabIndex = 121
        Me.btn_v180_UpdateSalary.TabStop = False
        Me.btn_v180_UpdateSalary.Text = "Update"
        Me.btn_v180_UpdateSalary.UseVisualStyleBackColor = True
        '
        'lbl_v180_UpdateSalary
        '
        Me.lbl_v180_UpdateSalary.AutoSize = True
        Me.lbl_v180_UpdateSalary.Location = New System.Drawing.Point(265, 303)
        Me.lbl_v180_UpdateSalary.Name = "lbl_v180_UpdateSalary"
        Me.lbl_v180_UpdateSalary.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_UpdateSalary.TabIndex = 120
        Me.lbl_v180_UpdateSalary.Text = "Progress...."
        '
        'Label75
        '
        Me.Label75.AutoSize = True
        Me.Label75.Location = New System.Drawing.Point(4, 283)
        Me.Label75.Name = "Label75"
        Me.Label75.Size = New System.Drawing.Size(118, 17)
        Me.Label75.TabIndex = 119
        Me.Label75.Text = "Update TblSalary"
        '
        'prg_v180_UpdateSalary
        '
        Me.prg_v180_UpdateSalary.Location = New System.Drawing.Point(317, 277)
        Me.prg_v180_UpdateSalary.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_UpdateSalary.Name = "prg_v180_UpdateSalary"
        Me.prg_v180_UpdateSalary.Size = New System.Drawing.Size(508, 23)
        Me.prg_v180_UpdateSalary.TabIndex = 118
        '
        'btn_v180_UpdateDiscAmount1
        '
        Me.btn_v180_UpdateDiscAmount1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_UpdateDiscAmount1.Location = New System.Drawing.Point(833, 229)
        Me.btn_v180_UpdateDiscAmount1.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_UpdateDiscAmount1.Name = "btn_v180_UpdateDiscAmount1"
        Me.btn_v180_UpdateDiscAmount1.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_UpdateDiscAmount1.TabIndex = 113
        Me.btn_v180_UpdateDiscAmount1.TabStop = False
        Me.btn_v180_UpdateDiscAmount1.Text = "Update"
        Me.btn_v180_UpdateDiscAmount1.UseVisualStyleBackColor = True
        '
        'Label57
        '
        Me.Label57.AutoSize = True
        Me.Label57.Location = New System.Drawing.Point(265, 258)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(81, 17)
        Me.Label57.TabIndex = 112
        Me.Label57.Text = "Progress...."
        '
        'Label58
        '
        Me.Label58.AutoSize = True
        Me.Label58.Location = New System.Drawing.Point(4, 238)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(228, 17)
        Me.Label58.TabIndex = 111
        Me.Label58.Text = "Update disc amount aop + pr quick"
        '
        'ProgressBar10
        '
        Me.ProgressBar10.Location = New System.Drawing.Point(317, 232)
        Me.ProgressBar10.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar10.Name = "ProgressBar10"
        Me.ProgressBar10.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar10.TabIndex = 110
        '
        'btn_v180_UpdateJurnalItem2
        '
        Me.btn_v180_UpdateJurnalItem2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_UpdateJurnalItem2.Location = New System.Drawing.Point(833, 184)
        Me.btn_v180_UpdateJurnalItem2.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_UpdateJurnalItem2.Name = "btn_v180_UpdateJurnalItem2"
        Me.btn_v180_UpdateJurnalItem2.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_UpdateJurnalItem2.TabIndex = 109
        Me.btn_v180_UpdateJurnalItem2.TabStop = False
        Me.btn_v180_UpdateJurnalItem2.Text = "Update"
        Me.btn_v180_UpdateJurnalItem2.UseVisualStyleBackColor = True
        '
        'Label50
        '
        Me.Label50.AutoSize = True
        Me.Label50.Location = New System.Drawing.Point(265, 213)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(81, 17)
        Me.Label50.TabIndex = 108
        Me.Label50.Text = "Progress...."
        '
        'Label51
        '
        Me.Label51.AutoSize = True
        Me.Label51.Location = New System.Drawing.Point(4, 193)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(299, 17)
        Me.Label51.TabIndex = 107
        Me.Label51.Text = "Update Tipe Item Detail Closing + Item Journal"
        '
        'ProgressBar9
        '
        Me.ProgressBar9.Location = New System.Drawing.Point(317, 187)
        Me.ProgressBar9.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar9.Name = "ProgressBar9"
        Me.ProgressBar9.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar9.TabIndex = 106
        '
        'btn_v180_MasterCust_TrimCode
        '
        Me.btn_v180_MasterCust_TrimCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_MasterCust_TrimCode.Location = New System.Drawing.Point(833, 139)
        Me.btn_v180_MasterCust_TrimCode.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_MasterCust_TrimCode.Name = "btn_v180_MasterCust_TrimCode"
        Me.btn_v180_MasterCust_TrimCode.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_MasterCust_TrimCode.TabIndex = 105
        Me.btn_v180_MasterCust_TrimCode.TabStop = False
        Me.btn_v180_MasterCust_TrimCode.Text = "Update"
        Me.btn_v180_MasterCust_TrimCode.UseVisualStyleBackColor = True
        '
        'lbl_v180_MasterCust_TrimCode
        '
        Me.lbl_v180_MasterCust_TrimCode.AutoSize = True
        Me.lbl_v180_MasterCust_TrimCode.Location = New System.Drawing.Point(265, 168)
        Me.lbl_v180_MasterCust_TrimCode.Name = "lbl_v180_MasterCust_TrimCode"
        Me.lbl_v180_MasterCust_TrimCode.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_MasterCust_TrimCode.TabIndex = 104
        Me.lbl_v180_MasterCust_TrimCode.Text = "Progress...."
        '
        'Label49
        '
        Me.Label49.AutoSize = True
        Me.Label49.Location = New System.Drawing.Point(4, 149)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(133, 17)
        Me.Label49.TabIndex = 103
        Me.Label49.Text = "Trim kode customer"
        '
        'prg_v180_MasterCust_TrimCode
        '
        Me.prg_v180_MasterCust_TrimCode.Location = New System.Drawing.Point(265, 143)
        Me.prg_v180_MasterCust_TrimCode.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_MasterCust_TrimCode.Name = "prg_v180_MasterCust_TrimCode"
        Me.prg_v180_MasterCust_TrimCode.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_MasterCust_TrimCode.TabIndex = 102
        '
        'btn_v180_QuotationLimit
        '
        Me.btn_v180_QuotationLimit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_QuotationLimit.Location = New System.Drawing.Point(833, 95)
        Me.btn_v180_QuotationLimit.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_QuotationLimit.Name = "btn_v180_QuotationLimit"
        Me.btn_v180_QuotationLimit.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_QuotationLimit.TabIndex = 101
        Me.btn_v180_QuotationLimit.TabStop = False
        Me.btn_v180_QuotationLimit.Text = "Update"
        Me.btn_v180_QuotationLimit.UseVisualStyleBackColor = True
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(265, 124)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(81, 17)
        Me.Label46.TabIndex = 100
        Me.Label46.Text = "Progress...."
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Location = New System.Drawing.Point(4, 105)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(182, 17)
        Me.Label47.TabIndex = 99
        Me.Label47.Text = "Pastikan limit quotation lvl 2"
        '
        'ProgressBar8
        '
        Me.ProgressBar8.Location = New System.Drawing.Point(265, 99)
        Me.ProgressBar8.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar8.Name = "ProgressBar8"
        Me.ProgressBar8.Size = New System.Drawing.Size(560, 23)
        Me.ProgressBar8.TabIndex = 98
        '
        'btn_v180_JournalPaymentMethod
        '
        Me.btn_v180_JournalPaymentMethod.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_JournalPaymentMethod.Location = New System.Drawing.Point(833, 51)
        Me.btn_v180_JournalPaymentMethod.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_JournalPaymentMethod.Name = "btn_v180_JournalPaymentMethod"
        Me.btn_v180_JournalPaymentMethod.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_JournalPaymentMethod.TabIndex = 97
        Me.btn_v180_JournalPaymentMethod.TabStop = False
        Me.btn_v180_JournalPaymentMethod.Text = "Update"
        Me.btn_v180_JournalPaymentMethod.UseVisualStyleBackColor = True
        '
        'Label43
        '
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(265, 80)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(81, 17)
        Me.Label43.TabIndex = 96
        Me.Label43.Text = "Progress...."
        '
        'Label45
        '
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(4, 61)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(307, 17)
        Me.Label45.TabIndex = 95
        Me.Label45.Text = "Pastikan kolom payment method default value ''"
        '
        'ProgressBar7
        '
        Me.ProgressBar7.Location = New System.Drawing.Point(317, 55)
        Me.ProgressBar7.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar7.Name = "ProgressBar7"
        Me.ProgressBar7.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar7.TabIndex = 94
        '
        'btn_v180_BundlingDate
        '
        Me.btn_v180_BundlingDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_v180_BundlingDate.Location = New System.Drawing.Point(833, 7)
        Me.btn_v180_BundlingDate.Margin = New System.Windows.Forms.Padding(4)
        Me.btn_v180_BundlingDate.Name = "btn_v180_BundlingDate"
        Me.btn_v180_BundlingDate.Size = New System.Drawing.Size(83, 28)
        Me.btn_v180_BundlingDate.TabIndex = 93
        Me.btn_v180_BundlingDate.TabStop = False
        Me.btn_v180_BundlingDate.Text = "Update"
        Me.btn_v180_BundlingDate.UseVisualStyleBackColor = True
        '
        'lbl_v180_BundlingDate
        '
        Me.lbl_v180_BundlingDate.AutoSize = True
        Me.lbl_v180_BundlingDate.Location = New System.Drawing.Point(265, 36)
        Me.lbl_v180_BundlingDate.Name = "lbl_v180_BundlingDate"
        Me.lbl_v180_BundlingDate.Size = New System.Drawing.Size(81, 17)
        Me.lbl_v180_BundlingDate.TabIndex = 92
        Me.lbl_v180_BundlingDate.Text = "Progress...."
        '
        'Label44
        '
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(4, 17)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(201, 17)
        Me.Label44.TabIndex = 91
        Me.Label44.Text = "Update Date pd tabel bundling"
        '
        'prg_v180_BundlingDate
        '
        Me.prg_v180_BundlingDate.Location = New System.Drawing.Point(265, 11)
        Me.prg_v180_BundlingDate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg_v180_BundlingDate.Name = "prg_v180_BundlingDate"
        Me.prg_v180_BundlingDate.Size = New System.Drawing.Size(560, 23)
        Me.prg_v180_BundlingDate.TabIndex = 90
        '
        'TabPage10
        '
        Me.TabPage10.Controls.Add(Me.Button14)
        Me.TabPage10.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.TabPage10.Controls.Add(Me.btnFilter)
        Me.TabPage10.Controls.Add(Me.TxtDesc)
        Me.TabPage10.Controls.Add(Me.btnFindCodeNo)
        Me.TabPage10.Controls.Add(Me.Label55)
        Me.TabPage10.Controls.Add(Me.TxtCodeNo)
        Me.TabPage10.Controls.Add(Me.TabControl2)
        Me.TabPage10.Location = New System.Drawing.Point(4, 25)
        Me.TabPage10.Name = "TabPage10"
        Me.TabPage10.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage10.Size = New System.Drawing.Size(923, 603)
        Me.TabPage10.TabIndex = 9
        Me.TabPage10.Text = "SC vs Qty"
        Me.TabPage10.UseVisualStyleBackColor = True
        '
        'Button14
        '
        Me.Button14.Location = New System.Drawing.Point(6, 274)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(129, 23)
        Me.Button14.TabIndex = 7
        Me.Button14.Text = "Del SC CMSI"
        Me.Button14.UseVisualStyleBackColor = True
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(173, 6)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(50, 22)
        Me.TxtPrimaryCodeNo.TabIndex = 6
        Me.TxtPrimaryCodeNo.Text = "-1"
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(538, 4)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(75, 23)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'TxtDesc
        '
        Me.TxtDesc.Location = New System.Drawing.Point(299, 5)
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.ReadOnly = True
        Me.TxtDesc.Size = New System.Drawing.Size(233, 22)
        Me.TxtDesc.TabIndex = 4
        '
        'btnFindCodeNo
        '
        Me.btnFindCodeNo.Location = New System.Drawing.Point(218, 5)
        Me.btnFindCodeNo.Name = "btnFindCodeNo"
        Me.btnFindCodeNo.Size = New System.Drawing.Size(75, 23)
        Me.btnFindCodeNo.TabIndex = 3
        Me.btnFindCodeNo.Text = "Find"
        Me.btnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Label55
        '
        Me.Label55.AutoSize = True
        Me.Label55.Location = New System.Drawing.Point(7, 11)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(63, 17)
        Me.Label55.TabIndex = 2
        Me.Label55.Text = "Code No"
        '
        'TxtCodeNo
        '
        Me.TxtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCodeNo.Location = New System.Drawing.Point(76, 6)
        Me.TxtCodeNo.Name = "TxtCodeNo"
        Me.TxtCodeNo.Size = New System.Drawing.Size(136, 22)
        Me.TxtCodeNo.TabIndex = 1
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage11)
        Me.TabControl2.Controls.Add(Me.TabPage12)
        Me.TabControl2.Location = New System.Drawing.Point(6, 34)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(911, 234)
        Me.TabControl2.TabIndex = 0
        '
        'TabPage11
        '
        Me.TabPage11.Controls.Add(Me.dgvSC)
        Me.TabPage11.Location = New System.Drawing.Point(4, 25)
        Me.TabPage11.Name = "TabPage11"
        Me.TabPage11.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage11.Size = New System.Drawing.Size(903, 205)
        Me.TabPage11.TabIndex = 0
        Me.TabPage11.Text = "Stock Card"
        Me.TabPage11.UseVisualStyleBackColor = True
        '
        'dgvSC
        '
        Me.dgvSC.AllowUserToAddRows = False
        Me.dgvSC.AllowUserToDeleteRows = False
        Me.dgvSC.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvSC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSC.Location = New System.Drawing.Point(3, 3)
        Me.dgvSC.Name = "dgvSC"
        Me.dgvSC.ReadOnly = True
        Me.dgvSC.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvSC.RowTemplate.Height = 24
        Me.dgvSC.Size = New System.Drawing.Size(897, 199)
        Me.dgvSC.TabIndex = 0
        '
        'TabPage12
        '
        Me.TabPage12.Controls.Add(Me.dgvQty)
        Me.TabPage12.Location = New System.Drawing.Point(4, 25)
        Me.TabPage12.Name = "TabPage12"
        Me.TabPage12.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage12.Size = New System.Drawing.Size(903, 205)
        Me.TabPage12.TabIndex = 1
        Me.TabPage12.Text = "Qty"
        Me.TabPage12.UseVisualStyleBackColor = True
        '
        'dgvQty
        '
        Me.dgvQty.AllowUserToAddRows = False
        Me.dgvQty.AllowUserToDeleteRows = False
        Me.dgvQty.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvQty.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvQty.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvQty.Location = New System.Drawing.Point(3, 3)
        Me.dgvQty.Name = "dgvQty"
        Me.dgvQty.ReadOnly = True
        Me.dgvQty.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvQty.RowTemplate.Height = 24
        Me.dgvQty.Size = New System.Drawing.Size(897, 199)
        Me.dgvQty.TabIndex = 1
        '
        'TabPage13
        '
        Me.TabPage13.Controls.Add(Me.btnUpdate2018_ItemKartuAPAR)
        Me.TabPage13.Controls.Add(Me.Label87)
        Me.TabPage13.Controls.Add(Me.Label89)
        Me.TabPage13.Controls.Add(Me.ProgressBar20)
        Me.TabPage13.Controls.Add(Me.LinkLabel3)
        Me.TabPage13.Controls.Add(Me.LinkLabel2)
        Me.TabPage13.Controls.Add(Me.LinkLabel1)
        Me.TabPage13.Controls.Add(Me.btnUpdate2018_ItemProdQtyMonthlyClosing)
        Me.TabPage13.Controls.Add(Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing)
        Me.TabPage13.Controls.Add(Me.Label88)
        Me.TabPage13.Controls.Add(Me.prg2018_UpdateItemProdQtyMonthlyClosing)
        Me.TabPage13.Controls.Add(Me.btnUpdate2018_PrimarySalesP)
        Me.TabPage13.Controls.Add(Me.Label80)
        Me.TabPage13.Controls.Add(Me.Label81)
        Me.TabPage13.Controls.Add(Me.ProgressBar17)
        Me.TabPage13.Controls.Add(Me.btnUpdate2018_AOP)
        Me.TabPage13.Controls.Add(Me.Label70)
        Me.TabPage13.Controls.Add(Me.Label72)
        Me.TabPage13.Controls.Add(Me.ProgressBar14)
        Me.TabPage13.Controls.Add(Me.Button16)
        Me.TabPage13.Controls.Add(Me.Label65)
        Me.TabPage13.Controls.Add(Me.Label67)
        Me.TabPage13.Controls.Add(Me.ProgressBar13)
        Me.TabPage13.Controls.Add(Me.btn2018_UpdateSODO)
        Me.TabPage13.Controls.Add(Me.LblProgress2018_UpdateSODO)
        Me.TabPage13.Controls.Add(Me.Label66)
        Me.TabPage13.Controls.Add(Me.ProgressBar12)
        Me.TabPage13.Controls.Add(Me.btn2018_UpdateMO)
        Me.TabPage13.Controls.Add(Me.LblProgress2018_UpdateMO)
        Me.TabPage13.Controls.Add(Me.Label63)
        Me.TabPage13.Controls.Add(Me.prg2018_UpdateMO)
        Me.TabPage13.Controls.Add(Me.Button15)
        Me.TabPage13.Controls.Add(Me.Label60)
        Me.TabPage13.Controls.Add(Me.Label61)
        Me.TabPage13.Controls.Add(Me.ProgressBar11)
        Me.TabPage13.Location = New System.Drawing.Point(4, 25)
        Me.TabPage13.Name = "TabPage13"
        Me.TabPage13.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage13.Size = New System.Drawing.Size(923, 603)
        Me.TabPage13.TabIndex = 10
        Me.TabPage13.Text = "v.2018.0.0"
        Me.TabPage13.UseVisualStyleBackColor = True
        '
        'LinkLabel3
        '
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(4, 371)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(391, 17)
        Me.LinkLabel3.TabIndex = 149
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Check Consistency: Item Product Qty vs Item Closing Monthly"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(4, 404)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(367, 17)
        Me.LinkLabel2.TabIndex = 147
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Update Status Item WO: Set Finish that Progress = 100%"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(4, 337)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(364, 17)
        Me.LinkLabel1.TabIndex = 146
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Update Item Purchase Monthly Closing: Item Product Qty"
        '
        'btnUpdate2018_ItemProdQtyMonthlyClosing
        '
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Location = New System.Drawing.Point(833, 277)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Name = "btnUpdate2018_ItemProdQtyMonthlyClosing"
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Size = New System.Drawing.Size(83, 28)
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.TabIndex = 145
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.TabStop = False
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.Text = "Update"
        Me.btnUpdate2018_ItemProdQtyMonthlyClosing.UseVisualStyleBackColor = True
        '
        'LblProgress2018_UpdateItemProdQtyMonthlyClosing
        '
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.AutoSize = True
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Location = New System.Drawing.Point(317, 306)
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Name = "LblProgress2018_UpdateItemProdQtyMonthlyClosing"
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Size = New System.Drawing.Size(81, 17)
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.TabIndex = 144
        Me.LblProgress2018_UpdateItemProdQtyMonthlyClosing.Text = "Progress...."
        '
        'Label88
        '
        Me.Label88.AutoSize = True
        Me.Label88.Location = New System.Drawing.Point(4, 286)
        Me.Label88.Name = "Label88"
        Me.Label88.Size = New System.Drawing.Size(267, 17)
        Me.Label88.TabIndex = 143
        Me.Label88.Text = "Update Item Produ Qty _ Monthly Closing"
        '
        'prg2018_UpdateItemProdQtyMonthlyClosing
        '
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Location = New System.Drawing.Point(317, 280)
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Name = "prg2018_UpdateItemProdQtyMonthlyClosing"
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.Size = New System.Drawing.Size(508, 23)
        Me.prg2018_UpdateItemProdQtyMonthlyClosing.TabIndex = 142
        '
        'btnUpdate2018_PrimarySalesP
        '
        Me.btnUpdate2018_PrimarySalesP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate2018_PrimarySalesP.Location = New System.Drawing.Point(833, 232)
        Me.btnUpdate2018_PrimarySalesP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate2018_PrimarySalesP.Name = "btnUpdate2018_PrimarySalesP"
        Me.btnUpdate2018_PrimarySalesP.Size = New System.Drawing.Size(83, 28)
        Me.btnUpdate2018_PrimarySalesP.TabIndex = 141
        Me.btnUpdate2018_PrimarySalesP.TabStop = False
        Me.btnUpdate2018_PrimarySalesP.Text = "Update"
        Me.btnUpdate2018_PrimarySalesP.UseVisualStyleBackColor = True
        '
        'Label80
        '
        Me.Label80.AutoSize = True
        Me.Label80.Location = New System.Drawing.Point(317, 261)
        Me.Label80.Name = "Label80"
        Me.Label80.Size = New System.Drawing.Size(81, 17)
        Me.Label80.TabIndex = 140
        Me.Label80.Text = "Progress...."
        '
        'Label81
        '
        Me.Label81.AutoSize = True
        Me.Label81.Location = New System.Drawing.Point(4, 241)
        Me.Label81.Name = "Label81"
        Me.Label81.Size = New System.Drawing.Size(255, 17)
        Me.Label81.TabIndex = 139
        Me.Label81.Text = "Update Primary Sales Person - Item Inv"
        '
        'ProgressBar17
        '
        Me.ProgressBar17.Location = New System.Drawing.Point(317, 235)
        Me.ProgressBar17.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar17.Name = "ProgressBar17"
        Me.ProgressBar17.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar17.TabIndex = 138
        '
        'btnUpdate2018_AOP
        '
        Me.btnUpdate2018_AOP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate2018_AOP.Location = New System.Drawing.Point(833, 187)
        Me.btnUpdate2018_AOP.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate2018_AOP.Name = "btnUpdate2018_AOP"
        Me.btnUpdate2018_AOP.Size = New System.Drawing.Size(83, 28)
        Me.btnUpdate2018_AOP.TabIndex = 137
        Me.btnUpdate2018_AOP.TabStop = False
        Me.btnUpdate2018_AOP.Text = "Update"
        Me.btnUpdate2018_AOP.UseVisualStyleBackColor = True
        '
        'Label70
        '
        Me.Label70.AutoSize = True
        Me.Label70.Location = New System.Drawing.Point(317, 216)
        Me.Label70.Name = "Label70"
        Me.Label70.Size = New System.Drawing.Size(81, 17)
        Me.Label70.TabIndex = 136
        Me.Label70.Text = "Progress...."
        '
        'Label72
        '
        Me.Label72.AutoSize = True
        Me.Label72.Location = New System.Drawing.Point(4, 196)
        Me.Label72.Name = "Label72"
        Me.Label72.Size = New System.Drawing.Size(138, 17)
        Me.Label72.TabIndex = 135
        Me.Label72.Text = "Update Header AOP"
        '
        'ProgressBar14
        '
        Me.ProgressBar14.Location = New System.Drawing.Point(317, 190)
        Me.ProgressBar14.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar14.Name = "ProgressBar14"
        Me.ProgressBar14.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar14.TabIndex = 134
        '
        'Button16
        '
        Me.Button16.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button16.Location = New System.Drawing.Point(833, 142)
        Me.Button16.Margin = New System.Windows.Forms.Padding(4)
        Me.Button16.Name = "Button16"
        Me.Button16.Size = New System.Drawing.Size(83, 28)
        Me.Button16.TabIndex = 133
        Me.Button16.TabStop = False
        Me.Button16.Text = "Update"
        Me.Button16.UseVisualStyleBackColor = True
        '
        'Label65
        '
        Me.Label65.AutoSize = True
        Me.Label65.Location = New System.Drawing.Point(317, 171)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(81, 17)
        Me.Label65.TabIndex = 132
        Me.Label65.Text = "Progress...."
        '
        'Label67
        '
        Me.Label67.AutoSize = True
        Me.Label67.Location = New System.Drawing.Point(4, 151)
        Me.Label67.Name = "Label67"
        Me.Label67.Size = New System.Drawing.Size(108, 17)
        Me.Label67.TabIndex = 131
        Me.Label67.Text = "Update Item OT"
        '
        'ProgressBar13
        '
        Me.ProgressBar13.Location = New System.Drawing.Point(317, 145)
        Me.ProgressBar13.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar13.Name = "ProgressBar13"
        Me.ProgressBar13.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar13.TabIndex = 130
        '
        'btn2018_UpdateSODO
        '
        Me.btn2018_UpdateSODO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn2018_UpdateSODO.Location = New System.Drawing.Point(833, 97)
        Me.btn2018_UpdateSODO.Margin = New System.Windows.Forms.Padding(4)
        Me.btn2018_UpdateSODO.Name = "btn2018_UpdateSODO"
        Me.btn2018_UpdateSODO.Size = New System.Drawing.Size(83, 28)
        Me.btn2018_UpdateSODO.TabIndex = 129
        Me.btn2018_UpdateSODO.TabStop = False
        Me.btn2018_UpdateSODO.Text = "Update"
        Me.btn2018_UpdateSODO.UseVisualStyleBackColor = True
        '
        'LblProgress2018_UpdateSODO
        '
        Me.LblProgress2018_UpdateSODO.AutoSize = True
        Me.LblProgress2018_UpdateSODO.Location = New System.Drawing.Point(317, 126)
        Me.LblProgress2018_UpdateSODO.Name = "LblProgress2018_UpdateSODO"
        Me.LblProgress2018_UpdateSODO.Size = New System.Drawing.Size(81, 17)
        Me.LblProgress2018_UpdateSODO.TabIndex = 128
        Me.LblProgress2018_UpdateSODO.Text = "Progress...."
        '
        'Label66
        '
        Me.Label66.AutoSize = True
        Me.Label66.Location = New System.Drawing.Point(4, 106)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(200, 17)
        Me.Label66.TabIndex = 127
        Me.Label66.Text = "Identifier Cust Deliver SO - DO"
        '
        'ProgressBar12
        '
        Me.ProgressBar12.Location = New System.Drawing.Point(317, 100)
        Me.ProgressBar12.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar12.Name = "ProgressBar12"
        Me.ProgressBar12.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar12.TabIndex = 126
        '
        'btn2018_UpdateMO
        '
        Me.btn2018_UpdateMO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn2018_UpdateMO.Location = New System.Drawing.Point(833, 52)
        Me.btn2018_UpdateMO.Margin = New System.Windows.Forms.Padding(4)
        Me.btn2018_UpdateMO.Name = "btn2018_UpdateMO"
        Me.btn2018_UpdateMO.Size = New System.Drawing.Size(83, 28)
        Me.btn2018_UpdateMO.TabIndex = 125
        Me.btn2018_UpdateMO.TabStop = False
        Me.btn2018_UpdateMO.Text = "Update"
        Me.btn2018_UpdateMO.UseVisualStyleBackColor = True
        '
        'LblProgress2018_UpdateMO
        '
        Me.LblProgress2018_UpdateMO.AutoSize = True
        Me.LblProgress2018_UpdateMO.Location = New System.Drawing.Point(317, 81)
        Me.LblProgress2018_UpdateMO.Name = "LblProgress2018_UpdateMO"
        Me.LblProgress2018_UpdateMO.Size = New System.Drawing.Size(81, 17)
        Me.LblProgress2018_UpdateMO.TabIndex = 124
        Me.LblProgress2018_UpdateMO.Text = "Progress...."
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Location = New System.Drawing.Point(4, 61)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(135, 17)
        Me.Label63.TabIndex = 123
        Me.Label63.Text = "Update Auth Str MO"
        '
        'prg2018_UpdateMO
        '
        Me.prg2018_UpdateMO.Location = New System.Drawing.Point(317, 55)
        Me.prg2018_UpdateMO.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.prg2018_UpdateMO.Name = "prg2018_UpdateMO"
        Me.prg2018_UpdateMO.Size = New System.Drawing.Size(508, 23)
        Me.prg2018_UpdateMO.TabIndex = 122
        '
        'Button15
        '
        Me.Button15.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button15.Location = New System.Drawing.Point(833, 7)
        Me.Button15.Margin = New System.Windows.Forms.Padding(4)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(83, 28)
        Me.Button15.TabIndex = 121
        Me.Button15.TabStop = False
        Me.Button15.Text = "Update"
        Me.Button15.UseVisualStyleBackColor = True
        '
        'Label60
        '
        Me.Label60.AutoSize = True
        Me.Label60.Location = New System.Drawing.Point(317, 36)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(81, 17)
        Me.Label60.TabIndex = 120
        Me.Label60.Text = "Progress...."
        '
        'Label61
        '
        Me.Label61.AutoSize = True
        Me.Label61.Location = New System.Drawing.Point(4, 16)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(299, 17)
        Me.Label61.TabIndex = 119
        Me.Label61.Text = "Update Tipe Item Detail Closing + Item Journal"
        '
        'ProgressBar11
        '
        Me.ProgressBar11.Location = New System.Drawing.Point(317, 10)
        Me.ProgressBar11.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar11.Name = "ProgressBar11"
        Me.ProgressBar11.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar11.TabIndex = 118
        '
        'TabPage14
        '
        Me.TabPage14.Controls.Add(Me.Button19)
        Me.TabPage14.Controls.Add(Me.Button18)
        Me.TabPage14.Controls.Add(Me.Button17)
        Me.TabPage14.Location = New System.Drawing.Point(4, 25)
        Me.TabPage14.Name = "TabPage14"
        Me.TabPage14.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage14.Size = New System.Drawing.Size(923, 603)
        Me.TabPage14.TabIndex = 11
        Me.TabPage14.Text = "Closing Tools"
        Me.TabPage14.UseVisualStyleBackColor = True
        '
        'Button19
        '
        Me.Button19.Location = New System.Drawing.Point(6, 68)
        Me.Button19.Name = "Button19"
        Me.Button19.Size = New System.Drawing.Size(183, 25)
        Me.Button19.TabIndex = 2
        Me.Button19.Text = "Version Control"
        Me.Button19.UseVisualStyleBackColor = True
        '
        'Button18
        '
        Me.Button18.Location = New System.Drawing.Point(6, 37)
        Me.Button18.Name = "Button18"
        Me.Button18.Size = New System.Drawing.Size(183, 25)
        Me.Button18.TabIndex = 1
        Me.Button18.Text = "Reset Stock LMB 2016"
        Me.Button18.UseVisualStyleBackColor = True
        '
        'Button17
        '
        Me.Button17.Location = New System.Drawing.Point(6, 6)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(183, 25)
        Me.Button17.TabIndex = 0
        Me.Button17.Text = "Monthly Closing Tools"
        Me.Button17.UseVisualStyleBackColor = True
        '
        'TabPage15
        '
        Me.TabPage15.Controls.Add(Me.Label76)
        Me.TabPage15.Controls.Add(Me.dtpCheck_DOItem)
        Me.TabPage15.Controls.Add(Me.btnCheckItemDO)
        Me.TabPage15.Controls.Add(Me.lblCheck_ItemDO)
        Me.TabPage15.Controls.Add(Me.Label77)
        Me.TabPage15.Controls.Add(Me.prgCheck_ItemDO)
        Me.TabPage15.Location = New System.Drawing.Point(4, 25)
        Me.TabPage15.Name = "TabPage15"
        Me.TabPage15.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage15.Size = New System.Drawing.Size(923, 603)
        Me.TabPage15.TabIndex = 12
        Me.TabPage15.Text = "Data Check"
        Me.TabPage15.UseVisualStyleBackColor = True
        '
        'Label76
        '
        Me.Label76.AutoSize = True
        Me.Label76.Location = New System.Drawing.Point(259, 10)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(79, 17)
        Me.Label76.TabIndex = 63
        Me.Label76.Text = "Ref. Period"
        '
        'dtpCheck_DOItem
        '
        Me.dtpCheck_DOItem.CustomFormat = "dd MMM yyyy"
        Me.dtpCheck_DOItem.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpCheck_DOItem.Location = New System.Drawing.Point(344, 6)
        Me.dtpCheck_DOItem.Name = "dtpCheck_DOItem"
        Me.dtpCheck_DOItem.Size = New System.Drawing.Size(167, 22)
        Me.dtpCheck_DOItem.TabIndex = 62
        '
        'btnCheckItemDO
        '
        Me.btnCheckItemDO.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCheckItemDO.Location = New System.Drawing.Point(834, 4)
        Me.btnCheckItemDO.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCheckItemDO.Name = "btnCheckItemDO"
        Me.btnCheckItemDO.Size = New System.Drawing.Size(82, 28)
        Me.btnCheckItemDO.TabIndex = 61
        Me.btnCheckItemDO.TabStop = False
        Me.btnCheckItemDO.Text = "Update"
        Me.btnCheckItemDO.UseVisualStyleBackColor = True
        '
        'lblCheck_ItemDO
        '
        Me.lblCheck_ItemDO.AutoSize = True
        Me.lblCheck_ItemDO.Location = New System.Drawing.Point(259, 32)
        Me.lblCheck_ItemDO.Name = "lblCheck_ItemDO"
        Me.lblCheck_ItemDO.Size = New System.Drawing.Size(81, 17)
        Me.lblCheck_ItemDO.TabIndex = 60
        Me.lblCheck_ItemDO.Text = "Progress...."
        '
        'Label77
        '
        Me.Label77.AutoSize = True
        Me.Label77.Location = New System.Drawing.Point(6, 12)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(102, 17)
        Me.Label77.TabIndex = 59
        Me.Label77.Text = "Check Item DO"
        '
        'prgCheck_ItemDO
        '
        Me.prgCheck_ItemDO.Location = New System.Drawing.Point(517, 6)
        Me.prgCheck_ItemDO.Name = "prgCheck_ItemDO"
        Me.prgCheck_ItemDO.Size = New System.Drawing.Size(310, 23)
        Me.prgCheck_ItemDO.TabIndex = 58
        '
        'BackgroundWorker7
        '
        '
        'Button11
        '
        Me.Button11.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button11.Location = New System.Drawing.Point(834, 50)
        Me.Button11.Margin = New System.Windows.Forms.Padding(4)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(82, 28)
        Me.Button11.TabIndex = 21
        Me.Button11.TabStop = False
        Me.Button11.Text = "Update"
        Me.Button11.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(245, 58)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(81, 17)
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Progress...."
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 58)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(233, 17)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "Update Numbering Collection Letter"
        '
        'Button12
        '
        Me.Button12.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button12.Location = New System.Drawing.Point(834, 4)
        Me.Button12.Margin = New System.Windows.Forms.Padding(4)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(82, 28)
        Me.Button12.TabIndex = 17
        Me.Button12.TabStop = False
        Me.Button12.Text = "Update"
        Me.Button12.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(245, 32)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(81, 17)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "Progress...."
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(6, 12)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(161, 17)
        Me.Label12.TabIndex = 1
        Me.Label12.Text = "Update Stock Card Data"
        '
        'ProgressBar2
        '
        Me.ProgressBar2.Location = New System.Drawing.Point(245, 6)
        Me.ProgressBar2.Name = "ProgressBar2"
        Me.ProgressBar2.Size = New System.Drawing.Size(582, 23)
        Me.ProgressBar2.TabIndex = 0
        '
        'BgwBalikJurnal
        '
        '
        'BgwARPay
        '
        '
        'BgwChkInv
        '
        '
        'BgwClearTable
        '
        '
        'BgwUpdPrimarySalesPerson
        '
        '
        'TabPage16
        '
        Me.TabPage16.Controls.Add(Me.btnv2021)
        Me.TabPage16.Location = New System.Drawing.Point(4, 25)
        Me.TabPage16.Name = "TabPage16"
        Me.TabPage16.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage16.Size = New System.Drawing.Size(923, 603)
        Me.TabPage16.TabIndex = 13
        Me.TabPage16.Text = "v.2021"
        Me.TabPage16.UseVisualStyleBackColor = True
        '
        'btnv2021
        '
        Me.btnv2021.Location = New System.Drawing.Point(6, 6)
        Me.btnv2021.Name = "btnv2021"
        Me.btnv2021.Size = New System.Drawing.Size(75, 23)
        Me.btnv2021.TabIndex = 0
        Me.btnv2021.Text = "v.2021"
        Me.btnv2021.UseVisualStyleBackColor = True
        '
        'btnUpdate2018_ItemKartuAPAR
        '
        Me.btnUpdate2018_ItemKartuAPAR.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdate2018_ItemKartuAPAR.Location = New System.Drawing.Point(833, 426)
        Me.btnUpdate2018_ItemKartuAPAR.Margin = New System.Windows.Forms.Padding(4)
        Me.btnUpdate2018_ItemKartuAPAR.Name = "btnUpdate2018_ItemKartuAPAR"
        Me.btnUpdate2018_ItemKartuAPAR.Size = New System.Drawing.Size(83, 28)
        Me.btnUpdate2018_ItemKartuAPAR.TabIndex = 153
        Me.btnUpdate2018_ItemKartuAPAR.TabStop = False
        Me.btnUpdate2018_ItemKartuAPAR.Text = "Update"
        Me.btnUpdate2018_ItemKartuAPAR.UseVisualStyleBackColor = True
        '
        'Label87
        '
        Me.Label87.AutoSize = True
        Me.Label87.Location = New System.Drawing.Point(317, 455)
        Me.Label87.Name = "Label87"
        Me.Label87.Size = New System.Drawing.Size(81, 17)
        Me.Label87.TabIndex = 152
        Me.Label87.Text = "Progress...."
        '
        'Label89
        '
        Me.Label89.AutoSize = True
        Me.Label89.Location = New System.Drawing.Point(4, 435)
        Me.Label89.Name = "Label89"
        Me.Label89.Size = New System.Drawing.Size(135, 17)
        Me.Label89.TabIndex = 151
        Me.Label89.Text = "Update kartu AP AR"
        '
        'ProgressBar20
        '
        Me.ProgressBar20.Location = New System.Drawing.Point(317, 429)
        Me.ProgressBar20.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar20.Name = "ProgressBar20"
        Me.ProgressBar20.Size = New System.Drawing.Size(508, 23)
        Me.ProgressBar20.TabIndex = 150
        '
        'FrmUpdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(955, 656)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "FrmUpdate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Update Utility"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        Me.TabPage5.ResumeLayout(False)
        Me.TabPage5.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.dgvCT_DetailTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage6.ResumeLayout(False)
        Me.TabPage6.PerformLayout()
        Me.TabPage7.ResumeLayout(False)
        Me.TabPage7.PerformLayout()
        Me.TabPage8.ResumeLayout(False)
        Me.TabPage8.PerformLayout()
        Me.TabPage9.ResumeLayout(False)
        Me.TabPage9.PerformLayout()
        Me.TabPage10.ResumeLayout(False)
        Me.TabPage10.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage11.ResumeLayout(False)
        CType(Me.dgvSC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage12.ResumeLayout(False)
        CType(Me.dgvQty, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage13.ResumeLayout(False)
        Me.TabPage13.PerformLayout()
        Me.TabPage14.ResumeLayout(False)
        Me.TabPage15.ResumeLayout(False)
        Me.TabPage15.PerformLayout()
        Me.TabPage16.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ProgPOB As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgSOB As System.Windows.Forms.ProgressBar
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ProgUpdateDO1 As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Prog_SeriFakturPajak As System.Windows.Forms.ProgressBar
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Prog_UpdateNoPO As System.Windows.Forms.ProgressBar
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Prog_UpdateNoSO As System.Windows.Forms.ProgressBar
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents ProgressBar3 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents ProgressBar4 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents prgCost As System.Windows.Forms.ProgressBar
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker5 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents prgItemPR As System.Windows.Forms.ProgressBar
    Friend WithEvents btnItemPR As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker6 As System.ComponentModel.BackgroundWorker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents btnUpdateAccount_StockTaking As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtAccountDesc_StockTaking As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccount_StockTaking As System.Windows.Forms.TextBox
    Friend WithEvents btnFindAccount_StockTaking As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryAccount_StockTaking As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents TxtAccountDesc_Adjustment As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccount_Adjustment As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdateAccount_Adjustment As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtAccount_Adjustment As System.Windows.Forms.TextBox
    Friend WithEvents btnFindAccount_Adjustment As System.Windows.Forms.Button
    Friend WithEvents lblPrgStockCard As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents prgStockCard As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdatev160_SC As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker7 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnUpdatev160_UpdateNumberingCL As System.Windows.Forms.Button
    Friend WithEvents LblProgress_UpdateNumberingColletionLetter As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtAccountDesc_COGS As System.Windows.Forms.TextBox
    Friend WithEvents TextBox15 As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccount_COGS As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdateAccount_COGS As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TxtAccount_COGS As System.Windows.Forms.TextBox
    Friend WithEvents btnFindAccount_COGS As System.Windows.Forms.Button
    Friend WithEvents TxtAccountDesc_SalesReturn As System.Windows.Forms.TextBox
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccount_SalesReturn As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdateAccount_SalesReturn As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtAccount_SalesReturn As System.Windows.Forms.TextBox
    Friend WithEvents btnFindAccount_SalesReturn As System.Windows.Forms.Button
    Friend WithEvents TxtAccountDesc_Sales As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccount_Sales As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdateAccount_Sales As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtAccount_Sales As System.Windows.Forms.TextBox
    Friend WithEvents btnFindAccount_Sales As System.Windows.Forms.Button
    Friend WithEvents TxtAccountDesc_Inventory As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryAccount_Inventory As System.Windows.Forms.TextBox
    Friend WithEvents btnUpdateAccount_Inventory As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtAccount_Inventory As System.Windows.Forms.TextBox
    Friend WithEvents btnFindAccount_Inventory As System.Windows.Forms.Button
    Friend WithEvents btnUpdatev160_UpdateNumberingMP As System.Windows.Forms.Button
    Friend WithEvents LblProgress_UpdateNumberingMaterialPlanning As System.Windows.Forms.Label
    Friend WithEvents LblProgress_UpdateNumbering As System.Windows.Forms.Label
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar2 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btnPindahkanSalesPrice As System.Windows.Forms.Button
    Friend WithEvents LblProgres_PindahSalesPrice As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents prgUpdateSalesPrice As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker8 As System.ComponentModel.BackgroundWorker
    Friend WithEvents prgIdentifierSO As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdate_IdentifierSO As System.Windows.Forms.Button
    Friend WithEvents LblIdentifierSO As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker9 As System.ComponentModel.BackgroundWorker
    Friend WithEvents prgIdentifierPO As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdateIdentifierPO As System.Windows.Forms.Button
    Friend WithEvents lblProgressPO As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents prgUpdatev160_TambahTabelBaru As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdatev160_TambahTabelBaru As System.Windows.Forms.Button
    Friend WithEvents lblUpdatev160_TambahTabelBaru As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents btnUpdatev160_TBLMasterProduct As System.Windows.Forms.Button
    Friend WithEvents LblUpdatev160_TBLMasterProduct As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker10 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents btnUpdateSalesPerson As System.Windows.Forms.Button
    Friend WithEvents LblUpdateSalesPerson As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents prbUpdateSalesPerson As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker11 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnDeletePackingList As System.Windows.Forms.Button
    Friend WithEvents lblPackingList As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents prgPackingList As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker12 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnUpdateEmail As System.Windows.Forms.Button
    Friend WithEvents LblUpdateEmail As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents prgUpdateEmail As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker13 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar5 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdataHR As System.Windows.Forms.Button
    Friend WithEvents lblUpdatev160_HR As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents btnUpdate_NPWP_AOP As System.Windows.Forms.Button
    Friend WithEvents lblProgress_NPWP_AOP As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents PrgUpdate_NPWP_AOP As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker14 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Prg_UpdateDiscAmount As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdateDiscAmount As System.Windows.Forms.Button
    Friend WithEvents LblUpdateDiscAmount As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker15 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage5 As System.Windows.Forms.TabPage
    Friend WithEvents dgvCT_DetailTable As System.Windows.Forms.DataGridView
    Friend WithEvents btnCT_Save As System.Windows.Forms.Button
    Friend WithEvents btnCT_Clear As System.Windows.Forms.Button
    Friend WithEvents cboCT_SchemeName As System.Windows.Forms.ComboBox
    Friend WithEvents ColCT_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_TableName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_RowCount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCT_Clear As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents btnCT_Delete As System.Windows.Forms.Button
    Friend WithEvents TxtCT_Primary As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCT_CreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCT_CreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCT_ModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCT_ModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TxtCT_ConnectionString As System.Windows.Forms.TextBox
    Friend WithEvents btnCT_LoadTableName As System.Windows.Forms.Button
    Friend WithEvents chkCT_New As System.Windows.Forms.CheckBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TxtCT_DBName As System.Windows.Forms.TextBox
    Friend WithEvents TxtCT_NamaScheme As System.Windows.Forms.TextBox
    Friend WithEvents llbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage6 As System.Windows.Forms.TabPage
    Friend WithEvents chkLocation_tblmaster_product As System.Windows.Forms.CheckBox
    Friend WithEvents prgLocationCont As System.Windows.Forms.ProgressBar
    Friend WithEvents btnLocation_Start As System.Windows.Forms.Button
    Friend WithEvents ChkLocation_tblmaster_productqty As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltransaksi_stockinout_item As System.Windows.Forms.CheckBox
    Friend WithEvents chkLocation_tbltransaksi_stockbundling_item_after As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltransaksi_stockbundling_item_before As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltransaksi_stocktransfer_item As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_aop_item As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_booking_invitem As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_purchasereturn_item As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_salesinvoice_item As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_salesinvoice_item_copy As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_salesinvoice_item_import As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_salesreturn_item As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_salesreturn_item_copy As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_salesreturn_item_import As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_stock_taking_itembuffer As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_stock_taking_itempenyesuaian As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_stock_taking_itemstart As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltransaksi_stock_card As System.Windows.Forms.CheckBox
    Friend WithEvents ChkLocation_tbltrans_stock_taking_tag As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents btn_v170_quotation As System.Windows.Forms.Button
    Friend WithEvents lbl_v170_quotation As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents prg_v170_quotation As System.Windows.Forms.ProgressBar
    Friend WithEvents BackgroundWorker16 As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage8 As System.Windows.Forms.TabPage
    Friend WithEvents btn_v180_multiunit As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_multiunit As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_multiunit As System.Windows.Forms.ProgressBar
    Friend WithEvents BgwMultiunit As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnUpdate_UpdatePOB As System.Windows.Forms.Button
    Friend WithEvents lblProgress_UpdatePOB As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents PrgUpdate_UpdatePOB As System.Windows.Forms.ProgressBar
    Friend WithEvents bgwUpdatePOB As System.ComponentModel.BackgroundWorker
    Friend WithEvents btn_v180_master_cuti As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_master_cuti As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_master_cuti As System.Windows.Forms.ProgressBar
    Friend WithEvents bgw180_MasterCuti As System.ComponentModel.BackgroundWorker
    Friend WithEvents btn_v180_multiAcc As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_multiAcc As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_multiAcc As System.Windows.Forms.ProgressBar
    Friend WithEvents BgwMultiAcc As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnUpdate_UpdatePrOB As System.Windows.Forms.Button
    Friend WithEvents lblProgress_UpdatePrOB As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents PrgUpdate_PrOB As System.Windows.Forms.ProgressBar
    Friend WithEvents bgwUpdatePrOB As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents btn_v180_PR As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_PR As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_PR As System.Windows.Forms.ProgressBar
    Friend WithEvents BgwPR As System.ComponentModel.BackgroundWorker
    Friend WithEvents btn_v180_SC As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_SC As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_SC As System.Windows.Forms.ProgressBar
    Friend WithEvents Bgw_180_SC As System.ComponentModel.BackgroundWorker
    Friend WithEvents btn_v180_Salary As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_Salary As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_Salary As System.Windows.Forms.ProgressBar
    Friend WithEvents BgwSalary As System.ComponentModel.BackgroundWorker
    Friend WithEvents btn_v180_DetailClosing As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_DetailClosing As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_DetailClosing As System.Windows.Forms.ProgressBar
    Friend WithEvents btn_v180_ProsesBundling As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_ProsesBundling As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_ProsesBundling As System.Windows.Forms.ProgressBar
    Friend WithEvents btn_v180_DetailItemJurnal As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_DetailItemJurnal As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_DetailItemJurnal As System.Windows.Forms.ProgressBar
    Friend WithEvents Button13 As System.Windows.Forms.Button
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar6 As System.Windows.Forms.ProgressBar
    Friend WithEvents btn_v180_Subtotal As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_Subtotal As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_Subtotal As System.Windows.Forms.ProgressBar
    Friend WithEvents BgwSubtotal As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage9 As System.Windows.Forms.TabPage
    Friend WithEvents btn_v180_BundlingDate As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_BundlingDate As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_BundlingDate As System.Windows.Forms.ProgressBar
    Friend WithEvents btn_v180_JournalPaymentMethod As System.Windows.Forms.Button
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar7 As System.Windows.Forms.ProgressBar
    Friend WithEvents btn_v180_QuotationLimit As System.Windows.Forms.Button
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar8 As System.Windows.Forms.ProgressBar
    Friend WithEvents btn_v180_MasterCust_TrimCode As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_MasterCust_TrimCode As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_MasterCust_TrimCode As System.Windows.Forms.ProgressBar
    Friend WithEvents Bgw_MasterCust As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnDeleteMPSF As System.Windows.Forms.Button
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents btn_v180_UpdateJurnalItem2 As System.Windows.Forms.Button
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar9 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdatePenghasilanBruto As System.Windows.Forms.Button
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents prgUpdatePenghasilanBruto As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdatePrimaryAkun As System.Windows.Forms.Button
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents btnCekSalaryKomponen As System.Windows.Forms.Button
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents dtpRefMP As System.Windows.Forms.DateTimePicker
    Friend WithEvents TabPage10 As System.Windows.Forms.TabPage
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents TxtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage11 As System.Windows.Forms.TabPage
    Friend WithEvents dgvSC As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage12 As System.Windows.Forms.TabPage
    Friend WithEvents dgvQty As System.Windows.Forms.DataGridView
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Button14 As System.Windows.Forms.Button
    Friend WithEvents btnCekUnbalanceAccount As System.Windows.Forms.Button
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents btn_v180_UpdateDiscAmount1 As System.Windows.Forms.Button
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar10 As System.Windows.Forms.ProgressBar
    Friend WithEvents prgRefMP As System.Windows.Forms.ProgressBar
    Friend WithEvents prgRefClosing As System.Windows.Forms.ProgressBar
    Friend WithEvents dtpRefClosing As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnDeleteClosing As System.Windows.Forms.Button
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents TabPage13 As System.Windows.Forms.TabPage
    Friend WithEvents Button15 As System.Windows.Forms.Button
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar11 As System.Windows.Forms.ProgressBar
    Friend WithEvents btn2018_UpdateMO As System.Windows.Forms.Button
    Friend WithEvents LblProgress2018_UpdateMO As System.Windows.Forms.Label
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents prg2018_UpdateMO As System.Windows.Forms.ProgressBar
    Friend WithEvents Bgw2018_UpdateMO As System.ComponentModel.BackgroundWorker
    Friend WithEvents TxtLocControl_Desc As System.Windows.Forms.TextBox
    Friend WithEvents TxtLocControl_TxtPrimaryProduct As System.Windows.Forms.TextBox
    Friend WithEvents BtnLocControl_FindCodeNo As System.Windows.Forms.Button
    Friend WithEvents TxtLocControl_CodeNo As System.Windows.Forms.TextBox
    Friend WithEvents dtpBalikJurnal As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblProgressBalikJurnal As System.Windows.Forms.Label
    Friend WithEvents prgBalikJurnal As System.Windows.Forms.ProgressBar
    Friend WithEvents btnBalikJurnal As System.Windows.Forms.Button
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents LblProgressARPay As System.Windows.Forms.Label
    Friend WithEvents dtpARPay As System.Windows.Forms.DateTimePicker
    Friend WithEvents prgBayarAR As System.Windows.Forms.ProgressBar
    Friend WithEvents btnBayarAR As System.Windows.Forms.Button
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents BgwBalikJurnal As System.ComponentModel.BackgroundWorker
    Friend WithEvents BgwARPay As System.ComponentModel.BackgroundWorker
    Friend WithEvents btn2018_UpdateSODO As System.Windows.Forms.Button
    Friend WithEvents LblProgress2018_UpdateSODO As System.Windows.Forms.Label
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar12 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button16 As System.Windows.Forms.Button
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label67 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar13 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label68 As System.Windows.Forms.Label
    Friend WithEvents TxtLocControl_LocationDesc As System.Windows.Forms.TextBox
    Friend WithEvents TxtLocControl_PrimaryLocation As System.Windows.Forms.TextBox
    Friend WithEvents BtnLocControl_FindLocation As System.Windows.Forms.Button
    Friend WithEvents TxtLocControl_LocationCode As System.Windows.Forms.TextBox
    Friend WithEvents Label69 As System.Windows.Forms.Label
    Friend WithEvents TabPage14 As System.Windows.Forms.TabPage
    Friend WithEvents Button17 As System.Windows.Forms.Button
    Friend WithEvents Button18 As System.Windows.Forms.Button
    Friend WithEvents btnUpdate_ChkInv As System.Windows.Forms.Button
    Friend WithEvents lblProgress_ChkInv As System.Windows.Forms.Label
    Friend WithEvents Label71 As System.Windows.Forms.Label
    Friend WithEvents PrgUpdate_ChkInv As System.Windows.Forms.ProgressBar
    Friend WithEvents BgwChkInv As System.ComponentModel.BackgroundWorker
    Friend WithEvents Button19 As System.Windows.Forms.Button
    Friend WithEvents btnUpdate2018_AOP As System.Windows.Forms.Button
    Friend WithEvents Label70 As System.Windows.Forms.Label
    Friend WithEvents Label72 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar14 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button20 As System.Windows.Forms.Button
    Friend WithEvents Label73 As System.Windows.Forms.Label
    Friend WithEvents Label74 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar15 As System.Windows.Forms.ProgressBar
    Friend WithEvents prgClearTable As System.Windows.Forms.ProgressBar
    Friend WithEvents LblClearTable_Progress As System.Windows.Forms.Label
    Friend WithEvents BgwClearTable As System.ComponentModel.BackgroundWorker
    Friend WithEvents btn_v180_UpdateSalary As System.Windows.Forms.Button
    Friend WithEvents lbl_v180_UpdateSalary As System.Windows.Forms.Label
    Friend WithEvents Label75 As System.Windows.Forms.Label
    Friend WithEvents prg_v180_UpdateSalary As System.Windows.Forms.ProgressBar
    Friend WithEvents Bgw_v180_UpdateSalary As System.ComponentModel.BackgroundWorker
    Friend WithEvents TabPage15 As System.Windows.Forms.TabPage
    Friend WithEvents btnCheckItemDO As System.Windows.Forms.Button
    Friend WithEvents lblCheck_ItemDO As System.Windows.Forms.Label
    Friend WithEvents Label77 As System.Windows.Forms.Label
    Friend WithEvents prgCheck_ItemDO As System.Windows.Forms.ProgressBar
    Friend WithEvents Bgw_Check_ItemDO As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label76 As System.Windows.Forms.Label
    Friend WithEvents dtpCheck_DOItem As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button21 As System.Windows.Forms.Button
    Friend WithEvents Label78 As System.Windows.Forms.Label
    Friend WithEvents Label79 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar16 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdate2018_PrimarySalesP As System.Windows.Forms.Button
    Friend WithEvents Label80 As System.Windows.Forms.Label
    Friend WithEvents Label81 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar17 As System.Windows.Forms.ProgressBar
    Friend WithEvents BgwUpdPrimarySalesPerson As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnUpdateGetVersion As System.Windows.Forms.Button
    Friend WithEvents btnUpdateSetVersion As System.Windows.Forms.Button
    Friend WithEvents Label82 As System.Windows.Forms.Label
    Friend WithEvents btnUpdatePresentasiPayroll As System.Windows.Forms.Button
    Friend WithEvents Label83 As System.Windows.Forms.Label
    Friend WithEvents Label84 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar18 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdateSalaryOverride As System.Windows.Forms.Button
    Friend WithEvents Label85 As System.Windows.Forms.Label
    Friend WithEvents Label86 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar19 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnUpdate2018_ItemProdQtyMonthlyClosing As System.Windows.Forms.Button
    Friend WithEvents LblProgress2018_UpdateItemProdQtyMonthlyClosing As System.Windows.Forms.Label
    Friend WithEvents Label88 As System.Windows.Forms.Label
    Friend WithEvents prg2018_UpdateItemProdQtyMonthlyClosing As System.Windows.Forms.ProgressBar
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage16 As System.Windows.Forms.TabPage
    Friend WithEvents btnv2021 As System.Windows.Forms.Button
    Friend WithEvents btnUpdate2018_ItemKartuAPAR As System.Windows.Forms.Button
    Friend WithEvents Label87 As System.Windows.Forms.Label
    Friend WithEvents Label89 As System.Windows.Forms.Label
    Friend WithEvents ProgressBar20 As System.Windows.Forms.ProgressBar
End Class
