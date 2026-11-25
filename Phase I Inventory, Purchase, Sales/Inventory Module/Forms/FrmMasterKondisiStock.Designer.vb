<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKondisiStock
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
        Me.txtCodeNo = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnFindProduct = New System.Windows.Forms.Button
        Me.txtFreeStock = New System.Windows.Forms.TextBox
        Me.txtQtyOnHand = New System.Windows.Forms.TextBox
        Me.txtBookingInv = New System.Windows.Forms.TextBox
        Me.txtGIT = New System.Windows.Forms.TextBox
        Me.lblUnit1 = New System.Windows.Forms.Label
        Me.lblUnit2 = New System.Windows.Forms.Label
        Me.lblUnit4 = New System.Windows.Forms.Label
        Me.lblUnit3 = New System.Windows.Forms.Label
        Me.llbBookingStockInv = New System.Windows.Forms.LinkLabel
        Me.llbGIT = New System.Windows.Forms.LinkLabel
        Me.llbBookingGIT = New System.Windows.Forms.LinkLabel
        Me.llbQtyOnHand = New System.Windows.Forms.LinkLabel
        Me.lblPendingDO = New System.Windows.Forms.LinkLabel
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.lblUnit5 = New System.Windows.Forms.Label
        Me.txtBookingGIT = New System.Windows.Forms.TextBox
        Me.txtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.lblUnitPendingDO = New System.Windows.Forms.Label
        Me.txtPendingDO = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblUnitAvailable = New System.Windows.Forms.Label
        Me.txtAvailable = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtUnitPrice = New System.Windows.Forms.TextBox
        Me.cboUnitPrice = New System.Windows.Forms.ComboBox
        Me.llbPOB = New System.Windows.Forms.LinkLabel
        Me.LblUnit6 = New System.Windows.Forms.Label
        Me.TxtPOB = New System.Windows.Forms.TextBox
        Me.llbSOB = New System.Windows.Forms.LinkLabel
        Me.LblUnit7 = New System.Windows.Forms.Label
        Me.TxtSOB = New System.Windows.Forms.TextBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSpec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.dgvItemPendingDO = New System.Windows.Forms.DataGridView
        Me.Stock_detail_viewDataGridView = New System.Windows.Forms.DataGridView
        Me.dgvItemPOB = New System.Windows.Forms.DataGridView
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.dgvDataRMReq = New System.Windows.Forms.DataGridView
        Me.ColRMReq_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_MONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_MODate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_JONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_WONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_Process = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_BalanceQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_UnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_IsWIP = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColRMReq_IdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRMReq_IdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvDataMOB = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMODate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProses = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTargetFinishDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProgress = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColIdentifierJO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierWO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvItemSOB = New System.Windows.Forms.DataGridView
        Me.ColSOB_SONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_CustPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_CustomerCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_CustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_OrderBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_PromiseToDeliver = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColSOB_PrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_IdentifierSOItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSOB_PrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DgvPurcReq = New System.Windows.Forms.DataGridView
        Me.ColPurcReq_PriorityLevel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_PromiseToDeliver = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_RefParty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_DocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_Description = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_QtyOB = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_BalanceQtyUsing = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_Remarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_PickUpOrder = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColPurcReq_PrimaryPickUpOrder = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_PrimaryPR = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurcReq_PrimaryPrItemDel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.llbPurcReqBalance = New System.Windows.Forms.LinkLabel
        Me.LblUnit8 = New System.Windows.Forms.Label
        Me.TxtPurcReqBalance = New System.Windows.Forms.TextBox
        Me.ChkIncludeUnAuthPO = New System.Windows.Forms.CheckBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.lblUnit9 = New System.Windows.Forms.Label
        Me.TxtRawMaterialRequest = New System.Windows.Forms.TextBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.lblUnit10 = New System.Windows.Forms.Label
        Me.TxtMOB = New System.Windows.Forms.TextBox
        Me.chkShowWONo = New System.Windows.Forms.CheckBox
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemPendingDO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Stock_detail_viewDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemPOB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvDataRMReq, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvDataMOB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvItemSOB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DgvPurcReq, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtCodeNo
        '
        Me.txtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodeNo.Location = New System.Drawing.Point(83, 12)
        Me.txtCodeNo.Name = "txtCodeNo"
        Me.txtCodeNo.Size = New System.Drawing.Size(173, 20)
        Me.txtCodeNo.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 15)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Code No"
        '
        'btnFindProduct
        '
        Me.btnFindProduct.Location = New System.Drawing.Point(261, 12)
        Me.btnFindProduct.Name = "btnFindProduct"
        Me.btnFindProduct.Size = New System.Drawing.Size(33, 20)
        Me.btnFindProduct.TabIndex = 1
        Me.btnFindProduct.Text = "...."
        Me.btnFindProduct.UseVisualStyleBackColor = True
        '
        'txtFreeStock
        '
        Me.txtFreeStock.Location = New System.Drawing.Point(164, 161)
        Me.txtFreeStock.Name = "txtFreeStock"
        Me.txtFreeStock.ReadOnly = True
        Me.txtFreeStock.Size = New System.Drawing.Size(71, 20)
        Me.txtFreeStock.TabIndex = 5
        Me.txtFreeStock.TabStop = False
        Me.txtFreeStock.Text = "0"
        Me.txtFreeStock.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtQtyOnHand
        '
        Me.txtQtyOnHand.Location = New System.Drawing.Point(164, 64)
        Me.txtQtyOnHand.Name = "txtQtyOnHand"
        Me.txtQtyOnHand.ReadOnly = True
        Me.txtQtyOnHand.Size = New System.Drawing.Size(71, 20)
        Me.txtQtyOnHand.TabIndex = 3
        Me.txtQtyOnHand.TabStop = False
        Me.txtQtyOnHand.Text = "0"
        Me.txtQtyOnHand.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtBookingInv
        '
        Me.txtBookingInv.Location = New System.Drawing.Point(164, 90)
        Me.txtBookingInv.Name = "txtBookingInv"
        Me.txtBookingInv.ReadOnly = True
        Me.txtBookingInv.Size = New System.Drawing.Size(71, 20)
        Me.txtBookingInv.TabIndex = 4
        Me.txtBookingInv.TabStop = False
        Me.txtBookingInv.Text = "0"
        Me.txtBookingInv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtGIT
        '
        Me.txtGIT.Location = New System.Drawing.Point(164, 288)
        Me.txtGIT.Name = "txtGIT"
        Me.txtGIT.ReadOnly = True
        Me.txtGIT.Size = New System.Drawing.Size(71, 20)
        Me.txtGIT.TabIndex = 9
        Me.txtGIT.TabStop = False
        Me.txtGIT.Text = "0"
        Me.txtGIT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblUnit1
        '
        Me.lblUnit1.AutoSize = True
        Me.lblUnit1.Location = New System.Drawing.Point(240, 163)
        Me.lblUnit1.Name = "lblUnit1"
        Me.lblUnit1.Size = New System.Drawing.Size(31, 13)
        Me.lblUnit1.TabIndex = 13
        Me.lblUnit1.Text = "Units"
        '
        'lblUnit2
        '
        Me.lblUnit2.AutoSize = True
        Me.lblUnit2.Location = New System.Drawing.Point(240, 93)
        Me.lblUnit2.Name = "lblUnit2"
        Me.lblUnit2.Size = New System.Drawing.Size(31, 13)
        Me.lblUnit2.TabIndex = 14
        Me.lblUnit2.Text = "Units"
        '
        'lblUnit4
        '
        Me.lblUnit4.AutoSize = True
        Me.lblUnit4.Location = New System.Drawing.Point(240, 290)
        Me.lblUnit4.Name = "lblUnit4"
        Me.lblUnit4.Size = New System.Drawing.Size(31, 13)
        Me.lblUnit4.TabIndex = 16
        Me.lblUnit4.Text = "Units"
        '
        'lblUnit3
        '
        Me.lblUnit3.AutoSize = True
        Me.lblUnit3.Location = New System.Drawing.Point(240, 67)
        Me.lblUnit3.Name = "lblUnit3"
        Me.lblUnit3.Size = New System.Drawing.Size(31, 13)
        Me.lblUnit3.TabIndex = 15
        Me.lblUnit3.Text = "Units"
        '
        'llbBookingStockInv
        '
        Me.llbBookingStockInv.AutoSize = True
        Me.llbBookingStockInv.Location = New System.Drawing.Point(12, 93)
        Me.llbBookingStockInv.Name = "llbBookingStockInv"
        Me.llbBookingStockInv.Size = New System.Drawing.Size(93, 13)
        Me.llbBookingStockInv.TabIndex = 14
        Me.llbBookingStockInv.TabStop = True
        Me.llbBookingStockInv.Text = "Booking Inventory"
        '
        'llbGIT
        '
        Me.llbGIT.AutoSize = True
        Me.llbGIT.Location = New System.Drawing.Point(12, 290)
        Me.llbGIT.Name = "llbGIT"
        Me.llbGIT.Size = New System.Drawing.Size(85, 13)
        Me.llbGIT.TabIndex = 16
        Me.llbGIT.TabStop = True
        Me.llbGIT.Text = "Goods In Transit"
        '
        'llbBookingGIT
        '
        Me.llbBookingGIT.AutoSize = True
        Me.llbBookingGIT.Location = New System.Drawing.Point(277, 97)
        Me.llbBookingGIT.Name = "llbBookingGIT"
        Me.llbBookingGIT.Size = New System.Drawing.Size(67, 13)
        Me.llbBookingGIT.TabIndex = 10
        Me.llbBookingGIT.TabStop = True
        Me.llbBookingGIT.Text = "Booking GIT"
        Me.llbBookingGIT.Visible = False
        '
        'llbQtyOnHand
        '
        Me.llbQtyOnHand.AutoSize = True
        Me.llbQtyOnHand.Location = New System.Drawing.Point(12, 67)
        Me.llbQtyOnHand.Name = "llbQtyOnHand"
        Me.llbQtyOnHand.Size = New System.Drawing.Size(69, 13)
        Me.llbQtyOnHand.TabIndex = 12
        Me.llbQtyOnHand.TabStop = True
        Me.llbQtyOnHand.Text = "Qty On Hand"
        '
        'lblPendingDO
        '
        Me.lblPendingDO.AutoSize = True
        Me.lblPendingDO.Location = New System.Drawing.Point(12, 189)
        Me.lblPendingDO.Name = "lblPendingDO"
        Me.lblPendingDO.Size = New System.Drawing.Size(65, 13)
        Me.lblPendingDO.TabIndex = 6
        Me.lblPendingDO.TabStop = True
        Me.lblPendingDO.Text = "Pending DO"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 163)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Free Stock"
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(261, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(33, 20)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'lblUnit5
        '
        Me.lblUnit5.AutoSize = True
        Me.lblUnit5.Location = New System.Drawing.Point(507, 93)
        Me.lblUnit5.Name = "lblUnit5"
        Me.lblUnit5.Size = New System.Drawing.Size(31, 13)
        Me.lblUnit5.TabIndex = 25
        Me.lblUnit5.Text = "Units"
        Me.lblUnit5.Visible = False
        '
        'txtBookingGIT
        '
        Me.txtBookingGIT.Location = New System.Drawing.Point(381, 90)
        Me.txtBookingGIT.Name = "txtBookingGIT"
        Me.txtBookingGIT.ReadOnly = True
        Me.txtBookingGIT.Size = New System.Drawing.Size(120, 20)
        Me.txtBookingGIT.TabIndex = 11
        Me.txtBookingGIT.TabStop = False
        Me.txtBookingGIT.Text = "0"
        Me.txtBookingGIT.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtBookingGIT.Visible = False
        '
        'txtPrimaryCodeNo
        '
        Me.txtPrimaryCodeNo.Location = New System.Drawing.Point(83, 12)
        Me.txtPrimaryCodeNo.Name = "txtPrimaryCodeNo"
        Me.txtPrimaryCodeNo.Size = New System.Drawing.Size(58, 20)
        Me.txtPrimaryCodeNo.TabIndex = 27
        Me.txtPrimaryCodeNo.TabStop = False
        Me.txtPrimaryCodeNo.Text = "-1"
        Me.txtPrimaryCodeNo.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 41)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Description"
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(83, 38)
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(212, 20)
        Me.txtDesc.TabIndex = 2
        Me.txtDesc.TabStop = False
        '
        'lblUnitPendingDO
        '
        Me.lblUnitPendingDO.AutoSize = True
        Me.lblUnitPendingDO.Location = New System.Drawing.Point(240, 189)
        Me.lblUnitPendingDO.Name = "lblUnitPendingDO"
        Me.lblUnitPendingDO.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitPendingDO.TabIndex = 34
        Me.lblUnitPendingDO.Text = "Units"
        '
        'txtPendingDO
        '
        Me.txtPendingDO.Location = New System.Drawing.Point(164, 187)
        Me.txtPendingDO.Name = "txtPendingDO"
        Me.txtPendingDO.ReadOnly = True
        Me.txtPendingDO.Size = New System.Drawing.Size(71, 20)
        Me.txtPendingDO.TabIndex = 6
        Me.txtPendingDO.TabStop = False
        Me.txtPendingDO.Text = "0"
        Me.txtPendingDO.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 215)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(102, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Available To Deliver"
        '
        'lblUnitAvailable
        '
        Me.lblUnitAvailable.AutoSize = True
        Me.lblUnitAvailable.Location = New System.Drawing.Point(240, 215)
        Me.lblUnitAvailable.Name = "lblUnitAvailable"
        Me.lblUnitAvailable.Size = New System.Drawing.Size(31, 13)
        Me.lblUnitAvailable.TabIndex = 36
        Me.lblUnitAvailable.Text = "Units"
        '
        'txtAvailable
        '
        Me.txtAvailable.Location = New System.Drawing.Point(164, 213)
        Me.txtAvailable.Name = "txtAvailable"
        Me.txtAvailable.ReadOnly = True
        Me.txtAvailable.Size = New System.Drawing.Size(71, 20)
        Me.txtAvailable.TabIndex = 7
        Me.txtAvailable.TabStop = False
        Me.txtAvailable.Text = "0"
        Me.txtAvailable.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 388)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "Sales Price"
        Me.Label4.Visible = False
        '
        'TxtUnitPrice
        '
        Me.TxtUnitPrice.Location = New System.Drawing.Point(164, 385)
        Me.TxtUnitPrice.Name = "TxtUnitPrice"
        Me.TxtUnitPrice.ReadOnly = True
        Me.TxtUnitPrice.Size = New System.Drawing.Size(71, 20)
        Me.TxtUnitPrice.TabIndex = 12
        Me.TxtUnitPrice.TabStop = False
        Me.TxtUnitPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtUnitPrice.Visible = False
        '
        'cboUnitPrice
        '
        Me.cboUnitPrice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUnitPrice.FormattingEnabled = True
        Me.cboUnitPrice.Location = New System.Drawing.Point(77, 385)
        Me.cboUnitPrice.Name = "cboUnitPrice"
        Me.cboUnitPrice.Size = New System.Drawing.Size(66, 21)
        Me.cboUnitPrice.TabIndex = 11
        Me.cboUnitPrice.TabStop = False
        Me.cboUnitPrice.Visible = False
        '
        'llbPOB
        '
        Me.llbPOB.AutoSize = True
        Me.llbPOB.Location = New System.Drawing.Point(12, 316)
        Me.llbPOB.Name = "llbPOB"
        Me.llbPOB.Size = New System.Drawing.Size(123, 13)
        Me.llbPOB.TabIndex = 18
        Me.llbPOB.TabStop = True
        Me.llbPOB.Text = "Purchase Order Balance"
        '
        'LblUnit6
        '
        Me.LblUnit6.AutoSize = True
        Me.LblUnit6.Location = New System.Drawing.Point(240, 316)
        Me.LblUnit6.Name = "LblUnit6"
        Me.LblUnit6.Size = New System.Drawing.Size(31, 13)
        Me.LblUnit6.TabIndex = 43
        Me.LblUnit6.Text = "Units"
        '
        'TxtPOB
        '
        Me.TxtPOB.Location = New System.Drawing.Point(164, 314)
        Me.TxtPOB.Name = "TxtPOB"
        Me.TxtPOB.ReadOnly = True
        Me.TxtPOB.Size = New System.Drawing.Size(71, 20)
        Me.TxtPOB.TabIndex = 10
        Me.TxtPOB.TabStop = False
        Me.TxtPOB.Text = "0"
        Me.TxtPOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'llbSOB
        '
        Me.llbSOB.AutoSize = True
        Me.llbSOB.Location = New System.Drawing.Point(12, 241)
        Me.llbSOB.Name = "llbSOB"
        Me.llbSOB.Size = New System.Drawing.Size(104, 13)
        Me.llbSOB.TabIndex = 44
        Me.llbSOB.TabStop = True
        Me.llbSOB.Text = "Sales Order Balance"
        '
        'LblUnit7
        '
        Me.LblUnit7.AutoSize = True
        Me.LblUnit7.Location = New System.Drawing.Point(240, 241)
        Me.LblUnit7.Name = "LblUnit7"
        Me.LblUnit7.Size = New System.Drawing.Size(31, 13)
        Me.LblUnit7.TabIndex = 46
        Me.LblUnit7.Text = "Units"
        '
        'TxtSOB
        '
        Me.TxtSOB.Location = New System.Drawing.Point(164, 239)
        Me.TxtSOB.Name = "TxtSOB"
        Me.TxtSOB.ReadOnly = True
        Me.TxtSOB.Size = New System.Drawing.Size(71, 20)
        Me.TxtSOB.TabIndex = 8
        Me.TxtSOB.TabStop = False
        Me.TxtSOB.Text = "0"
        Me.TxtSOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCodeNo, Me.ColSpec, Me.ColDesc, Me.ColQty})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 16)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(549, 461)
        Me.DataGridView1.TabIndex = 1
        Me.DataGridView1.TabStop = False
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 74
        '
        'ColSpec
        '
        Me.ColSpec.HeaderText = "Specification"
        Me.ColSpec.Name = "ColSpec"
        Me.ColSpec.ReadOnly = True
        Me.ColSpec.Width = 93
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.Width = 85
        '
        'ColQty
        '
        Me.ColQty.HeaderText = "Total Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.Width = 75
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(3, 16)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(549, 461)
        Me.dgvItem.TabIndex = 0
        Me.dgvItem.TabStop = False
        '
        'dgvItemPendingDO
        '
        Me.dgvItemPendingDO.AllowUserToAddRows = False
        Me.dgvItemPendingDO.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.dgvItemPendingDO.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItemPendingDO.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemPendingDO.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemPendingDO.Location = New System.Drawing.Point(3, 16)
        Me.dgvItemPendingDO.Name = "dgvItemPendingDO"
        Me.dgvItemPendingDO.ReadOnly = True
        Me.dgvItemPendingDO.RowTemplate.Height = 24
        Me.dgvItemPendingDO.Size = New System.Drawing.Size(549, 461)
        Me.dgvItemPendingDO.TabIndex = 2
        '
        'Stock_detail_viewDataGridView
        '
        Me.Stock_detail_viewDataGridView.AllowUserToAddRows = False
        Me.Stock_detail_viewDataGridView.AllowUserToDeleteRows = False
        Me.Stock_detail_viewDataGridView.AllowUserToResizeColumns = False
        Me.Stock_detail_viewDataGridView.AllowUserToResizeRows = False
        Me.Stock_detail_viewDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.Stock_detail_viewDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.Stock_detail_viewDataGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Stock_detail_viewDataGridView.Location = New System.Drawing.Point(3, 16)
        Me.Stock_detail_viewDataGridView.Name = "Stock_detail_viewDataGridView"
        Me.Stock_detail_viewDataGridView.ReadOnly = True
        Me.Stock_detail_viewDataGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.Stock_detail_viewDataGridView.RowTemplate.Height = 24
        Me.Stock_detail_viewDataGridView.Size = New System.Drawing.Size(549, 461)
        Me.Stock_detail_viewDataGridView.TabIndex = 0
        Me.Stock_detail_viewDataGridView.TabStop = False
        '
        'dgvItemPOB
        '
        Me.dgvItemPOB.AllowUserToAddRows = False
        Me.dgvItemPOB.AllowUserToDeleteRows = False
        Me.dgvItemPOB.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemPOB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemPOB.Location = New System.Drawing.Point(3, 16)
        Me.dgvItemPOB.Name = "dgvItemPOB"
        Me.dgvItemPOB.ReadOnly = True
        Me.dgvItemPOB.RowTemplate.Height = 24
        Me.dgvItemPOB.Size = New System.Drawing.Size(549, 461)
        Me.dgvItemPOB.TabIndex = 3
        Me.dgvItemPOB.TabStop = False
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvDataRMReq)
        Me.gboDetail.Controls.Add(Me.dgvDataMOB)
        Me.gboDetail.Controls.Add(Me.dgvItemSOB)
        Me.gboDetail.Controls.Add(Me.dgvItemPOB)
        Me.gboDetail.Controls.Add(Me.Stock_detail_viewDataGridView)
        Me.gboDetail.Controls.Add(Me.dgvItemPendingDO)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.DgvPurcReq)
        Me.gboDetail.Controls.Add(Me.DataGridView1)
        Me.gboDetail.Location = New System.Drawing.Point(300, 2)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(555, 480)
        Me.gboDetail.TabIndex = 13
        Me.gboDetail.TabStop = False
        '
        'dgvDataRMReq
        '
        Me.dgvDataRMReq.AllowUserToAddRows = False
        Me.dgvDataRMReq.AllowUserToDeleteRows = False
        Me.dgvDataRMReq.AllowUserToResizeColumns = False
        Me.dgvDataRMReq.AllowUserToResizeRows = False
        Me.dgvDataRMReq.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDataRMReq.ColumnHeadersHeight = 25
        Me.dgvDataRMReq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDataRMReq.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColRMReq_No, Me.ColRMReq_MONo, Me.ColRMReq_MODate, Me.ColRMReq_JONo, Me.ColRMReq_WONo, Me.ColRMReq_Process, Me.ColRMReq_BalanceQty, Me.ColRMReq_UnitCode, Me.ColRMReq_IsWIP, Me.ColRMReq_IdentifierJO, Me.ColRMReq_IdentifierWO})
        Me.dgvDataRMReq.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDataRMReq.Location = New System.Drawing.Point(3, 16)
        Me.dgvDataRMReq.Name = "dgvDataRMReq"
        Me.dgvDataRMReq.ReadOnly = True
        Me.dgvDataRMReq.RowHeadersVisible = False
        Me.dgvDataRMReq.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDataRMReq.RowTemplate.Height = 24
        Me.dgvDataRMReq.Size = New System.Drawing.Size(549, 461)
        Me.dgvDataRMReq.TabIndex = 11
        '
        'ColRMReq_No
        '
        Me.ColRMReq_No.HeaderText = "No"
        Me.ColRMReq_No.Name = "ColRMReq_No"
        Me.ColRMReq_No.ReadOnly = True
        Me.ColRMReq_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRMReq_No.Width = 27
        '
        'ColRMReq_MONo
        '
        Me.ColRMReq_MONo.HeaderText = "MO No"
        Me.ColRMReq_MONo.Name = "ColRMReq_MONo"
        Me.ColRMReq_MONo.ReadOnly = True
        Me.ColRMReq_MONo.Width = 66
        '
        'ColRMReq_MODate
        '
        Me.ColRMReq_MODate.HeaderText = "MO Date"
        Me.ColRMReq_MODate.Name = "ColRMReq_MODate"
        Me.ColRMReq_MODate.ReadOnly = True
        Me.ColRMReq_MODate.Width = 75
        '
        'ColRMReq_JONo
        '
        Me.ColRMReq_JONo.HeaderText = "JO No"
        Me.ColRMReq_JONo.Name = "ColRMReq_JONo"
        Me.ColRMReq_JONo.ReadOnly = True
        Me.ColRMReq_JONo.Width = 62
        '
        'ColRMReq_WONo
        '
        Me.ColRMReq_WONo.HeaderText = "WO No"
        Me.ColRMReq_WONo.Name = "ColRMReq_WONo"
        Me.ColRMReq_WONo.ReadOnly = True
        Me.ColRMReq_WONo.Width = 68
        '
        'ColRMReq_Process
        '
        Me.ColRMReq_Process.HeaderText = "Process"
        Me.ColRMReq_Process.Name = "ColRMReq_Process"
        Me.ColRMReq_Process.ReadOnly = True
        Me.ColRMReq_Process.Width = 70
        '
        'ColRMReq_BalanceQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRMReq_BalanceQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColRMReq_BalanceQty.HeaderText = "Balance Qty"
        Me.ColRMReq_BalanceQty.Name = "ColRMReq_BalanceQty"
        Me.ColRMReq_BalanceQty.ReadOnly = True
        Me.ColRMReq_BalanceQty.Width = 90
        '
        'ColRMReq_UnitCode
        '
        Me.ColRMReq_UnitCode.HeaderText = "Unit Code"
        Me.ColRMReq_UnitCode.Name = "ColRMReq_UnitCode"
        Me.ColRMReq_UnitCode.ReadOnly = True
        Me.ColRMReq_UnitCode.Width = 79
        '
        'ColRMReq_IsWIP
        '
        Me.ColRMReq_IsWIP.HeaderText = "Is WIP"
        Me.ColRMReq_IsWIP.Name = "ColRMReq_IsWIP"
        Me.ColRMReq_IsWIP.ReadOnly = True
        Me.ColRMReq_IsWIP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRMReq_IsWIP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRMReq_IsWIP.Width = 64
        '
        'ColRMReq_IdentifierJO
        '
        Me.ColRMReq_IdentifierJO.HeaderText = "Identifier JO"
        Me.ColRMReq_IdentifierJO.Name = "ColRMReq_IdentifierJO"
        Me.ColRMReq_IdentifierJO.ReadOnly = True
        Me.ColRMReq_IdentifierJO.Visible = False
        Me.ColRMReq_IdentifierJO.Width = 88
        '
        'ColRMReq_IdentifierWO
        '
        Me.ColRMReq_IdentifierWO.HeaderText = "Identifier WO"
        Me.ColRMReq_IdentifierWO.Name = "ColRMReq_IdentifierWO"
        Me.ColRMReq_IdentifierWO.ReadOnly = True
        Me.ColRMReq_IdentifierWO.Visible = False
        Me.ColRMReq_IdentifierWO.Width = 94
        '
        'dgvDataMOB
        '
        Me.dgvDataMOB.AllowUserToAddRows = False
        Me.dgvDataMOB.AllowUserToDeleteRows = False
        Me.dgvDataMOB.AllowUserToResizeColumns = False
        Me.dgvDataMOB.AllowUserToResizeRows = False
        Me.dgvDataMOB.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDataMOB.ColumnHeadersHeight = 25
        Me.dgvDataMOB.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDataMOB.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColMONo, Me.ColMODate, Me.ColJONo, Me.ColWONo, Me.ColProses, Me.ColBalanceQty, Me.ColUnitCode, Me.ColTargetFinishDate, Me.ColProgress, Me.ColIdentifierJO, Me.ColIdentifierWO})
        Me.dgvDataMOB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvDataMOB.Location = New System.Drawing.Point(3, 16)
        Me.dgvDataMOB.Name = "dgvDataMOB"
        Me.dgvDataMOB.ReadOnly = True
        Me.dgvDataMOB.RowHeadersVisible = False
        Me.dgvDataMOB.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDataMOB.RowTemplate.Height = 24
        Me.dgvDataMOB.Size = New System.Drawing.Size(549, 461)
        Me.dgvDataMOB.TabIndex = 7
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColMONo
        '
        Me.ColMONo.HeaderText = "MO No"
        Me.ColMONo.Name = "ColMONo"
        Me.ColMONo.ReadOnly = True
        Me.ColMONo.Width = 66
        '
        'ColMODate
        '
        Me.ColMODate.HeaderText = "MO Date"
        Me.ColMODate.Name = "ColMODate"
        Me.ColMODate.ReadOnly = True
        Me.ColMODate.Width = 75
        '
        'ColJONo
        '
        Me.ColJONo.HeaderText = "JO No"
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.ReadOnly = True
        Me.ColJONo.Width = 62
        '
        'ColWONo
        '
        Me.ColWONo.HeaderText = "WO No"
        Me.ColWONo.Name = "ColWONo"
        Me.ColWONo.ReadOnly = True
        Me.ColWONo.Width = 68
        '
        'ColProses
        '
        Me.ColProses.HeaderText = "Process"
        Me.ColProses.Name = "ColProses"
        Me.ColProses.ReadOnly = True
        Me.ColProses.Width = 70
        '
        'ColBalanceQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalanceQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColBalanceQty.HeaderText = "Balance Qty"
        Me.ColBalanceQty.Name = "ColBalanceQty"
        Me.ColBalanceQty.ReadOnly = True
        Me.ColBalanceQty.Width = 90
        '
        'ColUnitCode
        '
        Me.ColUnitCode.HeaderText = "Unit Code"
        Me.ColUnitCode.Name = "ColUnitCode"
        Me.ColUnitCode.ReadOnly = True
        Me.ColUnitCode.Width = 79
        '
        'ColTargetFinishDate
        '
        Me.ColTargetFinishDate.HeaderText = "Target Finish Date"
        Me.ColTargetFinishDate.Name = "ColTargetFinishDate"
        Me.ColTargetFinishDate.ReadOnly = True
        Me.ColTargetFinishDate.Visible = False
        Me.ColTargetFinishDate.Width = 119
        '
        'ColProgress
        '
        Me.ColProgress.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColProgress.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColProgress.HeaderText = "Progress"
        Me.ColProgress.Name = "ColProgress"
        Me.ColProgress.ReadOnly = True
        Me.ColProgress.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColProgress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColProgress.Visible = False
        Me.ColProgress.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColProgress.Width = 73
        '
        'ColIdentifierJO
        '
        Me.ColIdentifierJO.HeaderText = "Identifier JO"
        Me.ColIdentifierJO.Name = "ColIdentifierJO"
        Me.ColIdentifierJO.ReadOnly = True
        Me.ColIdentifierJO.Visible = False
        Me.ColIdentifierJO.Width = 88
        '
        'ColIdentifierWO
        '
        Me.ColIdentifierWO.HeaderText = "IdentifierWO"
        Me.ColIdentifierWO.Name = "ColIdentifierWO"
        Me.ColIdentifierWO.ReadOnly = True
        Me.ColIdentifierWO.Visible = False
        Me.ColIdentifierWO.Width = 91
        '
        'dgvItemSOB
        '
        Me.dgvItemSOB.AllowUserToAddRows = False
        Me.dgvItemSOB.AllowUserToDeleteRows = False
        Me.dgvItemSOB.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItemSOB.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColSOB_SONo, Me.ColSOB_CustPO, Me.ColSOB_Date, Me.ColSOB_CustomerCode, Me.ColSOB_CustomerName, Me.ColSOB_OrderBalance, Me.ColSOB_PromiseToDeliver, Me.ColSOB_PrimarySO, Me.ColSOB_IdentifierSOItem, Me.ColSOB_PrimaryUnitMulti})
        Me.dgvItemSOB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItemSOB.Location = New System.Drawing.Point(3, 16)
        Me.dgvItemSOB.Name = "dgvItemSOB"
        Me.dgvItemSOB.ReadOnly = True
        Me.dgvItemSOB.RowTemplate.Height = 24
        Me.dgvItemSOB.Size = New System.Drawing.Size(549, 461)
        Me.dgvItemSOB.TabIndex = 0
        Me.dgvItemSOB.TabStop = False
        '
        'ColSOB_SONo
        '
        Me.ColSOB_SONo.HeaderText = "SO Number"
        Me.ColSOB_SONo.Name = "ColSOB_SONo"
        Me.ColSOB_SONo.ReadOnly = True
        Me.ColSOB_SONo.Width = 87
        '
        'ColSOB_CustPO
        '
        Me.ColSOB_CustPO.HeaderText = "Customer PO Number"
        Me.ColSOB_CustPO.Name = "ColSOB_CustPO"
        Me.ColSOB_CustPO.ReadOnly = True
        Me.ColSOB_CustPO.Width = 134
        '
        'ColSOB_Date
        '
        Me.ColSOB_Date.HeaderText = "Date"
        Me.ColSOB_Date.Name = "ColSOB_Date"
        Me.ColSOB_Date.ReadOnly = True
        Me.ColSOB_Date.Width = 55
        '
        'ColSOB_CustomerCode
        '
        Me.ColSOB_CustomerCode.HeaderText = "Customer Code"
        Me.ColSOB_CustomerCode.Name = "ColSOB_CustomerCode"
        Me.ColSOB_CustomerCode.ReadOnly = True
        Me.ColSOB_CustomerCode.Width = 104
        '
        'ColSOB_CustomerName
        '
        Me.ColSOB_CustomerName.HeaderText = "Customer Name"
        Me.ColSOB_CustomerName.Name = "ColSOB_CustomerName"
        Me.ColSOB_CustomerName.ReadOnly = True
        Me.ColSOB_CustomerName.Width = 107
        '
        'ColSOB_OrderBalance
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSOB_OrderBalance.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColSOB_OrderBalance.HeaderText = "Order Balance"
        Me.ColSOB_OrderBalance.Name = "ColSOB_OrderBalance"
        Me.ColSOB_OrderBalance.ReadOnly = True
        Me.ColSOB_OrderBalance.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        '
        'ColSOB_PromiseToDeliver
        '
        Me.ColSOB_PromiseToDeliver.HeaderText = "Promise To Deliver"
        Me.ColSOB_PromiseToDeliver.Name = "ColSOB_PromiseToDeliver"
        Me.ColSOB_PromiseToDeliver.ReadOnly = True
        Me.ColSOB_PromiseToDeliver.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        Me.ColSOB_PromiseToDeliver.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColSOB_PromiseToDeliver.Width = 121
        '
        'ColSOB_PrimarySO
        '
        Me.ColSOB_PrimarySO.HeaderText = "Primary SO"
        Me.ColSOB_PrimarySO.Name = "ColSOB_PrimarySO"
        Me.ColSOB_PrimarySO.ReadOnly = True
        Me.ColSOB_PrimarySO.Visible = False
        Me.ColSOB_PrimarySO.Width = 84
        '
        'ColSOB_IdentifierSOItem
        '
        Me.ColSOB_IdentifierSOItem.HeaderText = "Identifier SO Item"
        Me.ColSOB_IdentifierSOItem.Name = "ColSOB_IdentifierSOItem"
        Me.ColSOB_IdentifierSOItem.ReadOnly = True
        Me.ColSOB_IdentifierSOItem.Visible = False
        Me.ColSOB_IdentifierSOItem.Width = 113
        '
        'ColSOB_PrimaryUnitMulti
        '
        Me.ColSOB_PrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColSOB_PrimaryUnitMulti.Name = "ColSOB_PrimaryUnitMulti"
        Me.ColSOB_PrimaryUnitMulti.ReadOnly = True
        Me.ColSOB_PrimaryUnitMulti.Visible = False
        Me.ColSOB_PrimaryUnitMulti.Width = 113
        '
        'DgvPurcReq
        '
        Me.DgvPurcReq.AllowUserToAddRows = False
        Me.DgvPurcReq.AllowUserToDeleteRows = False
        Me.DgvPurcReq.AllowUserToResizeColumns = False
        Me.DgvPurcReq.AllowUserToResizeRows = False
        Me.DgvPurcReq.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvPurcReq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvPurcReq.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPurcReq_PriorityLevel, Me.ColPurcReq_PromiseToDeliver, Me.ColPurcReq_RefParty, Me.ColPurcReq_DocumentNo, Me.ColPurcReq_Date, Me.ColPurcReq_CodeNo, Me.ColPurcReq_Description, Me.ColPurcReq_QtyOB, Me.ColPurcReq_BalanceQtyUsing, Me.ColPurcReq_Remarks, Me.ColPurcReq_PickUpOrder, Me.ColPurcReq_PrimaryPickUpOrder, Me.ColPurcReq_PrimaryPR, Me.ColPurcReq_PrimaryPrItemDel})
        Me.DgvPurcReq.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvPurcReq.Location = New System.Drawing.Point(3, 16)
        Me.DgvPurcReq.Name = "DgvPurcReq"
        Me.DgvPurcReq.ReadOnly = True
        Me.DgvPurcReq.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvPurcReq.RowTemplate.Height = 24
        Me.DgvPurcReq.Size = New System.Drawing.Size(549, 461)
        Me.DgvPurcReq.TabIndex = 4
        Me.DgvPurcReq.TabStop = False
        '
        'ColPurcReq_PriorityLevel
        '
        Me.ColPurcReq_PriorityLevel.HeaderText = "Prority"
        Me.ColPurcReq_PriorityLevel.Name = "ColPurcReq_PriorityLevel"
        Me.ColPurcReq_PriorityLevel.ReadOnly = True
        Me.ColPurcReq_PriorityLevel.Width = 61
        '
        'ColPurcReq_PromiseToDeliver
        '
        Me.ColPurcReq_PromiseToDeliver.HeaderText = "Promise To Deliver"
        Me.ColPurcReq_PromiseToDeliver.Name = "ColPurcReq_PromiseToDeliver"
        Me.ColPurcReq_PromiseToDeliver.ReadOnly = True
        Me.ColPurcReq_PromiseToDeliver.Width = 121
        '
        'ColPurcReq_RefParty
        '
        Me.ColPurcReq_RefParty.HeaderText = "Ref. Party"
        Me.ColPurcReq_RefParty.Name = "ColPurcReq_RefParty"
        Me.ColPurcReq_RefParty.ReadOnly = True
        Me.ColPurcReq_RefParty.Width = 79
        '
        'ColPurcReq_DocumentNo
        '
        Me.ColPurcReq_DocumentNo.HeaderText = "Purch Request No"
        Me.ColPurcReq_DocumentNo.Name = "ColPurcReq_DocumentNo"
        Me.ColPurcReq_DocumentNo.ReadOnly = True
        Me.ColPurcReq_DocumentNo.Width = 120
        '
        'ColPurcReq_Date
        '
        Me.ColPurcReq_Date.HeaderText = "Date"
        Me.ColPurcReq_Date.Name = "ColPurcReq_Date"
        Me.ColPurcReq_Date.ReadOnly = True
        Me.ColPurcReq_Date.Width = 55
        '
        'ColPurcReq_CodeNo
        '
        Me.ColPurcReq_CodeNo.HeaderText = "Code No"
        Me.ColPurcReq_CodeNo.Name = "ColPurcReq_CodeNo"
        Me.ColPurcReq_CodeNo.ReadOnly = True
        Me.ColPurcReq_CodeNo.Visible = False
        Me.ColPurcReq_CodeNo.Width = 74
        '
        'ColPurcReq_Description
        '
        Me.ColPurcReq_Description.HeaderText = "Description"
        Me.ColPurcReq_Description.Name = "ColPurcReq_Description"
        Me.ColPurcReq_Description.ReadOnly = True
        Me.ColPurcReq_Description.Visible = False
        Me.ColPurcReq_Description.Width = 85
        '
        'ColPurcReq_QtyOB
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPurcReq_QtyOB.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColPurcReq_QtyOB.HeaderText = "Balance Qty"
        Me.ColPurcReq_QtyOB.Name = "ColPurcReq_QtyOB"
        Me.ColPurcReq_QtyOB.ReadOnly = True
        Me.ColPurcReq_QtyOB.Width = 90
        '
        'ColPurcReq_BalanceQtyUsing
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPurcReq_BalanceQtyUsing.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColPurcReq_BalanceQtyUsing.HeaderText = "Waiting Using"
        Me.ColPurcReq_BalanceQtyUsing.Name = "ColPurcReq_BalanceQtyUsing"
        Me.ColPurcReq_BalanceQtyUsing.ReadOnly = True
        Me.ColPurcReq_BalanceQtyUsing.Width = 98
        '
        'ColPurcReq_Remarks
        '
        Me.ColPurcReq_Remarks.HeaderText = "Remarks"
        Me.ColPurcReq_Remarks.Name = "ColPurcReq_Remarks"
        Me.ColPurcReq_Remarks.ReadOnly = True
        Me.ColPurcReq_Remarks.Width = 74
        '
        'ColPurcReq_PickUpOrder
        '
        Me.ColPurcReq_PickUpOrder.HeaderText = "Pick Up Order No"
        Me.ColPurcReq_PickUpOrder.Name = "ColPurcReq_PickUpOrder"
        Me.ColPurcReq_PickUpOrder.ReadOnly = True
        Me.ColPurcReq_PickUpOrder.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPurcReq_PickUpOrder.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColPurcReq_PickUpOrder.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColPurcReq_PickUpOrder.Width = 116
        '
        'ColPurcReq_PrimaryPickUpOrder
        '
        Me.ColPurcReq_PrimaryPickUpOrder.HeaderText = "Primary Pick Up Order"
        Me.ColPurcReq_PrimaryPickUpOrder.Name = "ColPurcReq_PrimaryPickUpOrder"
        Me.ColPurcReq_PrimaryPickUpOrder.ReadOnly = True
        Me.ColPurcReq_PrimaryPickUpOrder.Visible = False
        Me.ColPurcReq_PrimaryPickUpOrder.Width = 136
        '
        'ColPurcReq_PrimaryPR
        '
        Me.ColPurcReq_PrimaryPR.HeaderText = "Primary PR"
        Me.ColPurcReq_PrimaryPR.Name = "ColPurcReq_PrimaryPR"
        Me.ColPurcReq_PrimaryPR.ReadOnly = True
        Me.ColPurcReq_PrimaryPR.Visible = False
        Me.ColPurcReq_PrimaryPR.Width = 84
        '
        'ColPurcReq_PrimaryPrItemDel
        '
        Me.ColPurcReq_PrimaryPrItemDel.HeaderText = "Primary Pr Item Del"
        Me.ColPurcReq_PrimaryPrItemDel.Name = "ColPurcReq_PrimaryPrItemDel"
        Me.ColPurcReq_PrimaryPrItemDel.ReadOnly = True
        Me.ColPurcReq_PrimaryPrItemDel.Visible = False
        Me.ColPurcReq_PrimaryPrItemDel.Width = 121
        '
        'llbPurcReqBalance
        '
        Me.llbPurcReqBalance.AutoSize = True
        Me.llbPurcReqBalance.Location = New System.Drawing.Point(12, 363)
        Me.llbPurcReqBalance.Name = "llbPurcReqBalance"
        Me.llbPurcReqBalance.Size = New System.Drawing.Size(137, 13)
        Me.llbPurcReqBalance.TabIndex = 48
        Me.llbPurcReqBalance.TabStop = True
        Me.llbPurcReqBalance.Text = "Purchase Request Balance"
        '
        'LblUnit8
        '
        Me.LblUnit8.AutoSize = True
        Me.LblUnit8.Location = New System.Drawing.Point(240, 363)
        Me.LblUnit8.Name = "LblUnit8"
        Me.LblUnit8.Size = New System.Drawing.Size(31, 13)
        Me.LblUnit8.TabIndex = 49
        Me.LblUnit8.Text = "Units"
        '
        'TxtPurcReqBalance
        '
        Me.TxtPurcReqBalance.Location = New System.Drawing.Point(164, 362)
        Me.TxtPurcReqBalance.Name = "TxtPurcReqBalance"
        Me.TxtPurcReqBalance.ReadOnly = True
        Me.TxtPurcReqBalance.Size = New System.Drawing.Size(71, 20)
        Me.TxtPurcReqBalance.TabIndex = 47
        Me.TxtPurcReqBalance.TabStop = False
        Me.TxtPurcReqBalance.Text = "0"
        Me.TxtPurcReqBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ChkIncludeUnAuthPO
        '
        Me.ChkIncludeUnAuthPO.AutoSize = True
        Me.ChkIncludeUnAuthPO.Location = New System.Drawing.Point(15, 340)
        Me.ChkIncludeUnAuthPO.Name = "ChkIncludeUnAuthPO"
        Me.ChkIncludeUnAuthPO.Size = New System.Drawing.Size(145, 17)
        Me.ChkIncludeUnAuthPO.TabIndex = 50
        Me.ChkIncludeUnAuthPO.Text = "Include Unauthorized PO"
        Me.ChkIncludeUnAuthPO.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(12, 117)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(112, 13)
        Me.LinkLabel1.TabIndex = 52
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Raw Material Request"
        '
        'lblUnit9
        '
        Me.lblUnit9.AutoSize = True
        Me.lblUnit9.Location = New System.Drawing.Point(240, 117)
        Me.lblUnit9.Name = "lblUnit9"
        Me.lblUnit9.Size = New System.Drawing.Size(31, 13)
        Me.lblUnit9.TabIndex = 53
        Me.lblUnit9.Text = "Units"
        '
        'TxtRawMaterialRequest
        '
        Me.TxtRawMaterialRequest.Location = New System.Drawing.Point(164, 115)
        Me.TxtRawMaterialRequest.Name = "TxtRawMaterialRequest"
        Me.TxtRawMaterialRequest.ReadOnly = True
        Me.TxtRawMaterialRequest.Size = New System.Drawing.Size(71, 20)
        Me.TxtRawMaterialRequest.TabIndex = 51
        Me.TxtRawMaterialRequest.TabStop = False
        Me.TxtRawMaterialRequest.Text = "0"
        Me.TxtRawMaterialRequest.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LinkLabel2
        '
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(12, 266)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(146, 13)
        Me.LinkLabel2.TabIndex = 55
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Manufacturing Order Balance"
        '
        'lblUnit10
        '
        Me.lblUnit10.AutoSize = True
        Me.lblUnit10.Location = New System.Drawing.Point(240, 266)
        Me.lblUnit10.Name = "lblUnit10"
        Me.lblUnit10.Size = New System.Drawing.Size(31, 13)
        Me.lblUnit10.TabIndex = 56
        Me.lblUnit10.Text = "Units"
        '
        'TxtMOB
        '
        Me.TxtMOB.Location = New System.Drawing.Point(164, 263)
        Me.TxtMOB.Name = "TxtMOB"
        Me.TxtMOB.ReadOnly = True
        Me.TxtMOB.Size = New System.Drawing.Size(71, 20)
        Me.TxtMOB.TabIndex = 54
        Me.TxtMOB.TabStop = False
        Me.TxtMOB.Text = "0"
        Me.TxtMOB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'chkShowWONo
        '
        Me.chkShowWONo.AutoSize = True
        Me.chkShowWONo.Location = New System.Drawing.Point(164, 138)
        Me.chkShowWONo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chkShowWONo.Name = "chkShowWONo"
        Me.chkShowWONo.Size = New System.Drawing.Size(92, 17)
        Me.chkShowWONo.TabIndex = 12
        Me.chkShowWONo.Text = "Show WO No"
        Me.chkShowWONo.UseVisualStyleBackColor = True
        '
        'FrmMasterKondisiStock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(859, 484)
        Me.Controls.Add(Me.chkShowWONo)
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.lblUnit10)
        Me.Controls.Add(Me.TxtMOB)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.lblUnit9)
        Me.Controls.Add(Me.TxtRawMaterialRequest)
        Me.Controls.Add(Me.ChkIncludeUnAuthPO)
        Me.Controls.Add(Me.llbPurcReqBalance)
        Me.Controls.Add(Me.LblUnit8)
        Me.Controls.Add(Me.TxtPurcReqBalance)
        Me.Controls.Add(Me.llbSOB)
        Me.Controls.Add(Me.LblUnit7)
        Me.Controls.Add(Me.TxtSOB)
        Me.Controls.Add(Me.llbPOB)
        Me.Controls.Add(Me.LblUnit6)
        Me.Controls.Add(Me.TxtPOB)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.cboUnitPrice)
        Me.Controls.Add(Me.TxtUnitPrice)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblUnitAvailable)
        Me.Controls.Add(Me.txtAvailable)
        Me.Controls.Add(Me.lblUnitPendingDO)
        Me.Controls.Add(Me.txtPendingDO)
        Me.Controls.Add(Me.lblPendingDO)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.txtPrimaryCodeNo)
        Me.Controls.Add(Me.lblUnit5)
        Me.Controls.Add(Me.txtBookingGIT)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.llbGIT)
        Me.Controls.Add(Me.llbBookingStockInv)
        Me.Controls.Add(Me.lblUnit4)
        Me.Controls.Add(Me.lblUnit3)
        Me.Controls.Add(Me.lblUnit2)
        Me.Controls.Add(Me.lblUnit1)
        Me.Controls.Add(Me.txtGIT)
        Me.Controls.Add(Me.txtQtyOnHand)
        Me.Controls.Add(Me.txtBookingInv)
        Me.Controls.Add(Me.txtFreeStock)
        Me.Controls.Add(Me.btnFindProduct)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtCodeNo)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.llbQtyOnHand)
        Me.Controls.Add(Me.llbBookingGIT)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(875, 463)
        Me.Name = "FrmMasterKondisiStock"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product Information"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemPendingDO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Stock_detail_viewDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemPOB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboDetail.ResumeLayout(False)
        CType(Me.dgvDataRMReq, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvDataMOB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvItemSOB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DgvPurcReq, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnFindProduct As System.Windows.Forms.Button
    Friend WithEvents txtFreeStock As System.Windows.Forms.TextBox
    Friend WithEvents txtQtyOnHand As System.Windows.Forms.TextBox
    Friend WithEvents txtBookingInv As System.Windows.Forms.TextBox
    Friend WithEvents txtGIT As System.Windows.Forms.TextBox
    Friend WithEvents lblUnit1 As System.Windows.Forms.Label
    Friend WithEvents lblUnit2 As System.Windows.Forms.Label
    Friend WithEvents lblUnit4 As System.Windows.Forms.Label
    Friend WithEvents lblUnit3 As System.Windows.Forms.Label
    Friend WithEvents llbBookingStockInv As System.Windows.Forms.LinkLabel
    Friend WithEvents llbGIT As System.Windows.Forms.LinkLabel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents llbBookingGIT As System.Windows.Forms.LinkLabel
    Friend WithEvents lblUnit5 As System.Windows.Forms.Label
    Friend WithEvents txtBookingGIT As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents llbQtyOnHand As System.Windows.Forms.LinkLabel
    Friend WithEvents lblUnitPendingDO As System.Windows.Forms.Label
    Friend WithEvents txtPendingDO As System.Windows.Forms.TextBox
    Friend WithEvents lblPendingDO As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblUnitAvailable As System.Windows.Forms.Label
    Friend WithEvents txtAvailable As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtUnitPrice As System.Windows.Forms.TextBox
    Friend WithEvents cboUnitPrice As System.Windows.Forms.ComboBox
    Public WithEvents txtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents llbPOB As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUnit6 As System.Windows.Forms.Label
    Friend WithEvents TxtPOB As System.Windows.Forms.TextBox
    Friend WithEvents llbSOB As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUnit7 As System.Windows.Forms.Label
    Friend WithEvents TxtSOB As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSpec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents dgvItemPendingDO As System.Windows.Forms.DataGridView
    Friend WithEvents Stock_detail_viewDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents dgvItemPOB As System.Windows.Forms.DataGridView
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents dgvItemSOB As System.Windows.Forms.DataGridView
    Friend WithEvents llbPurcReqBalance As System.Windows.Forms.LinkLabel
    Friend WithEvents LblUnit8 As System.Windows.Forms.Label
    Friend WithEvents TxtPurcReqBalance As System.Windows.Forms.TextBox
    Friend WithEvents DgvPurcReq As System.Windows.Forms.DataGridView
    Friend WithEvents ChkIncludeUnAuthPO As System.Windows.Forms.CheckBox
    Friend WithEvents ColPurcReq_PriorityLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_PromiseToDeliver As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_RefParty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_DocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_Description As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_QtyOB As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_BalanceQtyUsing As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_Remarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_PickUpOrder As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColPurcReq_PrimaryPickUpOrder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_PrimaryPR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurcReq_PrimaryPrItemDel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_SONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_CustPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_CustomerCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_CustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_OrderBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_PromiseToDeliver As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColSOB_PrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_IdentifierSOItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSOB_PrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents lblUnit9 As System.Windows.Forms.Label
    Friend WithEvents TxtRawMaterialRequest As System.Windows.Forms.TextBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents lblUnit10 As System.Windows.Forms.Label
    Friend WithEvents TxtMOB As System.Windows.Forms.TextBox
    Friend WithEvents dgvDataMOB As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMODate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTargetFinishDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProgress As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColIdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgvDataRMReq As System.Windows.Forms.DataGridView
    Friend WithEvents ColRMReq_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_MONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_MODate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_JONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_WONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_Process As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_BalanceQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_UnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_IsWIP As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColRMReq_IdentifierJO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRMReq_IdentifierWO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkShowWONo As System.Windows.Forms.CheckBox
End Class
