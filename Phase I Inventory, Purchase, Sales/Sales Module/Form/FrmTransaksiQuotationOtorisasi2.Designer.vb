<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiQuotationOtorisasi2
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
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbExit = New System.Windows.Forms.ToolStripButton
        Me.tsbAuthorize = New System.Windows.Forms.ToolStripButton
        Me.TsbPrint = New System.Windows.Forms.ToolStripButton
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.txtTaxAmount = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.chkTax = New System.Windows.Forms.CheckBox
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.txtTax = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtSubTotal = New System.Windows.Forms.TextBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtContactPerson = New System.Windows.Forms.TextBox
        Me.btnCLose = New System.Windows.Forms.Button
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.llbAuthorizationDetail = New System.Windows.Forms.LinkLabel
        Me.txtAutohirizationStatus = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.cboPaymentDueDateReference = New System.Windows.Forms.ComboBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupplierCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriceByMultiUnit_Str = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryCont = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDown = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDownPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUpPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUseSecondary = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitsMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDownFromPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUpFromPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriceByMultiUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtNoNPWP = New System.Windows.Forms.TextBox
        Me.TxtInquiryNo = New System.Windows.Forms.TextBox
        Me.LlbAddress = New System.Windows.Forms.LinkLabel
        Me.txtPaymentDueDate = New System.Windows.Forms.TextBox
        Me.TxtStockInNo = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.txtCustomerAddIdentifier = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtPaymentTerm = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.txtPhoneNo = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboCustomerAddress = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.LlbViewAddressCustomerShipTo = New System.Windows.Forms.LinkLabel
        Me.TxtIdentifierAddressCustomerShipTo = New System.Windows.Forms.TextBox
        Me.CboAddressCustomerShipTo = New System.Windows.Forms.ComboBox
        Me.TxtCustomerDescShipTo = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCustomerShipTo = New System.Windows.Forms.TextBox
        Me.RboCustomerShipTo = New System.Windows.Forms.RadioButton
        Me.TxtCustomerCodeShipTo = New System.Windows.Forms.TextBox
        Me.LlbViewAddressShipTo = New System.Windows.Forms.LinkLabel
        Me.CboAddressShipTo = New System.Windows.Forms.ComboBox
        Me.TxtIdentifierAddressShipTo = New System.Windows.Forms.TextBox
        Me.RboAddressShipTo = New System.Windows.Forms.RadioButton
        Me.TxtDPFromPercent = New System.Windows.Forms.TextBox
        Me.TxtDPAmount = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.chkUseDP = New System.Windows.Forms.CheckBox
        Me.TxtDPPercent = New System.Windows.Forms.TextBox
        Me.txtTermsNCondition = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.TabControl2 = New System.Windows.Forms.TabControl
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.DgvOtorisasi = New System.Windows.Forms.DataGridView
        Me.ColDgvAuth_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvAuth_Authorize = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDgvAuth_QuoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvAuth_CustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbUncheckAll_Athorize = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll_Authorize = New System.Windows.Forms.LinkLabel
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvCancelOtorisasi = New System.Windows.Forms.DataGridView
        Me.ColDgvUnAuth_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvUnAuth_Cancel = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDgvUnAuth_QuoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDgvUnAuth_CustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LlbUncheckAll_Cancel = New System.Windows.Forms.LinkLabel
        Me.LlbCheckAll_Cancel = New System.Windows.Forms.LinkLabel
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.DgvAlreadyAuthorize = New System.Windows.Forms.DataGridView
        Me.ColAA_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAA_Print = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAA_QuoNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAA_CustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAA_Status = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.TxtNoQuotation = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.BtnFindQuo = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.DtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.DtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.ChkQuoDate = New System.Windows.Forms.CheckBox
        Me.DgvUser = New System.Windows.Forms.DataGridView
        Me.ColCreatedBy_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCreatedBy_Name = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CboStatusQuotation = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtPrimaryQuo = New System.Windows.Forms.TextBox
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer
        Me.ToolStrip2 = New System.Windows.Forms.ToolStrip
        Me.tsbRemove = New System.Windows.Forms.ToolStripButton
        Me.tsbShowRemoved = New System.Windows.Forms.ToolStripButton
        Me.BgwUser = New System.ComponentModel.BackgroundWorker
        Me.ToolStrip1.SuspendLayout()
        Me.StatusStrip2.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.DgvOtorisasi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvCancelOtorisasi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        CType(Me.DgvAlreadyAuthorize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.ToolStrip2.SuspendLayout()
        Me.SuspendLayout()
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(248, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbExit, Me.tsbAuthorize, Me.TsbPrint})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip1.Size = New System.Drawing.Size(441, 33)
        Me.ToolStrip1.TabIndex = 59
        '
        'tsbExit
        '
        Me.tsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExit.Name = "tsbExit"
        Me.tsbExit.Size = New System.Drawing.Size(29, 30)
        Me.tsbExit.Text = "E&xit"
        Me.tsbExit.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tsbAuthorize
        '
        Me.tsbAuthorize.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAuthorize.Name = "tsbAuthorize"
        Me.tsbAuthorize.Size = New System.Drawing.Size(62, 30)
        Me.tsbAuthorize.Text = "Authorize"
        Me.tsbAuthorize.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'TsbPrint
        '
        Me.TsbPrint.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsbPrint.Name = "TsbPrint"
        Me.TsbPrint.Size = New System.Drawing.Size(36, 30)
        Me.TsbPrint.Text = "Print"
        Me.TsbPrint.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(248, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(248, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'txtTaxAmount
        '
        Me.txtTaxAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxAmount.Location = New System.Drawing.Point(181, 106)
        Me.txtTaxAmount.MaxLength = 50
        Me.txtTaxAmount.Name = "txtTaxAmount"
        Me.txtTaxAmount.ReadOnly = True
        Me.txtTaxAmount.Size = New System.Drawing.Size(108, 20)
        Me.txtTaxAmount.TabIndex = 27
        Me.txtTaxAmount.Text = "0"
        Me.txtTaxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(160, 107)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(15, 13)
        Me.Label25.TabIndex = 152
        Me.Label25.Text = "%"
        '
        'chkTax
        '
        Me.chkTax.AutoSize = True
        Me.chkTax.Enabled = False
        Me.chkTax.Location = New System.Drawing.Point(6, 107)
        Me.chkTax.Name = "chkTax"
        Me.chkTax.Size = New System.Drawing.Size(44, 17)
        Me.chkTax.TabIndex = 25
        Me.chkTax.Text = "Tax"
        Me.chkTax.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 498)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(1010, 22)
        Me.StatusStrip2.TabIndex = 58
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(248, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'txtTax
        '
        Me.txtTax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTax.Location = New System.Drawing.Point(117, 105)
        Me.txtTax.MaxLength = 50
        Me.txtTax.Name = "txtTax"
        Me.txtTax.ReadOnly = True
        Me.txtTax.Size = New System.Drawing.Size(36, 20)
        Me.txtTax.TabIndex = 26
        Me.txtTax.Text = "0"
        Me.txtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(3, 132)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(31, 13)
        Me.Label24.TabIndex = 150
        Me.Label24.Text = "Total"
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotal.Location = New System.Drawing.Point(117, 128)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(173, 20)
        Me.TxtTotal.TabIndex = 28
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(3, 81)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(53, 13)
        Me.Label15.TabIndex = 148
        Me.Label15.Text = "Sub Total"
        '
        'txtSubTotal
        '
        Me.txtSubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSubTotal.Location = New System.Drawing.Point(117, 77)
        Me.txtSubTotal.MaxLength = 50
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.ReadOnly = True
        Me.txtSubTotal.Size = New System.Drawing.Size(173, 20)
        Me.txtSubTotal.TabIndex = 24
        Me.txtSubTotal.Text = "0"
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(117, 153)
        Me.TxtRemark.MaxLength = 50
        Me.TxtRemark.Multiline = True
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.ReadOnly = True
        Me.TxtRemark.Size = New System.Drawing.Size(173, 19)
        Me.TxtRemark.TabIndex = 29
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(2, 155)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Remarks"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 83)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 13)
        Me.Label7.TabIndex = 138
        Me.Label7.Text = "Contact Person"
        '
        'txtContactPerson
        '
        Me.txtContactPerson.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtContactPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtContactPerson.Location = New System.Drawing.Point(107, 77)
        Me.txtContactPerson.MaxLength = 150
        Me.txtContactPerson.Name = "txtContactPerson"
        Me.txtContactPerson.ReadOnly = True
        Me.txtContactPerson.Size = New System.Drawing.Size(232, 20)
        Me.txtContactPerson.TabIndex = 6
        '
        'btnCLose
        '
        Me.btnCLose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCLose.Location = New System.Drawing.Point(83, 275)
        Me.btnCLose.Name = "btnCLose"
        Me.btnCLose.Size = New System.Drawing.Size(122, 42)
        Me.btnCLose.TabIndex = 155
        Me.btnCLose.Text = "Button1"
        Me.btnCLose.UseVisualStyleBackColor = True
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Customer"
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(143, 27)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCustomer.TabIndex = 133
        Me.txtPrimaryCustomer.TabStop = False
        Me.txtPrimaryCustomer.Visible = False
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(107, 27)
        Me.txtCustomerCode.MaxLength = 50
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.ReadOnly = True
        Me.txtCustomerCode.Size = New System.Drawing.Size(65, 20)
        Me.txtCustomerCode.TabIndex = 2
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerName.Location = New System.Drawing.Point(185, 27)
        Me.TxtCustomerName.MaxLength = 150
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(154, 20)
        Me.TxtCustomerName.TabIndex = 3
        Me.TxtCustomerName.TabStop = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryCurrency.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(108, 225)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(29, 20)
        Me.txtPrimaryCurrency.TabIndex = 126
        Me.txtPrimaryCurrency.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCurrency.Visible = False
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(6, 229)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(77, 13)
        Me.Label22.TabIndex = 125
        Me.Label22.Text = "Currency Code"
        '
        'llbAuthorizationDetail
        '
        Me.llbAuthorizationDetail.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.llbAuthorizationDetail.AutoSize = True
        Me.llbAuthorizationDetail.Location = New System.Drawing.Point(282, 152)
        Me.llbAuthorizationDetail.Name = "llbAuthorizationDetail"
        Me.llbAuthorizationDetail.Size = New System.Drawing.Size(57, 13)
        Me.llbAuthorizationDetail.TabIndex = 10
        Me.llbAuthorizationDetail.TabStop = True
        Me.llbAuthorizationDetail.Text = "view detail"
        '
        'txtAutohirizationStatus
        '
        Me.txtAutohirizationStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAutohirizationStatus.Location = New System.Drawing.Point(107, 149)
        Me.txtAutohirizationStatus.MaxLength = 50
        Me.txtAutohirizationStatus.Name = "txtAutohirizationStatus"
        Me.txtAutohirizationStatus.ReadOnly = True
        Me.txtAutohirizationStatus.Size = New System.Drawing.Size(167, 20)
        Me.txtAutohirizationStatus.TabIndex = 9
        Me.txtAutohirizationStatus.TabStop = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(6, 155)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(101, 13)
        Me.Label21.TabIndex = 120
        Me.Label21.Text = "Authorization Status"
        '
        'cboPaymentDueDateReference
        '
        Me.cboPaymentDueDateReference.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboPaymentDueDateReference.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaymentDueDateReference.Enabled = False
        Me.cboPaymentDueDateReference.FormattingEnabled = True
        Me.cboPaymentDueDateReference.Items.AddRange(New Object() {"PO Date", "Invoice Date", "Bill Of Lading"})
        Me.cboPaymentDueDateReference.Location = New System.Drawing.Point(253, 195)
        Me.cboPaymentDueDateReference.Name = "cboPaymentDueDateReference"
        Me.cboPaymentDueDateReference.Size = New System.Drawing.Size(86, 21)
        Me.cboPaymentDueDateReference.TabIndex = 13
        '
        'gboDetail
        '
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Controls.Add(Me.btnCLose)
        Me.gboDetail.Controls.Add(Me.Button2)
        Me.gboDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gboDetail.Location = New System.Drawing.Point(0, 0)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(650, 449)
        Me.gboDetail.TabIndex = 57
        Me.gboDetail.TabStop = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvItem.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColSupplierCodeNo, Me.ColSuppDesc, Me.ColQtyMulti, Me.ColUnitsMulti, Me.ColRatioDesc, Me.ColQty, Me.ColUnit, Me.ColUnitPrice, Me.ColPriceByMultiUnit_Str, Me.ColDiscount, Me.ColDiscAmount, Me.ColNettPrice, Me.ColAmount, Me.ColDeliveryCont, Me.ColDeliveryDate, Me.ColRemark, Me.ColToleranceDown, Me.ColToleranceDownPerc, Me.ColToleranceUp, Me.ColToleranceUpPerc, Me.ColPrimaryCodeNo, Me.ColPrimaryUnits, Me.ColUseSecondary, Me.ColDiscFromPerct, Me.ColIdentifier, Me.ColPrimaryUnitsMulti, Me.ColRatioDefault, Me.ColRatioReference, Me.ColToleranceDownFromPerc, Me.ColToleranceUpFromPerc, Me.ColPriceByMultiUnit})
        Me.dgvItem.Location = New System.Drawing.Point(6, 269)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(641, 173)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 169
        Me.dgvItem.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.Width = 85
        '
        'ColSupplierCodeNo
        '
        Me.ColSupplierCodeNo.HeaderText = "Customer's Code No"
        Me.ColSupplierCodeNo.Name = "ColSupplierCodeNo"
        Me.ColSupplierCodeNo.Width = 128
        '
        'ColSuppDesc
        '
        Me.ColSuppDesc.HeaderText = "Customer's Description"
        Me.ColSuppDesc.Name = "ColSuppDesc"
        Me.ColSuppDesc.Width = 139
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColUnitsMulti
        '
        Me.ColUnitsMulti.HeaderText = "Unit"
        Me.ColUnitsMulti.Name = "ColUnitsMulti"
        Me.ColUnitsMulti.Width = 51
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRatioDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRatioDesc.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.Width = 57
        '
        'ColQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQty.HeaderText = "Qty Default"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.Width = 85
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units Default"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 93
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.Width = 78
        '
        'ColPriceByMultiUnit_Str
        '
        Me.ColPriceByMultiUnit_Str.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColPriceByMultiUnit_Str.HeaderText = ""
        Me.ColPriceByMultiUnit_Str.Name = "ColPriceByMultiUnit_Str"
        Me.ColPriceByMultiUnit_Str.ReadOnly = True
        Me.ColPriceByMultiUnit_Str.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPriceByMultiUnit_Str.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColPriceByMultiUnit_Str.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColPriceByMultiUnit_Str.Width = 19
        '
        'ColDiscount
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.Width = 91
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColDiscAmount.HeaderText = "Disc Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.Width = 92
        '
        'ColNettPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.Width = 79
        '
        'ColAmount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
        '
        'ColDeliveryCont
        '
        Me.ColDeliveryCont.HeaderText = "Delivery Date Control"
        Me.ColDeliveryCont.Name = "ColDeliveryCont"
        Me.ColDeliveryCont.Width = 113
        '
        'ColDeliveryDate
        '
        Me.ColDeliveryDate.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.HeaderText = "Delivery Date"
        Me.ColDeliveryDate.Name = "ColDeliveryDate"
        Me.ColDeliveryDate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDeliveryDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColDeliveryDate.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColDeliveryDate.Width = 96
        '
        'ColRemark
        '
        Me.ColRemark.HeaderText = "Remarks"
        Me.ColRemark.Name = "ColRemark"
        Me.ColRemark.Width = 74
        '
        'ColToleranceDown
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceDown.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColToleranceDown.HeaderText = "Tolerance Down"
        Me.ColToleranceDown.Name = "ColToleranceDown"
        Me.ColToleranceDown.Width = 111
        '
        'ColToleranceDownPerc
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceDownPerc.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColToleranceDownPerc.HeaderText = "Tolerance Down (%)"
        Me.ColToleranceDownPerc.Name = "ColToleranceDownPerc"
        Me.ColToleranceDownPerc.Width = 128
        '
        'ColToleranceUp
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceUp.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColToleranceUp.HeaderText = "Tolerance Up"
        Me.ColToleranceUp.Name = "ColToleranceUp"
        Me.ColToleranceUp.Width = 97
        '
        'ColToleranceUpPerc
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceUpPerc.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColToleranceUpPerc.HeaderText = "Tolerance Up (%)"
        Me.ColToleranceUpPerc.Name = "ColToleranceUpPerc"
        Me.ColToleranceUpPerc.Width = 114
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 132
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 113
        '
        'ColUseSecondary
        '
        Me.ColUseSecondary.HeaderText = "ColUseSecondary"
        Me.ColUseSecondary.Name = "ColUseSecondary"
        Me.ColUseSecondary.Visible = False
        Me.ColUseSecondary.Width = 127
        '
        'ColDiscFromPerct
        '
        Me.ColDiscFromPerct.HeaderText = "Disc From Perct"
        Me.ColDiscFromPerct.Name = "ColDiscFromPerct"
        Me.ColDiscFromPerct.ReadOnly = True
        Me.ColDiscFromPerct.Visible = False
        Me.ColDiscFromPerct.Width = 133
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.ReadOnly = True
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 87
        '
        'ColPrimaryUnitsMulti
        '
        Me.ColPrimaryUnitsMulti.HeaderText = "Primary Units Multi"
        Me.ColPrimaryUnitsMulti.Name = "ColPrimaryUnitsMulti"
        Me.ColPrimaryUnitsMulti.Visible = False
        Me.ColPrimaryUnitsMulti.Width = 150
        '
        'ColRatioDefault
        '
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 115
        '
        'ColRatioReference
        '
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 136
        '
        'ColToleranceDownFromPerc
        '
        Me.ColToleranceDownFromPerc.HeaderText = "Tolerance Down From Perc"
        Me.ColToleranceDownFromPerc.Name = "ColToleranceDownFromPerc"
        Me.ColToleranceDownFromPerc.Visible = False
        Me.ColToleranceDownFromPerc.Width = 205
        '
        'ColToleranceUpFromPerc
        '
        Me.ColToleranceUpFromPerc.HeaderText = "Tolerance Up From Perc"
        Me.ColToleranceUpFromPerc.Name = "ColToleranceUpFromPerc"
        Me.ColToleranceUpFromPerc.Visible = False
        Me.ColToleranceUpFromPerc.Width = 188
        '
        'ColPriceByMultiUnit
        '
        Me.ColPriceByMultiUnit.HeaderText = "Price by Multi Unit"
        Me.ColPriceByMultiUnit.Name = "ColPriceByMultiUnit"
        Me.ColPriceByMultiUnit.ReadOnly = True
        Me.ColPriceByMultiUnit.Visible = False
        Me.ColPriceByMultiUnit.Width = 146
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(7, 11)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCurrencyCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryStockIn)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label11)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtNoNPWP)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtInquiryNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LlbAddress)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPaymentDueDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtStockInNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label18)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label20)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerAddIdentifier)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboPaymentDueDateReference)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label23)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPaymentTerm)
        Me.SplitContainer1.Panel1.Controls.Add(Me.dtpDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label22)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPhoneNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryCurrency)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label21)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtContactPerson)
        Me.SplitContainer1.Panel1.Controls.Add(Me.llbAuthorizationDetail)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.cboCustomerAddress)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtPrimaryCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label17)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.txtAutohirizationStatus)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.LlbViewAddressCustomerShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtIdentifierAddressCustomerShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboAddressCustomerShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCustomerDescShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryCustomerShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RboCustomerShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtCustomerCodeShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.LlbViewAddressShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.CboAddressShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtIdentifierAddressShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RboAddressShipTo)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDPFromPercent)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDPAmount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label13)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkUseDP)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtDPPercent)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtTermsNCondition)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtTaxAmount)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label25)
        Me.SplitContainer1.Panel2.Controls.Add(Me.chkTax)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtTax)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label24)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtTotal)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label15)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.txtSubTotal)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemark)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label12)
        Me.SplitContainer1.Size = New System.Drawing.Size(638, 253)
        Me.SplitContainer1.SplitterDistance = 342
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 168
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(171, 228)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(75, 13)
        Me.Label8.TabIndex = 169
        Me.Label8.Text = "Payment Term"
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrencyCode.Location = New System.Drawing.Point(107, 225)
        Me.txtCurrencyCode.MaxLength = 50
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.ReadOnly = True
        Me.txtCurrencyCode.Size = New System.Drawing.Size(58, 20)
        Me.txtCurrencyCode.TabIndex = 14
        Me.txtCurrencyCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(125, 3)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryStockIn.TabIndex = 47
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(6, 132)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(57, 13)
        Me.Label11.TabIndex = 168
        Me.Label11.Text = "NPWP No"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtNoNPWP
        '
        Me.txtNoNPWP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNoNPWP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNoNPWP.Location = New System.Drawing.Point(107, 125)
        Me.txtNoNPWP.MaxLength = 50
        Me.txtNoNPWP.Name = "txtNoNPWP"
        Me.txtNoNPWP.ReadOnly = True
        Me.txtNoNPWP.Size = New System.Drawing.Size(232, 20)
        Me.txtNoNPWP.TabIndex = 8
        '
        'TxtInquiryNo
        '
        Me.TxtInquiryNo.AcceptsTab = True
        Me.TxtInquiryNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtInquiryNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtInquiryNo.Location = New System.Drawing.Point(107, 173)
        Me.TxtInquiryNo.MaxLength = 50
        Me.TxtInquiryNo.Name = "TxtInquiryNo"
        Me.TxtInquiryNo.ReadOnly = True
        Me.TxtInquiryNo.Size = New System.Drawing.Size(232, 20)
        Me.TxtInquiryNo.TabIndex = 11
        '
        'LlbAddress
        '
        Me.LlbAddress.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbAddress.AutoSize = True
        Me.LlbAddress.Location = New System.Drawing.Point(282, 54)
        Me.LlbAddress.Name = "LlbAddress"
        Me.LlbAddress.Size = New System.Drawing.Size(57, 13)
        Me.LlbAddress.TabIndex = 5
        Me.LlbAddress.TabStop = True
        Me.LlbAddress.Text = "view detail"
        '
        'txtPaymentDueDate
        '
        Me.txtPaymentDueDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPaymentDueDate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaymentDueDate.Location = New System.Drawing.Point(107, 197)
        Me.txtPaymentDueDate.MaxLength = 50
        Me.txtPaymentDueDate.Name = "txtPaymentDueDate"
        Me.txtPaymentDueDate.ReadOnly = True
        Me.txtPaymentDueDate.Size = New System.Drawing.Size(58, 20)
        Me.txtPaymentDueDate.TabIndex = 12
        Me.txtPaymentDueDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtStockInNo
        '
        Me.TxtStockInNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStockInNo.Location = New System.Drawing.Point(107, 4)
        Me.TxtStockInNo.MaxLength = 50
        Me.TxtStockInNo.Name = "TxtStockInNo"
        Me.TxtStockInNo.ReadOnly = True
        Me.TxtStockInNo.Size = New System.Drawing.Size(135, 20)
        Me.TxtStockInNo.TabIndex = 0
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(6, 203)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(97, 13)
        Me.Label18.TabIndex = 117
        Me.Label18.Text = "Payment Due Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Quotation No"
        '
        'Label20
        '
        Me.Label20.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(171, 199)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(55, 13)
        Me.Label20.TabIndex = 118
        Me.Label20.Text = "days from "
        '
        'txtCustomerAddIdentifier
        '
        Me.txtCustomerAddIdentifier.AcceptsTab = True
        Me.txtCustomerAddIdentifier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerAddIdentifier.Location = New System.Drawing.Point(76, 53)
        Me.txtCustomerAddIdentifier.MaxLength = 50
        Me.txtCustomerAddIdentifier.Name = "txtCustomerAddIdentifier"
        Me.txtCustomerAddIdentifier.Size = New System.Drawing.Size(28, 20)
        Me.txtCustomerAddIdentifier.TabIndex = 163
        Me.txtCustomerAddIdentifier.Visible = False
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(6, 109)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(55, 13)
        Me.Label23.TabIndex = 141
        Me.Label23.Text = "Phone No"
        '
        'txtPaymentTerm
        '
        Me.txtPaymentTerm.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPaymentTerm.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaymentTerm.Location = New System.Drawing.Point(253, 225)
        Me.txtPaymentTerm.MaxLength = 50
        Me.txtPaymentTerm.Name = "txtPaymentTerm"
        Me.txtPaymentTerm.ReadOnly = True
        Me.txtPaymentTerm.Size = New System.Drawing.Size(85, 20)
        Me.txtPaymentTerm.TabIndex = 15
        Me.txtPaymentTerm.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dtpDate
        '
        Me.dtpDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpDate.CustomFormat = "dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(248, 4)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(89, 20)
        Me.dtpDate.TabIndex = 1
        '
        'txtPhoneNo
        '
        Me.txtPhoneNo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPhoneNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPhoneNo.Location = New System.Drawing.Point(107, 101)
        Me.txtPhoneNo.MaxLength = 50
        Me.txtPhoneNo.Name = "txtPhoneNo"
        Me.txtPhoneNo.ReadOnly = True
        Me.txtPhoneNo.Size = New System.Drawing.Size(232, 20)
        Me.txtPhoneNo.TabIndex = 7
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 13)
        Me.Label5.TabIndex = 136
        Me.Label5.Text = "Address Code"
        '
        'cboCustomerAddress
        '
        Me.cboCustomerAddress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCustomerAddress.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerAddress.Enabled = False
        Me.cboCustomerAddress.FormattingEnabled = True
        Me.cboCustomerAddress.Location = New System.Drawing.Point(107, 51)
        Me.cboCustomerAddress.Name = "cboCustomerAddress"
        Me.cboCustomerAddress.Size = New System.Drawing.Size(167, 21)
        Me.cboCustomerAddress.TabIndex = 4
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(6, 179)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(92, 13)
        Me.Label17.TabIndex = 146
        Me.Label17.Text = "Customer's Inquiry"
        '
        'LlbViewAddressCustomerShipTo
        '
        Me.LlbViewAddressCustomerShipTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbViewAddressCustomerShipTo.AutoSize = True
        Me.LlbViewAddressCustomerShipTo.Location = New System.Drawing.Point(263, 57)
        Me.LlbViewAddressCustomerShipTo.Name = "LlbViewAddressCustomerShipTo"
        Me.LlbViewAddressCustomerShipTo.Size = New System.Drawing.Size(29, 13)
        Me.LlbViewAddressCustomerShipTo.TabIndex = 23
        Me.LlbViewAddressCustomerShipTo.TabStop = True
        Me.LlbViewAddressCustomerShipTo.Text = "view"
        '
        'TxtIdentifierAddressCustomerShipTo
        '
        Me.TxtIdentifierAddressCustomerShipTo.Location = New System.Drawing.Point(40, 57)
        Me.TxtIdentifierAddressCustomerShipTo.Name = "TxtIdentifierAddressCustomerShipTo"
        Me.TxtIdentifierAddressCustomerShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtIdentifierAddressCustomerShipTo.TabIndex = 223
        Me.TxtIdentifierAddressCustomerShipTo.TabStop = False
        Me.TxtIdentifierAddressCustomerShipTo.Text = "-1"
        Me.TxtIdentifierAddressCustomerShipTo.Visible = False
        '
        'CboAddressCustomerShipTo
        '
        Me.CboAddressCustomerShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAddressCustomerShipTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAddressCustomerShipTo.Enabled = False
        Me.CboAddressCustomerShipTo.FormattingEnabled = True
        Me.CboAddressCustomerShipTo.Location = New System.Drawing.Point(117, 54)
        Me.CboAddressCustomerShipTo.Name = "CboAddressCustomerShipTo"
        Me.CboAddressCustomerShipTo.Size = New System.Drawing.Size(137, 21)
        Me.CboAddressCustomerShipTo.TabIndex = 22
        '
        'TxtCustomerDescShipTo
        '
        Me.TxtCustomerDescShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerDescShipTo.Location = New System.Drawing.Point(196, 32)
        Me.TxtCustomerDescShipTo.MaxLength = 150
        Me.TxtCustomerDescShipTo.Name = "TxtCustomerDescShipTo"
        Me.TxtCustomerDescShipTo.ReadOnly = True
        Me.TxtCustomerDescShipTo.Size = New System.Drawing.Size(93, 20)
        Me.TxtCustomerDescShipTo.TabIndex = 21
        Me.TxtCustomerDescShipTo.TabStop = False
        '
        'TxtPrimaryCustomerShipTo
        '
        Me.TxtPrimaryCustomerShipTo.Location = New System.Drawing.Point(117, 32)
        Me.TxtPrimaryCustomerShipTo.Name = "TxtPrimaryCustomerShipTo"
        Me.TxtPrimaryCustomerShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCustomerShipTo.TabIndex = 218
        Me.TxtPrimaryCustomerShipTo.TabStop = False
        Me.TxtPrimaryCustomerShipTo.Text = "-1"
        Me.TxtPrimaryCustomerShipTo.Visible = False
        '
        'RboCustomerShipTo
        '
        Me.RboCustomerShipTo.AutoSize = True
        Me.RboCustomerShipTo.Enabled = False
        Me.RboCustomerShipTo.Location = New System.Drawing.Point(3, 33)
        Me.RboCustomerShipTo.Name = "RboCustomerShipTo"
        Me.RboCustomerShipTo.Size = New System.Drawing.Size(105, 17)
        Me.RboCustomerShipTo.TabIndex = 19
        Me.RboCustomerShipTo.Text = "Ship to Customer"
        Me.RboCustomerShipTo.UseVisualStyleBackColor = True
        '
        'TxtCustomerCodeShipTo
        '
        Me.TxtCustomerCodeShipTo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCodeShipTo.Location = New System.Drawing.Point(117, 32)
        Me.TxtCustomerCodeShipTo.MaxLength = 50
        Me.TxtCustomerCodeShipTo.Name = "TxtCustomerCodeShipTo"
        Me.TxtCustomerCodeShipTo.ReadOnly = True
        Me.TxtCustomerCodeShipTo.Size = New System.Drawing.Size(73, 20)
        Me.TxtCustomerCodeShipTo.TabIndex = 20
        '
        'LlbViewAddressShipTo
        '
        Me.LlbViewAddressShipTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbViewAddressShipTo.AutoSize = True
        Me.LlbViewAddressShipTo.Location = New System.Drawing.Point(263, 10)
        Me.LlbViewAddressShipTo.Name = "LlbViewAddressShipTo"
        Me.LlbViewAddressShipTo.Size = New System.Drawing.Size(29, 13)
        Me.LlbViewAddressShipTo.TabIndex = 18
        Me.LlbViewAddressShipTo.TabStop = True
        Me.LlbViewAddressShipTo.Text = "view"
        '
        'CboAddressShipTo
        '
        Me.CboAddressShipTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboAddressShipTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboAddressShipTo.Enabled = False
        Me.CboAddressShipTo.FormattingEnabled = True
        Me.CboAddressShipTo.Location = New System.Drawing.Point(117, 6)
        Me.CboAddressShipTo.Name = "CboAddressShipTo"
        Me.CboAddressShipTo.Size = New System.Drawing.Size(137, 21)
        Me.CboAddressShipTo.TabIndex = 17
        '
        'TxtIdentifierAddressShipTo
        '
        Me.TxtIdentifierAddressShipTo.Location = New System.Drawing.Point(96, 6)
        Me.TxtIdentifierAddressShipTo.Name = "TxtIdentifierAddressShipTo"
        Me.TxtIdentifierAddressShipTo.Size = New System.Drawing.Size(25, 20)
        Me.TxtIdentifierAddressShipTo.TabIndex = 213
        Me.TxtIdentifierAddressShipTo.TabStop = False
        Me.TxtIdentifierAddressShipTo.Text = "-1"
        Me.TxtIdentifierAddressShipTo.Visible = False
        '
        'RboAddressShipTo
        '
        Me.RboAddressShipTo.AutoSize = True
        Me.RboAddressShipTo.Checked = True
        Me.RboAddressShipTo.Enabled = False
        Me.RboAddressShipTo.Location = New System.Drawing.Point(3, 6)
        Me.RboAddressShipTo.Name = "RboAddressShipTo"
        Me.RboAddressShipTo.Size = New System.Drawing.Size(99, 17)
        Me.RboAddressShipTo.TabIndex = 16
        Me.RboAddressShipTo.TabStop = True
        Me.RboAddressShipTo.Text = "Ship to Address"
        Me.RboAddressShipTo.UseVisualStyleBackColor = True
        '
        'TxtDPFromPercent
        '
        Me.TxtDPFromPercent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDPFromPercent.Cursor = System.Windows.Forms.Cursors.Default
        Me.TxtDPFromPercent.Location = New System.Drawing.Point(56, 176)
        Me.TxtDPFromPercent.Name = "TxtDPFromPercent"
        Me.TxtDPFromPercent.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.TxtDPFromPercent.Size = New System.Drawing.Size(10, 20)
        Me.TxtDPFromPercent.TabIndex = 199
        Me.TxtDPFromPercent.TabStop = False
        Me.TxtDPFromPercent.Text = "1"
        Me.TxtDPFromPercent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtDPFromPercent.Visible = False
        '
        'TxtDPAmount
        '
        Me.TxtDPAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDPAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDPAmount.Location = New System.Drawing.Point(181, 176)
        Me.TxtDPAmount.MaxLength = 20
        Me.TxtDPAmount.Name = "TxtDPAmount"
        Me.TxtDPAmount.ReadOnly = True
        Me.TxtDPAmount.Size = New System.Drawing.Size(107, 20)
        Me.TxtDPAmount.TabIndex = 32
        Me.TxtDPAmount.Text = "0"
        Me.TxtDPAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(160, 179)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(15, 13)
        Me.Label13.TabIndex = 202
        Me.Label13.Text = "%"
        '
        'chkUseDP
        '
        Me.chkUseDP.AutoSize = True
        Me.chkUseDP.Enabled = False
        Me.chkUseDP.Location = New System.Drawing.Point(5, 178)
        Me.chkUseDP.Name = "chkUseDP"
        Me.chkUseDP.Size = New System.Drawing.Size(63, 17)
        Me.chkUseDP.TabIndex = 30
        Me.chkUseDP.Text = "Use DP"
        Me.chkUseDP.UseVisualStyleBackColor = True
        '
        'TxtDPPercent
        '
        Me.TxtDPPercent.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtDPPercent.Location = New System.Drawing.Point(117, 176)
        Me.TxtDPPercent.MaxLength = 9
        Me.TxtDPPercent.Name = "TxtDPPercent"
        Me.TxtDPPercent.ReadOnly = True
        Me.TxtDPPercent.Size = New System.Drawing.Size(36, 20)
        Me.TxtDPPercent.TabIndex = 31
        Me.TxtDPPercent.Text = "0"
        Me.TxtDPPercent.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtTermsNCondition
        '
        Me.txtTermsNCondition.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTermsNCondition.Location = New System.Drawing.Point(117, 199)
        Me.txtTermsNCondition.MaxLength = 50
        Me.txtTermsNCondition.Multiline = True
        Me.txtTermsNCondition.Name = "txtTermsNCondition"
        Me.txtTermsNCondition.ReadOnly = True
        Me.txtTermsNCondition.Size = New System.Drawing.Size(173, 50)
        Me.txtTermsNCondition.TabIndex = 33
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(2, 199)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(87, 13)
        Me.Label12.TabIndex = 154
        Me.Label12.Text = "Term && Condition"
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(431, 275)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 166
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(9, 10)
        Me.SplitContainer2.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer3)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.gboDetail)
        Me.SplitContainer2.Size = New System.Drawing.Size(1001, 449)
        Me.SplitContainer2.SplitterDistance = 348
        Me.SplitContainer2.SplitterWidth = 3
        Me.SplitContainer2.TabIndex = 169
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.TabControl2)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.GroupBox1)
        Me.SplitContainer3.Size = New System.Drawing.Size(348, 449)
        Me.SplitContainer3.SplitterDistance = 223
        Me.SplitContainer3.SplitterWidth = 3
        Me.SplitContainer3.TabIndex = 169
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage2)
        Me.TabControl2.Controls.Add(Me.TabPage3)
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.Location = New System.Drawing.Point(0, 0)
        Me.TabControl2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(348, 223)
        Me.TabControl2.TabIndex = 169
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.DgvOtorisasi)
        Me.TabPage2.Controls.Add(Me.LlbUncheckAll_Athorize)
        Me.TabPage2.Controls.Add(Me.LlbCheckAll_Authorize)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(340, 197)
        Me.TabPage2.TabIndex = 0
        Me.TabPage2.Text = "Authorize"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'DgvOtorisasi
        '
        Me.DgvOtorisasi.AllowUserToAddRows = False
        Me.DgvOtorisasi.AllowUserToDeleteRows = False
        Me.DgvOtorisasi.AllowUserToResizeColumns = False
        Me.DgvOtorisasi.AllowUserToResizeRows = False
        Me.DgvOtorisasi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvOtorisasi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvOtorisasi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvOtorisasi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDgvAuth_Primary, Me.ColDgvAuth_Authorize, Me.ColDgvAuth_QuoNo, Me.ColDgvAuth_CustomerName})
        Me.DgvOtorisasi.Location = New System.Drawing.Point(2, 2)
        Me.DgvOtorisasi.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvOtorisasi.MultiSelect = False
        Me.DgvOtorisasi.Name = "DgvOtorisasi"
        Me.DgvOtorisasi.RowHeadersVisible = False
        Me.DgvOtorisasi.RowTemplate.Height = 24
        Me.DgvOtorisasi.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvOtorisasi.Size = New System.Drawing.Size(336, 178)
        Me.DgvOtorisasi.TabIndex = 170
        '
        'ColDgvAuth_Primary
        '
        Me.ColDgvAuth_Primary.HeaderText = "Primary"
        Me.ColDgvAuth_Primary.Name = "ColDgvAuth_Primary"
        Me.ColDgvAuth_Primary.ReadOnly = True
        Me.ColDgvAuth_Primary.Visible = False
        Me.ColDgvAuth_Primary.Width = 66
        '
        'ColDgvAuth_Authorize
        '
        Me.ColDgvAuth_Authorize.HeaderText = "Authorize"
        Me.ColDgvAuth_Authorize.Name = "ColDgvAuth_Authorize"
        Me.ColDgvAuth_Authorize.Width = 57
        '
        'ColDgvAuth_QuoNo
        '
        Me.ColDgvAuth_QuoNo.HeaderText = "Quotation No"
        Me.ColDgvAuth_QuoNo.MinimumWidth = 64
        Me.ColDgvAuth_QuoNo.Name = "ColDgvAuth_QuoNo"
        Me.ColDgvAuth_QuoNo.ReadOnly = True
        Me.ColDgvAuth_QuoNo.Width = 95
        '
        'ColDgvAuth_CustomerName
        '
        Me.ColDgvAuth_CustomerName.HeaderText = "Customer"
        Me.ColDgvAuth_CustomerName.Name = "ColDgvAuth_CustomerName"
        Me.ColDgvAuth_CustomerName.ReadOnly = True
        Me.ColDgvAuth_CustomerName.Width = 76
        '
        'LlbUncheckAll_Athorize
        '
        Me.LlbUncheckAll_Athorize.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll_Athorize.AutoSize = True
        Me.LlbUncheckAll_Athorize.Location = New System.Drawing.Point(267, 183)
        Me.LlbUncheckAll_Athorize.Name = "LlbUncheckAll_Athorize"
        Me.LlbUncheckAll_Athorize.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll_Athorize.TabIndex = 172
        Me.LlbUncheckAll_Athorize.TabStop = True
        Me.LlbUncheckAll_Athorize.Text = "Uncheck All"
        '
        'LlbCheckAll_Authorize
        '
        Me.LlbCheckAll_Authorize.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll_Authorize.AutoSize = True
        Me.LlbCheckAll_Authorize.Location = New System.Drawing.Point(211, 183)
        Me.LlbCheckAll_Authorize.Name = "LlbCheckAll_Authorize"
        Me.LlbCheckAll_Authorize.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll_Authorize.TabIndex = 171
        Me.LlbCheckAll_Authorize.TabStop = True
        Me.LlbCheckAll_Authorize.Text = "Check All"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvCancelOtorisasi)
        Me.TabPage3.Controls.Add(Me.LlbUncheckAll_Cancel)
        Me.TabPage3.Controls.Add(Me.LlbCheckAll_Cancel)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(340, 197)
        Me.TabPage3.TabIndex = 1
        Me.TabPage3.Text = "Cancel Authorize"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'DgvCancelOtorisasi
        '
        Me.DgvCancelOtorisasi.AllowUserToAddRows = False
        Me.DgvCancelOtorisasi.AllowUserToDeleteRows = False
        Me.DgvCancelOtorisasi.AllowUserToResizeColumns = False
        Me.DgvCancelOtorisasi.AllowUserToResizeRows = False
        Me.DgvCancelOtorisasi.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvCancelOtorisasi.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvCancelOtorisasi.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvCancelOtorisasi.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDgvUnAuth_Primary, Me.ColDgvUnAuth_Cancel, Me.ColDgvUnAuth_QuoNo, Me.ColDgvUnAuth_CustomerName})
        Me.DgvCancelOtorisasi.Location = New System.Drawing.Point(2, 2)
        Me.DgvCancelOtorisasi.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvCancelOtorisasi.MultiSelect = False
        Me.DgvCancelOtorisasi.Name = "DgvCancelOtorisasi"
        Me.DgvCancelOtorisasi.RowHeadersVisible = False
        Me.DgvCancelOtorisasi.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvCancelOtorisasi.RowTemplate.Height = 24
        Me.DgvCancelOtorisasi.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvCancelOtorisasi.Size = New System.Drawing.Size(336, 178)
        Me.DgvCancelOtorisasi.TabIndex = 173
        '
        'ColDgvUnAuth_Primary
        '
        Me.ColDgvUnAuth_Primary.HeaderText = "Primary"
        Me.ColDgvUnAuth_Primary.Name = "ColDgvUnAuth_Primary"
        Me.ColDgvUnAuth_Primary.ReadOnly = True
        Me.ColDgvUnAuth_Primary.Visible = False
        Me.ColDgvUnAuth_Primary.Width = 66
        '
        'ColDgvUnAuth_Cancel
        '
        Me.ColDgvUnAuth_Cancel.HeaderText = "Cancel Authorize"
        Me.ColDgvUnAuth_Cancel.Name = "ColDgvUnAuth_Cancel"
        Me.ColDgvUnAuth_Cancel.Width = 93
        '
        'ColDgvUnAuth_QuoNo
        '
        Me.ColDgvUnAuth_QuoNo.HeaderText = "Quotation No"
        Me.ColDgvUnAuth_QuoNo.Name = "ColDgvUnAuth_QuoNo"
        Me.ColDgvUnAuth_QuoNo.ReadOnly = True
        Me.ColDgvUnAuth_QuoNo.Width = 95
        '
        'ColDgvUnAuth_CustomerName
        '
        Me.ColDgvUnAuth_CustomerName.HeaderText = "Customer"
        Me.ColDgvUnAuth_CustomerName.Name = "ColDgvUnAuth_CustomerName"
        Me.ColDgvUnAuth_CustomerName.ReadOnly = True
        Me.ColDgvUnAuth_CustomerName.Width = 76
        '
        'LlbUncheckAll_Cancel
        '
        Me.LlbUncheckAll_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll_Cancel.AutoSize = True
        Me.LlbUncheckAll_Cancel.Location = New System.Drawing.Point(267, 183)
        Me.LlbUncheckAll_Cancel.Name = "LlbUncheckAll_Cancel"
        Me.LlbUncheckAll_Cancel.Size = New System.Drawing.Size(65, 13)
        Me.LlbUncheckAll_Cancel.TabIndex = 172
        Me.LlbUncheckAll_Cancel.TabStop = True
        Me.LlbUncheckAll_Cancel.Text = "Uncheck All"
        '
        'LlbCheckAll_Cancel
        '
        Me.LlbCheckAll_Cancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll_Cancel.AutoSize = True
        Me.LlbCheckAll_Cancel.Location = New System.Drawing.Point(211, 183)
        Me.LlbCheckAll_Cancel.Name = "LlbCheckAll_Cancel"
        Me.LlbCheckAll_Cancel.Size = New System.Drawing.Size(52, 13)
        Me.LlbCheckAll_Cancel.TabIndex = 171
        Me.LlbCheckAll_Cancel.TabStop = True
        Me.LlbCheckAll_Cancel.Text = "Check All"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DgvAlreadyAuthorize)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(340, 197)
        Me.TabPage4.TabIndex = 2
        Me.TabPage4.Text = "Already Authorize"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DgvAlreadyAuthorize
        '
        Me.DgvAlreadyAuthorize.AllowUserToAddRows = False
        Me.DgvAlreadyAuthorize.AllowUserToDeleteRows = False
        Me.DgvAlreadyAuthorize.AllowUserToResizeColumns = False
        Me.DgvAlreadyAuthorize.AllowUserToResizeRows = False
        Me.DgvAlreadyAuthorize.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvAlreadyAuthorize.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvAlreadyAuthorize.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAA_Primary, Me.ColAA_Print, Me.ColAA_QuoNo, Me.ColAA_CustomerName, Me.ColAA_Status})
        Me.DgvAlreadyAuthorize.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvAlreadyAuthorize.Location = New System.Drawing.Point(2, 2)
        Me.DgvAlreadyAuthorize.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvAlreadyAuthorize.MultiSelect = False
        Me.DgvAlreadyAuthorize.Name = "DgvAlreadyAuthorize"
        Me.DgvAlreadyAuthorize.RowHeadersVisible = False
        Me.DgvAlreadyAuthorize.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvAlreadyAuthorize.RowTemplate.Height = 24
        Me.DgvAlreadyAuthorize.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvAlreadyAuthorize.Size = New System.Drawing.Size(336, 193)
        Me.DgvAlreadyAuthorize.TabIndex = 172
        '
        'ColAA_Primary
        '
        Me.ColAA_Primary.HeaderText = "Primary"
        Me.ColAA_Primary.Name = "ColAA_Primary"
        Me.ColAA_Primary.ReadOnly = True
        Me.ColAA_Primary.Visible = False
        Me.ColAA_Primary.Width = 66
        '
        'ColAA_Print
        '
        Me.ColAA_Print.HeaderText = "Print"
        Me.ColAA_Print.Name = "ColAA_Print"
        Me.ColAA_Print.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAA_Print.Width = 34
        '
        'ColAA_QuoNo
        '
        Me.ColAA_QuoNo.HeaderText = "Quotation No"
        Me.ColAA_QuoNo.Name = "ColAA_QuoNo"
        Me.ColAA_QuoNo.ReadOnly = True
        Me.ColAA_QuoNo.Width = 95
        '
        'ColAA_CustomerName
        '
        Me.ColAA_CustomerName.HeaderText = "Customer"
        Me.ColAA_CustomerName.Name = "ColAA_CustomerName"
        Me.ColAA_CustomerName.ReadOnly = True
        Me.ColAA_CustomerName.Width = 76
        '
        'ColAA_Status
        '
        Me.ColAA_Status.HeaderText = "Status"
        Me.ColAA_Status.Name = "ColAA_Status"
        Me.ColAA_Status.ReadOnly = True
        Me.ColAA_Status.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColAA_Status.Width = 62
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnFilter)
        Me.GroupBox1.Controls.Add(Me.TxtNoQuotation)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.BtnFindQuo)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.DgvUser)
        Me.GroupBox1.Controls.Add(Me.CboStatusQuotation)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryQuo)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(348, 223)
        Me.GroupBox1.TabIndex = 169
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(295, 12)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(45, 23)
        Me.BtnFilter.TabIndex = 171
        Me.BtnFilter.TabStop = False
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'TxtNoQuotation
        '
        Me.TxtNoQuotation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNoQuotation.Location = New System.Drawing.Point(84, 38)
        Me.TxtNoQuotation.Name = "TxtNoQuotation"
        Me.TxtNoQuotation.Size = New System.Drawing.Size(207, 20)
        Me.TxtNoQuotation.TabIndex = 174
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(8, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(37, 13)
        Me.Label9.TabIndex = 172
        Me.Label9.Text = "Status"
        '
        'BtnFindQuo
        '
        Me.BtnFindQuo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindQuo.Location = New System.Drawing.Point(295, 37)
        Me.BtnFindQuo.Name = "BtnFindQuo"
        Me.BtnFindQuo.Size = New System.Drawing.Size(45, 22)
        Me.BtnFindQuo.TabIndex = 171
        Me.BtnFindQuo.TabStop = False
        Me.BtnFindQuo.Text = "...."
        Me.BtnFindQuo.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.DtpTo)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.DtpFrom)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.ChkQuoDate)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 59)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox2.Size = New System.Drawing.Size(340, 64)
        Me.GroupBox2.TabIndex = 175
        Me.GroupBox2.TabStop = False
        '
        'DtpTo
        '
        Me.DtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpTo.CustomFormat = "dd MMM yyyy"
        Me.DtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpTo.Location = New System.Drawing.Point(80, 40)
        Me.DtpTo.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpTo.Name = "DtpTo"
        Me.DtpTo.Size = New System.Drawing.Size(255, 20)
        Me.DtpTo.TabIndex = 177
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 43)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 170
        Me.Label3.Text = "To"
        '
        'DtpFrom
        '
        Me.DtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DtpFrom.CustomFormat = "dd MMM yyyy"
        Me.DtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpFrom.Location = New System.Drawing.Point(80, 17)
        Me.DtpFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpFrom.Name = "DtpFrom"
        Me.DtpFrom.Size = New System.Drawing.Size(255, 20)
        Me.DtpFrom.TabIndex = 176
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 20)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 169
        Me.Label2.Text = "From"
        '
        'ChkQuoDate
        '
        Me.ChkQuoDate.AutoSize = True
        Me.ChkQuoDate.Location = New System.Drawing.Point(5, 0)
        Me.ChkQuoDate.Name = "ChkQuoDate"
        Me.ChkQuoDate.Size = New System.Drawing.Size(98, 17)
        Me.ChkQuoDate.TabIndex = 175
        Me.ChkQuoDate.Text = "Quotation Date"
        Me.ChkQuoDate.UseVisualStyleBackColor = True
        '
        'DgvUser
        '
        Me.DgvUser.AllowUserToAddRows = False
        Me.DgvUser.AllowUserToResizeColumns = False
        Me.DgvUser.AllowUserToResizeRows = False
        Me.DgvUser.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCreatedBy_Primary, Me.ColCreatedBy, Me.ColCreatedBy_Name})
        Me.DgvUser.Location = New System.Drawing.Point(4, 127)
        Me.DgvUser.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvUser.Name = "DgvUser"
        Me.DgvUser.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvUser.RowTemplate.Height = 24
        Me.DgvUser.Size = New System.Drawing.Size(340, 91)
        Me.DgvUser.TabIndex = 178
        '
        'ColCreatedBy_Primary
        '
        Me.ColCreatedBy_Primary.HeaderText = "Primary"
        Me.ColCreatedBy_Primary.Name = "ColCreatedBy_Primary"
        Me.ColCreatedBy_Primary.ReadOnly = True
        Me.ColCreatedBy_Primary.Visible = False
        Me.ColCreatedBy_Primary.Width = 66
        '
        'ColCreatedBy
        '
        Me.ColCreatedBy.HeaderText = "User Code"
        Me.ColCreatedBy.Name = "ColCreatedBy"
        Me.ColCreatedBy.Width = 82
        '
        'ColCreatedBy_Name
        '
        Me.ColCreatedBy_Name.HeaderText = "Name"
        Me.ColCreatedBy_Name.Name = "ColCreatedBy_Name"
        Me.ColCreatedBy_Name.ReadOnly = True
        Me.ColCreatedBy_Name.Width = 60
        '
        'CboStatusQuotation
        '
        Me.CboStatusQuotation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboStatusQuotation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboStatusQuotation.FormattingEnabled = True
        Me.CboStatusQuotation.Items.AddRange(New Object() {"All Quotation", "Auth. By Sales Staff", "Auth. By Sales Mgr", "Auth. By Gen. Mgr", "Unauthorized Only"})
        Me.CboStatusQuotation.Location = New System.Drawing.Point(84, 13)
        Me.CboStatusQuotation.Name = "CboStatusQuotation"
        Me.CboStatusQuotation.Size = New System.Drawing.Size(207, 21)
        Me.CboStatusQuotation.TabIndex = 173
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(8, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(70, 13)
        Me.Label10.TabIndex = 174
        Me.Label10.Text = "Quotation No"
        '
        'TxtPrimaryQuo
        '
        Me.TxtPrimaryQuo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPrimaryQuo.Location = New System.Drawing.Point(26, 37)
        Me.TxtPrimaryQuo.MaxLength = 50
        Me.TxtPrimaryQuo.Name = "TxtPrimaryQuo"
        Me.TxtPrimaryQuo.Size = New System.Drawing.Size(28, 20)
        Me.TxtPrimaryQuo.TabIndex = 167
        Me.TxtPrimaryQuo.TabStop = False
        Me.TxtPrimaryQuo.Visible = False
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 465)
        Me.SplitContainer4.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.ToolStrip2)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.ToolStrip1)
        Me.SplitContainer4.Size = New System.Drawing.Size(1010, 33)
        Me.SplitContainer4.SplitterDistance = 566
        Me.SplitContainer4.SplitterWidth = 3
        Me.SplitContainer4.TabIndex = 170
        '
        'ToolStrip2
        '
        Me.ToolStrip2.AutoSize = False
        Me.ToolStrip2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip2.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip2.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbRemove, Me.tsbShowRemoved})
        Me.ToolStrip2.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip2.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip2.Name = "ToolStrip2"
        Me.ToolStrip2.Size = New System.Drawing.Size(566, 33)
        Me.ToolStrip2.TabIndex = 60
        '
        'tsbRemove
        '
        Me.tsbRemove.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbRemove.Name = "tsbRemove"
        Me.tsbRemove.Size = New System.Drawing.Size(54, 30)
        Me.tsbRemove.Text = "Remove"
        Me.tsbRemove.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.tsbRemove.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.tsbRemove.Visible = False
        '
        'tsbShowRemoved
        '
        Me.tsbShowRemoved.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbShowRemoved.Name = "tsbShowRemoved"
        Me.tsbShowRemoved.Size = New System.Drawing.Size(93, 30)
        Me.tsbShowRemoved.Text = "Show Removed"
        Me.tsbShowRemoved.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        Me.tsbShowRemoved.Visible = False
        '
        'BgwUser
        '
        '
        'FrmTransaksiQuotationOtorisasi2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(1010, 520)
        Me.Controls.Add(Me.SplitContainer4)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.SplitContainer2)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(1026, 558)
        Me.Name = "FrmTransaksiQuotationOtorisasi2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Authorize Quotation"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.DgvOtorisasi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        CType(Me.DgvCancelOtorisasi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        CType(Me.DgvAlreadyAuthorize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.ResumeLayout(False)
        Me.ToolStrip2.ResumeLayout(False)
        Me.ToolStrip2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbAuthorize As System.Windows.Forms.ToolStripButton
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtTaxAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents chkTax As System.Windows.Forms.CheckBox
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtTax As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtContactPerson As System.Windows.Forms.TextBox
    Friend WithEvents btnCLose As System.Windows.Forms.Button
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents llbAuthorizationDetail As System.Windows.Forms.LinkLabel
    Friend WithEvents txtAutohirizationStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cboPaymentDueDateReference As System.Windows.Forms.ComboBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtStockInNo As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtPaymentTerm As System.Windows.Forms.TextBox
    Friend WithEvents txtPaymentDueDate As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents DtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents LlbUncheckAll_Athorize As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll_Authorize As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvCancelOtorisasi As System.Windows.Forms.DataGridView
    Friend WithEvents LlbUncheckAll_Cancel As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll_Cancel As System.Windows.Forms.LinkLabel
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents DgvAlreadyAuthorize As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents ToolStrip2 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbRemove As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbShowRemoved As System.Windows.Forms.ToolStripButton
    Friend WithEvents CboStatusQuotation As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtNoQuotation As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BtnFindQuo As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryQuo As System.Windows.Forms.TextBox
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents BgwUser As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkQuoDate As System.Windows.Forms.CheckBox
    Friend WithEvents ColCreatedBy_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCreatedBy_Name As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtInquiryNo As System.Windows.Forms.TextBox
    Friend WithEvents LlbAddress As System.Windows.Forms.LinkLabel
    Friend WithEvents txtCustomerAddIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtPhoneNo As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboCustomerAddress As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents txtNoNPWP As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtTermsNCondition As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents DgvOtorisasi As System.Windows.Forms.DataGridView
    Friend WithEvents ColDgvAuth_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvAuth_Authorize As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDgvAuth_QuoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvAuth_CustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_Cancel As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDgvUnAuth_QuoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDgvUnAuth_CustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents TsbPrint As System.Windows.Forms.ToolStripButton
    Friend WithEvents ColAA_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAA_Print As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAA_QuoNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAA_CustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAA_Status As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupplierCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriceByMultiUnit_Str As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryCont As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDown As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDownPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUpPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUseSecondary As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitsMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDownFromPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUpFromPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriceByMultiUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TxtDPFromPercent As System.Windows.Forms.TextBox
    Friend WithEvents TxtDPAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents chkUseDP As System.Windows.Forms.CheckBox
    Friend WithEvents TxtDPPercent As System.Windows.Forms.TextBox
    Friend WithEvents RboAddressShipTo As System.Windows.Forms.RadioButton
    Friend WithEvents TxtIdentifierAddressShipTo As System.Windows.Forms.TextBox
    Friend WithEvents CboAddressShipTo As System.Windows.Forms.ComboBox
    Friend WithEvents LlbViewAddressShipTo As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtPrimaryCustomerShipTo As System.Windows.Forms.TextBox
    Friend WithEvents RboCustomerShipTo As System.Windows.Forms.RadioButton
    Friend WithEvents TxtCustomerCodeShipTo As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerDescShipTo As System.Windows.Forms.TextBox
    Friend WithEvents LlbViewAddressCustomerShipTo As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtIdentifierAddressCustomerShipTo As System.Windows.Forms.TextBox
    Friend WithEvents CboAddressCustomerShipTo As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
End Class
