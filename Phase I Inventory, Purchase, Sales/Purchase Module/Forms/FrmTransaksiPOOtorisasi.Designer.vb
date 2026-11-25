<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPOOtorisasi
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
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.btnExit = New System.Windows.Forms.Button
        Me.chkBelumOtorisasi = New System.Windows.Forms.CheckedListBox
        Me.TabBlmOtorisasi = New System.Windows.Forms.TabPage
        Me.llbUncheckAllAuthorize = New System.Windows.Forms.LinkLabel
        Me.llbCheckAllAuthorize = New System.Windows.Forms.LinkLabel
        Me.Button1 = New System.Windows.Forms.Button
        Me.TabCancel = New System.Windows.Forms.TabPage
        Me.llbUncheckAllCancel = New System.Windows.Forms.LinkLabel
        Me.llbCheckAllCancel = New System.Windows.Forms.LinkLabel
        Me.chkSudahOtorisasi = New System.Windows.Forms.CheckedListBox
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDesc = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatioReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIPONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryCont = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColRemark = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryIPO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUserSecondaryUnits = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscFromPerct = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDown = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceDownPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColToleranceUpPerc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.txtSupplierAddIdentifier = New System.Windows.Forms.TextBox
        Me.txtTaxAmount = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.chkTax = New System.Windows.Forms.CheckBox
        Me.txtTax = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtSubTotal = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TxtShipVia = New System.Windows.Forms.TextBox
        Me.TxtRemark = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtPhoneNo = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtContactPerson = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboSupplierAddress = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimarySupplier = New System.Windows.Forms.TextBox
        Me.txtSupplierCode = New System.Windows.Forms.TextBox
        Me.TxtSupplierName = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.llbAuthorizationDetail = New System.Windows.Forms.LinkLabel
        Me.txtAutohirizationStatus = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.cboPaymentDueDateReference = New System.Windows.Forms.ComboBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.txtPaymentDueDate = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtPaymentTerm = New System.Windows.Forms.TextBox
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryStockIn = New System.Windows.Forms.TextBox
        Me.TxtStockInNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.StatusStrip2 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.tsbExit = New System.Windows.Forms.ToolStripButton
        Me.tsbAuthorize = New System.Windows.Forms.ToolStripButton
        Me.tsbRemove = New System.Windows.Forms.ToolStripButton
        Me.tsbShowRemoved = New System.Windows.Forms.ToolStripButton
        Me.TabBlmOtorisasi.SuspendLayout()
        Me.TabCancel.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip2.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(652, 37)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(62, 29)
        Me.btnExit.TabIndex = 43
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'chkBelumOtorisasi
        '
        Me.chkBelumOtorisasi.ColumnWidth = 180
        Me.chkBelumOtorisasi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkBelumOtorisasi.FormattingEnabled = True
        Me.chkBelumOtorisasi.Location = New System.Drawing.Point(3, 3)
        Me.chkBelumOtorisasi.MultiColumn = True
        Me.chkBelumOtorisasi.Name = "chkBelumOtorisasi"
        Me.chkBelumOtorisasi.Size = New System.Drawing.Size(850, 94)
        Me.chkBelumOtorisasi.TabIndex = 0
        '
        'TabBlmOtorisasi
        '
        Me.TabBlmOtorisasi.Controls.Add(Me.llbUncheckAllAuthorize)
        Me.TabBlmOtorisasi.Controls.Add(Me.llbCheckAllAuthorize)
        Me.TabBlmOtorisasi.Controls.Add(Me.chkBelumOtorisasi)
        Me.TabBlmOtorisasi.Controls.Add(Me.btnExit)
        Me.TabBlmOtorisasi.Controls.Add(Me.Button1)
        Me.TabBlmOtorisasi.Location = New System.Drawing.Point(4, 22)
        Me.TabBlmOtorisasi.Name = "TabBlmOtorisasi"
        Me.TabBlmOtorisasi.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabBlmOtorisasi.Size = New System.Drawing.Size(856, 105)
        Me.TabBlmOtorisasi.TabIndex = 0
        Me.TabBlmOtorisasi.Text = "Authorize"
        Me.TabBlmOtorisasi.UseVisualStyleBackColor = True
        '
        'llbUncheckAllAuthorize
        '
        Me.llbUncheckAllAuthorize.AutoSize = True
        Me.llbUncheckAllAuthorize.Location = New System.Drawing.Point(65, 91)
        Me.llbUncheckAllAuthorize.Name = "llbUncheckAllAuthorize"
        Me.llbUncheckAllAuthorize.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAllAuthorize.TabIndex = 154
        Me.llbUncheckAllAuthorize.TabStop = True
        Me.llbUncheckAllAuthorize.Text = "Uncheck All"
        '
        'llbCheckAllAuthorize
        '
        Me.llbCheckAllAuthorize.AutoSize = True
        Me.llbCheckAllAuthorize.Location = New System.Drawing.Point(5, 91)
        Me.llbCheckAllAuthorize.Name = "llbCheckAllAuthorize"
        Me.llbCheckAllAuthorize.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAllAuthorize.TabIndex = 153
        Me.llbCheckAllAuthorize.TabStop = True
        Me.llbCheckAllAuthorize.Text = "Check All"
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(262, 23)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(122, 42)
        Me.Button1.TabIndex = 155
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TabCancel
        '
        Me.TabCancel.Controls.Add(Me.llbUncheckAllCancel)
        Me.TabCancel.Controls.Add(Me.llbCheckAllCancel)
        Me.TabCancel.Controls.Add(Me.chkSudahOtorisasi)
        Me.TabCancel.Location = New System.Drawing.Point(4, 22)
        Me.TabCancel.Name = "TabCancel"
        Me.TabCancel.Padding = New System.Windows.Forms.Padding(3, 3, 3, 3)
        Me.TabCancel.Size = New System.Drawing.Size(862, 105)
        Me.TabCancel.TabIndex = 1
        Me.TabCancel.Text = "Cancel Authorize"
        Me.TabCancel.UseVisualStyleBackColor = True
        '
        'llbUncheckAllCancel
        '
        Me.llbUncheckAllCancel.AutoSize = True
        Me.llbUncheckAllCancel.Location = New System.Drawing.Point(65, 91)
        Me.llbUncheckAllCancel.Name = "llbUncheckAllCancel"
        Me.llbUncheckAllCancel.Size = New System.Drawing.Size(65, 13)
        Me.llbUncheckAllCancel.TabIndex = 57
        Me.llbUncheckAllCancel.TabStop = True
        Me.llbUncheckAllCancel.Text = "Uncheck All"
        '
        'llbCheckAllCancel
        '
        Me.llbCheckAllCancel.AutoSize = True
        Me.llbCheckAllCancel.Location = New System.Drawing.Point(5, 91)
        Me.llbCheckAllCancel.Name = "llbCheckAllCancel"
        Me.llbCheckAllCancel.Size = New System.Drawing.Size(52, 13)
        Me.llbCheckAllCancel.TabIndex = 56
        Me.llbCheckAllCancel.TabStop = True
        Me.llbCheckAllCancel.Text = "Check All"
        '
        'chkSudahOtorisasi
        '
        Me.chkSudahOtorisasi.ColumnWidth = 180
        Me.chkSudahOtorisasi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.chkSudahOtorisasi.FormattingEnabled = True
        Me.chkSudahOtorisasi.Location = New System.Drawing.Point(3, 3)
        Me.chkSudahOtorisasi.MultiColumn = True
        Me.chkSudahOtorisasi.Name = "chkSudahOtorisasi"
        Me.chkSudahOtorisasi.Size = New System.Drawing.Size(856, 94)
        Me.chkSudahOtorisasi.TabIndex = 1
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabBlmOtorisasi)
        Me.TabControl1.Controls.Add(Me.TabCancel)
        Me.TabControl1.Location = New System.Drawing.Point(12, 6)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(864, 131)
        Me.TabControl1.TabIndex = 0
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.dgvItem)
        Me.gboDetail.Controls.Add(Me.LinkLabel1)
        Me.gboDetail.Controls.Add(Me.txtSupplierAddIdentifier)
        Me.gboDetail.Controls.Add(Me.txtTaxAmount)
        Me.gboDetail.Controls.Add(Me.Label25)
        Me.gboDetail.Controls.Add(Me.chkTax)
        Me.gboDetail.Controls.Add(Me.txtTax)
        Me.gboDetail.Controls.Add(Me.Label24)
        Me.gboDetail.Controls.Add(Me.TxtTotal)
        Me.gboDetail.Controls.Add(Me.Label15)
        Me.gboDetail.Controls.Add(Me.txtSubTotal)
        Me.gboDetail.Controls.Add(Me.Label17)
        Me.gboDetail.Controls.Add(Me.TxtShipVia)
        Me.gboDetail.Controls.Add(Me.TxtRemark)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.Label23)
        Me.gboDetail.Controls.Add(Me.txtPhoneNo)
        Me.gboDetail.Controls.Add(Me.Label7)
        Me.gboDetail.Controls.Add(Me.txtContactPerson)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.cboSupplierAddress)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtPrimarySupplier)
        Me.gboDetail.Controls.Add(Me.txtSupplierCode)
        Me.gboDetail.Controls.Add(Me.TxtSupplierName)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCurrency)
        Me.gboDetail.Controls.Add(Me.Label22)
        Me.gboDetail.Controls.Add(Me.txtCurrencyCode)
        Me.gboDetail.Controls.Add(Me.llbAuthorizationDetail)
        Me.gboDetail.Controls.Add(Me.txtAutohirizationStatus)
        Me.gboDetail.Controls.Add(Me.Label21)
        Me.gboDetail.Controls.Add(Me.cboPaymentDueDateReference)
        Me.gboDetail.Controls.Add(Me.Label20)
        Me.gboDetail.Controls.Add(Me.Label18)
        Me.gboDetail.Controls.Add(Me.txtPaymentDueDate)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.txtPaymentTerm)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryStockIn)
        Me.gboDetail.Controls.Add(Me.TxtStockInNo)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.Button2)
        Me.gboDetail.Location = New System.Drawing.Point(12, 143)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(864, 356)
        Me.gboDetail.TabIndex = 1
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
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDescription, Me.ColQtyMulti, Me.ColUnitCodeMulti, Me.ColRatioDesc, Me.ColQty, Me.ColUnit, Me.ColRatioDefault, Me.ColRatioReference, Me.ColIPONo, Me.ColUnitPrice, Me.ColDiscount, Me.ColDiscAmount, Me.ColNettPrice, Me.ColAmount, Me.ColDeliveryCont, Me.ColDeliveryDate, Me.ColRemark, Me.ColPrimaryCodeNo, Me.ColPrimaryUnits, Me.ColPrimaryIPO, Me.colUserSecondaryUnits, Me.ColIdentifier, Me.ColDiscFromPerct, Me.ColPrimaryUnitMulti, Me.ColToleranceDown, Me.ColToleranceDownPerc, Me.ColToleranceUp, Me.ColToleranceUpPerc})
        Me.dgvItem.Location = New System.Drawing.Point(6, 219)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(852, 130)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 167
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
        'ColQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColUnitCodeMulti
        '
        Me.ColUnitCodeMulti.HeaderText = "Units"
        Me.ColUnitCodeMulti.Name = "ColUnitCodeMulti"
        Me.ColUnitCodeMulti.ReadOnly = True
        Me.ColUnitCodeMulti.Width = 56
        '
        'ColRatioDesc
        '
        Me.ColRatioDesc.ActiveLinkColor = System.Drawing.Color.Blue
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.ColRatioDesc.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColRatioDesc.HeaderText = "Ratio"
        Me.ColRatioDesc.Name = "ColRatioDesc"
        Me.ColRatioDesc.ReadOnly = True
        Me.ColRatioDesc.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRatioDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColRatioDesc.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColRatioDesc.Width = 57
        '
        'ColQty
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle4
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
        'ColRatioDefault
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioDefault.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColRatioDefault.HeaderText = "Ratio Default"
        Me.ColRatioDefault.Name = "ColRatioDefault"
        Me.ColRatioDefault.Visible = False
        Me.ColRatioDefault.Width = 94
        '
        'ColRatioReference
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColRatioReference.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColRatioReference.HeaderText = "Ratio Reference"
        Me.ColRatioReference.Name = "ColRatioReference"
        Me.ColRatioReference.Visible = False
        Me.ColRatioReference.Width = 110
        '
        'ColIPONo
        '
        Me.ColIPONo.HeaderText = "IPO. No"
        Me.ColIPONo.Name = "ColIPONo"
        Me.ColIPONo.Visible = False
        Me.ColIPONo.Width = 70
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.Width = 78
        '
        'ColDiscount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.Width = 91
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColDiscAmount.HeaderText = "Disc. Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.Width = 95
        '
        'ColNettPrice
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.Width = 79
        '
        'ColAmount
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle11
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
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 105
        '
        'ColPrimaryUnits
        '
        Me.ColPrimaryUnits.HeaderText = "PrimaryUnits"
        Me.ColPrimaryUnits.Name = "ColPrimaryUnits"
        Me.ColPrimaryUnits.Visible = False
        Me.ColPrimaryUnits.Width = 90
        '
        'ColPrimaryIPO
        '
        Me.ColPrimaryIPO.HeaderText = "Primary IPO"
        Me.ColPrimaryIPO.Name = "ColPrimaryIPO"
        Me.ColPrimaryIPO.Visible = False
        Me.ColPrimaryIPO.Width = 87
        '
        'colUserSecondaryUnits
        '
        Me.colUserSecondaryUnits.HeaderText = "User Secondary Units"
        Me.colUserSecondaryUnits.Name = "colUserSecondaryUnits"
        Me.colUserSecondaryUnits.Visible = False
        Me.colUserSecondaryUnits.Width = 116
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 72
        '
        'ColDiscFromPerct
        '
        Me.ColDiscFromPerct.HeaderText = "Disc From Perct"
        Me.ColDiscFromPerct.Name = "ColDiscFromPerct"
        Me.ColDiscFromPerct.Visible = False
        Me.ColDiscFromPerct.Width = 107
        '
        'ColPrimaryUnitMulti
        '
        Me.ColPrimaryUnitMulti.HeaderText = "Primary Unit Multi"
        Me.ColPrimaryUnitMulti.Name = "ColPrimaryUnitMulti"
        Me.ColPrimaryUnitMulti.ReadOnly = True
        Me.ColPrimaryUnitMulti.Visible = False
        Me.ColPrimaryUnitMulti.Width = 113
        '
        'ColToleranceDown
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceDown.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColToleranceDown.HeaderText = "Tolerance Down"
        Me.ColToleranceDown.Name = "ColToleranceDown"
        Me.ColToleranceDown.Width = 111
        '
        'ColToleranceDownPerc
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceDownPerc.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColToleranceDownPerc.HeaderText = "Tolerance Down (%)"
        Me.ColToleranceDownPerc.Name = "ColToleranceDownPerc"
        Me.ColToleranceDownPerc.Width = 128
        '
        'ColToleranceUp
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceUp.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColToleranceUp.HeaderText = "Tolerance Up"
        Me.ColToleranceUp.Name = "ColToleranceUp"
        Me.ColToleranceUp.Width = 97
        '
        'ColToleranceUpPerc
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColToleranceUpPerc.DefaultCellStyle = DataGridViewCellStyle15
        Me.ColToleranceUpPerc.HeaderText = "Tolerance Up (%)"
        Me.ColToleranceUpPerc.Name = "ColToleranceUpPerc"
        Me.ColToleranceUpPerc.Width = 114
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(373, 22)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(54, 13)
        Me.LinkLabel1.TabIndex = 165
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "detail sign"
        '
        'txtSupplierAddIdentifier
        '
        Me.txtSupplierAddIdentifier.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierAddIdentifier.Location = New System.Drawing.Point(80, 116)
        Me.txtSupplierAddIdentifier.MaxLength = 50
        Me.txtSupplierAddIdentifier.Name = "txtSupplierAddIdentifier"
        Me.txtSupplierAddIdentifier.Size = New System.Drawing.Size(28, 20)
        Me.txtSupplierAddIdentifier.TabIndex = 163
        Me.txtSupplierAddIdentifier.Visible = False
        '
        'txtTaxAmount
        '
        Me.txtTaxAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxAmount.Location = New System.Drawing.Point(655, 94)
        Me.txtTaxAmount.MaxLength = 50
        Me.txtTaxAmount.Name = "txtTaxAmount"
        Me.txtTaxAmount.ReadOnly = True
        Me.txtTaxAmount.Size = New System.Drawing.Size(197, 20)
        Me.txtTaxAmount.TabIndex = 16
        Me.txtTaxAmount.Text = "0"
        Me.txtTaxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(615, 97)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(15, 13)
        Me.Label25.TabIndex = 152
        Me.Label25.Text = "%"
        '
        'chkTax
        '
        Me.chkTax.AutoSize = True
        Me.chkTax.Enabled = False
        Me.chkTax.Location = New System.Drawing.Point(455, 96)
        Me.chkTax.Name = "chkTax"
        Me.chkTax.Size = New System.Drawing.Size(44, 17)
        Me.chkTax.TabIndex = 14
        Me.chkTax.Text = "Tax"
        Me.chkTax.UseVisualStyleBackColor = True
        '
        'txtTax
        '
        Me.txtTax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTax.Location = New System.Drawing.Point(567, 94)
        Me.txtTax.MaxLength = 50
        Me.txtTax.Name = "txtTax"
        Me.txtTax.ReadOnly = True
        Me.txtTax.Size = New System.Drawing.Size(46, 20)
        Me.txtTax.TabIndex = 15
        Me.txtTax.Text = "0"
        Me.txtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(452, 120)
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
        Me.TxtTotal.Location = New System.Drawing.Point(567, 117)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(285, 20)
        Me.TxtTotal.TabIndex = 17
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(452, 70)
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
        Me.txtSubTotal.Location = New System.Drawing.Point(567, 67)
        Me.txtSubTotal.MaxLength = 50
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.ReadOnly = True
        Me.txtSubTotal.Size = New System.Drawing.Size(285, 20)
        Me.txtSubTotal.TabIndex = 13
        Me.txtSubTotal.Text = "0"
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(452, 22)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(46, 13)
        Me.Label17.TabIndex = 146
        Me.Label17.Text = "Ship Via"
        '
        'TxtShipVia
        '
        Me.TxtShipVia.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtShipVia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtShipVia.Location = New System.Drawing.Point(567, 19)
        Me.TxtShipVia.MaxLength = 50
        Me.TxtShipVia.Name = "TxtShipVia"
        Me.TxtShipVia.ReadOnly = True
        Me.TxtShipVia.Size = New System.Drawing.Size(285, 20)
        Me.TxtShipVia.TabIndex = 10
        '
        'TxtRemark
        '
        Me.TxtRemark.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemark.Location = New System.Drawing.Point(567, 194)
        Me.TxtRemark.MaxLength = 50
        Me.TxtRemark.Name = "TxtRemark"
        Me.TxtRemark.ReadOnly = True
        Me.TxtRemark.Size = New System.Drawing.Size(285, 20)
        Me.TxtRemark.TabIndex = 21
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(452, 197)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 13)
        Me.Label6.TabIndex = 143
        Me.Label6.Text = "Remarks"
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(10, 172)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(55, 13)
        Me.Label23.TabIndex = 141
        Me.Label23.Text = "Phone No"
        '
        'txtPhoneNo
        '
        Me.txtPhoneNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPhoneNo.Location = New System.Drawing.Point(111, 169)
        Me.txtPhoneNo.MaxLength = 50
        Me.txtPhoneNo.Name = "txtPhoneNo"
        Me.txtPhoneNo.ReadOnly = True
        Me.txtPhoneNo.Size = New System.Drawing.Size(319, 20)
        Me.txtPhoneNo.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(10, 146)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 13)
        Me.Label7.TabIndex = 138
        Me.Label7.Text = "Contact Person"
        '
        'txtContactPerson
        '
        Me.txtContactPerson.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtContactPerson.Location = New System.Drawing.Point(111, 143)
        Me.txtContactPerson.MaxLength = 150
        Me.txtContactPerson.Name = "txtContactPerson"
        Me.txtContactPerson.ReadOnly = True
        Me.txtContactPerson.Size = New System.Drawing.Size(319, 20)
        Me.txtContactPerson.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 119)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(73, 13)
        Me.Label5.TabIndex = 136
        Me.Label5.Text = "Address Code"
        '
        'cboSupplierAddress
        '
        Me.cboSupplierAddress.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSupplierAddress.Enabled = False
        Me.cboSupplierAddress.FormattingEnabled = True
        Me.cboSupplierAddress.Location = New System.Drawing.Point(111, 116)
        Me.cboSupplierAddress.Name = "cboSupplierAddress"
        Me.cboSupplierAddress.Size = New System.Drawing.Size(319, 21)
        Me.cboSupplierAddress.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 97)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 134
        Me.Label2.Text = "Supplier Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 72)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Supplier Code"
        '
        'txtPrimarySupplier
        '
        Me.txtPrimarySupplier.Location = New System.Drawing.Point(147, 67)
        Me.txtPrimarySupplier.Name = "txtPrimarySupplier"
        Me.txtPrimarySupplier.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimarySupplier.TabIndex = 133
        Me.txtPrimarySupplier.TabStop = False
        Me.txtPrimarySupplier.Visible = False
        '
        'txtSupplierCode
        '
        Me.txtSupplierCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSupplierCode.Location = New System.Drawing.Point(111, 67)
        Me.txtSupplierCode.MaxLength = 50
        Me.txtSupplierCode.Name = "txtSupplierCode"
        Me.txtSupplierCode.ReadOnly = True
        Me.txtSupplierCode.Size = New System.Drawing.Size(319, 20)
        Me.txtSupplierCode.TabIndex = 2
        '
        'TxtSupplierName
        '
        Me.TxtSupplierName.Location = New System.Drawing.Point(111, 94)
        Me.TxtSupplierName.MaxLength = 150
        Me.TxtSupplierName.Name = "TxtSupplierName"
        Me.TxtSupplierName.ReadOnly = True
        Me.TxtSupplierName.Size = New System.Drawing.Size(319, 20)
        Me.TxtSupplierName.TabIndex = 3
        Me.TxtSupplierName.TabStop = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryCurrency.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(567, 42)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(26, 20)
        Me.txtPrimaryCurrency.TabIndex = 126
        Me.txtPrimaryCurrency.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCurrency.Visible = False
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(452, 46)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(77, 13)
        Me.Label22.TabIndex = 125
        Me.Label22.Text = "Currency Code"
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrencyCode.Location = New System.Drawing.Point(567, 42)
        Me.txtCurrencyCode.MaxLength = 50
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.ReadOnly = True
        Me.txtCurrencyCode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtCurrencyCode.Size = New System.Drawing.Size(285, 20)
        Me.txtCurrencyCode.TabIndex = 12
        Me.txtCurrencyCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'llbAuthorizationDetail
        '
        Me.llbAuthorizationDetail.AutoSize = True
        Me.llbAuthorizationDetail.Location = New System.Drawing.Point(373, 197)
        Me.llbAuthorizationDetail.Name = "llbAuthorizationDetail"
        Me.llbAuthorizationDetail.Size = New System.Drawing.Size(57, 13)
        Me.llbAuthorizationDetail.TabIndex = 9
        Me.llbAuthorizationDetail.TabStop = True
        Me.llbAuthorizationDetail.Text = "view detail"
        '
        'txtAutohirizationStatus
        '
        Me.txtAutohirizationStatus.Location = New System.Drawing.Point(111, 195)
        Me.txtAutohirizationStatus.MaxLength = 50
        Me.txtAutohirizationStatus.Name = "txtAutohirizationStatus"
        Me.txtAutohirizationStatus.ReadOnly = True
        Me.txtAutohirizationStatus.Size = New System.Drawing.Size(256, 20)
        Me.txtAutohirizationStatus.TabIndex = 8
        Me.txtAutohirizationStatus.TabStop = False
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(10, 198)
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
        Me.cboPaymentDueDateReference.Items.AddRange(New Object() {"Purchase Order Date", "Invoice Date", "Bill Of Leading"})
        Me.cboPaymentDueDateReference.Location = New System.Drawing.Point(684, 168)
        Me.cboPaymentDueDateReference.Name = "cboPaymentDueDateReference"
        Me.cboPaymentDueDateReference.Size = New System.Drawing.Size(168, 21)
        Me.cboPaymentDueDateReference.TabIndex = 20
        '
        'Label20
        '
        Me.Label20.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(623, 171)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(55, 13)
        Me.Label20.TabIndex = 118
        Me.Label20.Text = "days from "
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(452, 171)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(97, 13)
        Me.Label18.TabIndex = 117
        Me.Label18.Text = "Payment Due Date"
        '
        'txtPaymentDueDate
        '
        Me.txtPaymentDueDate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPaymentDueDate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaymentDueDate.Location = New System.Drawing.Point(567, 168)
        Me.txtPaymentDueDate.MaxLength = 50
        Me.txtPaymentDueDate.Name = "txtPaymentDueDate"
        Me.txtPaymentDueDate.ReadOnly = True
        Me.txtPaymentDueDate.Size = New System.Drawing.Size(50, 20)
        Me.txtPaymentDueDate.TabIndex = 19
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(452, 143)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(75, 13)
        Me.Label8.TabIndex = 113
        Me.Label8.Text = "Payment Term"
        '
        'txtPaymentTerm
        '
        Me.txtPaymentTerm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPaymentTerm.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPaymentTerm.Location = New System.Drawing.Point(567, 140)
        Me.txtPaymentTerm.MaxLength = 50
        Me.txtPaymentTerm.Name = "txtPaymentTerm"
        Me.txtPaymentTerm.ReadOnly = True
        Me.txtPaymentTerm.Size = New System.Drawing.Size(285, 20)
        Me.txtPaymentTerm.TabIndex = 18
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Enabled = False
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(111, 42)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(319, 20)
        Me.dtpDate.TabIndex = 1
        '
        'TxtPrimaryStockIn
        '
        Me.TxtPrimaryStockIn.Location = New System.Drawing.Point(151, 19)
        Me.TxtPrimaryStockIn.Name = "TxtPrimaryStockIn"
        Me.TxtPrimaryStockIn.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryStockIn.TabIndex = 47
        Me.TxtPrimaryStockIn.TabStop = False
        Me.TxtPrimaryStockIn.Visible = False
        '
        'TxtStockInNo
        '
        Me.TxtStockInNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtStockInNo.Location = New System.Drawing.Point(111, 19)
        Me.TxtStockInNo.MaxLength = 50
        Me.TxtStockInNo.Name = "TxtStockInNo"
        Me.TxtStockInNo.ReadOnly = True
        Me.TxtStockInNo.Size = New System.Drawing.Size(256, 20)
        Me.TxtStockInNo.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 46
        Me.Label3.Text = "Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(98, 13)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "Purchase Order No"
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(486, 275)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 166
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'StatusStrip2
        '
        Me.StatusStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip2.Location = New System.Drawing.Point(0, 537)
        Me.StatusStrip2.Name = "StatusStrip2"
        Me.StatusStrip2.Size = New System.Drawing.Size(896, 22)
        Me.StatusStrip2.TabIndex = 54
        Me.StatusStrip2.Text = "StatusStrip2"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(220, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(220, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(220, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(220, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'ToolStrip1
        '
        Me.ToolStrip1.AutoSize = False
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(30, 30)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbExit, Me.tsbAuthorize, Me.tsbRemove, Me.tsbShowRemoved})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 507)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStrip1.Size = New System.Drawing.Size(896, 30)
        Me.ToolStrip1.TabIndex = 55
        '
        'tsbExit
        '
        Me.tsbExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbExit.Name = "tsbExit"
        Me.tsbExit.Size = New System.Drawing.Size(29, 27)
        Me.tsbExit.Text = "E&xit"
        Me.tsbExit.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tsbAuthorize
        '
        Me.tsbAuthorize.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbAuthorize.Name = "tsbAuthorize"
        Me.tsbAuthorize.Size = New System.Drawing.Size(62, 27)
        Me.tsbAuthorize.Text = "Authorize"
        Me.tsbAuthorize.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tsbRemove
        '
        Me.tsbRemove.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbRemove.Name = "tsbRemove"
        Me.tsbRemove.Size = New System.Drawing.Size(54, 27)
        Me.tsbRemove.Text = "Remove"
        Me.tsbRemove.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'tsbShowRemoved
        '
        Me.tsbShowRemoved.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbShowRemoved.Name = "tsbShowRemoved"
        Me.tsbShowRemoved.Size = New System.Drawing.Size(93, 27)
        Me.tsbShowRemoved.Text = "Show Removed"
        Me.tsbShowRemoved.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage
        '
        'FrmTransaksiPOOtorisasi
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(896, 559)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.StatusStrip2)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.TabControl1)
        Me.Location = New System.Drawing.Point(6, 219)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(912, 597)
        Me.Name = "FrmTransaksiPOOtorisasi"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Order Authorization Manager"
        Me.TabBlmOtorisasi.ResumeLayout(False)
        Me.TabBlmOtorisasi.PerformLayout()
        Me.TabCancel.ResumeLayout(False)
        Me.TabCancel.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip2.ResumeLayout(False)
        Me.StatusStrip2.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents chkBelumOtorisasi As System.Windows.Forms.CheckedListBox
    Friend WithEvents TabBlmOtorisasi As System.Windows.Forms.TabPage
    Friend WithEvents TabCancel As System.Windows.Forms.TabPage
    Friend WithEvents chkSudahOtorisasi As System.Windows.Forms.CheckedListBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents txtTaxAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents chkTax As System.Windows.Forms.CheckBox
    Friend WithEvents txtTax As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TxtShipVia As System.Windows.Forms.TextBox
    Friend WithEvents TxtRemark As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtPhoneNo As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtContactPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboSupplierAddress As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimarySupplier As System.Windows.Forms.TextBox
    Friend WithEvents txtSupplierCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtSupplierName As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents llbAuthorizationDetail As System.Windows.Forms.LinkLabel
    Friend WithEvents txtAutohirizationStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cboPaymentDueDateReference As System.Windows.Forms.ComboBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtPaymentDueDate As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPaymentTerm As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryStockIn As System.Windows.Forms.TextBox
    Friend WithEvents TxtStockInNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip2 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsbExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbAuthorize As System.Windows.Forms.ToolStripButton
    Friend WithEvents llbUncheckAllAuthorize As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAllAuthorize As System.Windows.Forms.LinkLabel
    Friend WithEvents llbUncheckAllCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents llbCheckAllCancel As System.Windows.Forms.LinkLabel
    Friend WithEvents txtSupplierAddIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents tsbShowRemoved As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsbRemove As System.Windows.Forms.ToolStripButton
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDesc As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatioReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIPONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryCont As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColRemark As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryIPO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUserSecondaryUnits As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscFromPerct As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDown As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceDownPerc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColToleranceUpPerc As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
