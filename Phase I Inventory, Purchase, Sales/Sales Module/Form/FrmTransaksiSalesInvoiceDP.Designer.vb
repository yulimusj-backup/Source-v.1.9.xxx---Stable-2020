<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesInvoiceDP
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
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.txtTax = New System.Windows.Forms.TextBox
        Me.TxtTotal = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.btnEdit = New System.Windows.Forms.Button
        Me.txtTaxExchangeRate = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.txtExhangeRateToIDR = New System.Windows.Forms.TextBox
        Me.txtExchangeRateToUSD = New System.Windows.Forms.TextBox
        Me.txtTaxAmount = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.chkTax = New System.Windows.Forms.CheckBox
        Me.txtSubTotal = New System.Windows.Forms.TextBox
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel
        Me.txtRevisedInvoice = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtNoFakturPajak = New System.Windows.Forms.TextBox
        Me.txtDPPNilaiLain = New System.Windows.Forms.TextBox
        Me.chkRevisedInvoice = New System.Windows.Forms.CheckBox
        Me.LlbViewDetailTax = New System.Windows.Forms.LinkLabel
        Me.cboTaxtype = New System.Windows.Forms.ComboBox
        Me.cboNIKPASS = New System.Windows.Forms.ComboBox
        Me.txtNoNIKPass = New System.Windows.Forms.TextBox
        Me.txtTaxAmount1 = New System.Windows.Forms.TextBox
        Me.LblPrintStatus = New System.Windows.Forms.Label
        Me.txtPrimaryDPAccount = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboDPAccount = New System.Windows.Forms.ComboBox
        Me.cboPrimaryDPAccount = New System.Windows.Forms.ComboBox
        Me.lblCurrTax = New System.Windows.Forms.Label
        Me.txtPrimaryCurrTax = New System.Windows.Forms.TextBox
        Me.cboCurrTax = New System.Windows.Forms.ComboBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.cboPrimaryCurrTax = New System.Windows.Forms.ComboBox
        Me.txtPrimaryAccount = New System.Windows.Forms.TextBox
        Me.Label39 = New System.Windows.Forms.Label
        Me.cboAccount = New System.Windows.Forms.ComboBox
        Me.cboPrimaryAccount = New System.Windows.Forms.ComboBox
        Me.txtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPrimarySalesPerson = New System.Windows.Forms.TextBox
        Me.TxtSalesPersonCode = New System.Windows.Forms.TextBox
        Me.TxtSalesPersonName = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.TxtPaymentTerm = New System.Windows.Forms.TextBox
        Me.llbUpDateNPWP = New System.Windows.Forms.LinkLabel
        Me.Label16 = New System.Windows.Forms.Label
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer
        Me.TxtNamaSigned = New System.Windows.Forms.TextBox
        Me.TxtJabatanSigned = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.llbEditNoInvoice = New System.Windows.Forms.LinkLabel
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtNoNPWP = New System.Windows.Forms.TextBox
        Me.txtCustomerAddIdentifier = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.llbAddress = New System.Windows.Forms.LinkLabel
        Me.cboCustomerAddress = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.CboFakturPajak = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtDutyStamp = New System.Windows.Forms.TextBox
        Me.btnFindSalesOrder = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtPrimarySO = New System.Windows.Forms.TextBox
        Me.txtSONo = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.dtpDate = New System.Windows.Forms.DateTimePicker
        Me.TxtPrimaryInvoice = New System.Windows.Forms.TextBox
        Me.txtInvoiceNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.txtNoFakturPajakSederhana = New System.Windows.Forms.TextBox
        Me.cboPPNDipungut = New System.Windows.Forms.ComboBox
        Me.txtNoSeriFakturPajak = New System.Windows.Forms.TextBox
        Me.llbEditNoFakturPajak = New System.Windows.Forms.LinkLabel
        Me.btnPrintTaxInvoice = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.gbo_EFaktur = New System.Windows.Forms.GroupBox
        Me.Label41 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Jalan = New System.Windows.Forms.TextBox
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.Label42 = New System.Windows.Forms.Label
        Me.Label43 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_NoTelepon = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_KodePos = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_Propinsi = New System.Windows.Forms.TextBox
        Me.Label44 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Kabupaten = New System.Windows.Forms.TextBox
        Me.Label45 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Kecamatan = New System.Windows.Forms.TextBox
        Me.Label46 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_Kelurahan = New System.Windows.Forms.TextBox
        Me.Label47 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_RW = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_RT = New System.Windows.Forms.TextBox
        Me.Label48 = New System.Windows.Forms.Label
        Me.Label49 = New System.Windows.Forms.Label
        Me.Label50 = New System.Windows.Forms.Label
        Me.Txt_E_Faktur_No = New System.Windows.Forms.TextBox
        Me.Txt_E_Faktur_Blok = New System.Windows.Forms.TextBox
        Me.Label51 = New System.Windows.Forms.Label
        Me.StatusStrip1.SuspendLayout()
        Me.gboDetail.SuspendLayout()
        Me.TLayoutPanel2.SuspendLayout()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.gbo_EFaktur.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(399, 407)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(107, 32)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrint.Location = New System.Drawing.Point(743, 407)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(4)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(107, 32)
        Me.btnPrint.TabIndex = 5
        Me.btnPrint.Text = "&Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(857, 407)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(107, 32)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(283, 407)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(107, 32)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(628, 407)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(107, 32)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'txtTax
        '
        Me.txtTax.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTax.Location = New System.Drawing.Point(767, 112)
        Me.txtTax.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTax.MaxLength = 50
        Me.txtTax.Name = "txtTax"
        Me.txtTax.ReadOnly = True
        Me.txtTax.Size = New System.Drawing.Size(103, 22)
        Me.txtTax.TabIndex = 20
        Me.txtTax.Text = "0"
        Me.txtTax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TxtTotal
        '
        Me.TxtTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtTotal.Location = New System.Drawing.Point(767, 209)
        Me.TxtTotal.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTotal.MaxLength = 50
        Me.TxtTotal.Name = "TxtTotal"
        Me.TxtTotal.ReadOnly = True
        Me.TxtTotal.Size = New System.Drawing.Size(437, 22)
        Me.TxtTotal.TabIndex = 26
        Me.TxtTotal.TabStop = False
        Me.TxtTotal.Text = "0"
        Me.TxtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(612, 85)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(102, 17)
        Me.Label15.TabIndex = 32
        Me.Label15.Text = "Down Payment"
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(612, 214)
        Me.Label24.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(40, 17)
        Me.Label24.TabIndex = 48
        Me.Label24.Text = "Total"
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(767, 322)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtRemarks.MaxLength = 50
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(437, 22)
        Me.TxtRemarks.TabIndex = 29
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(612, 327)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 17)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Remarks"
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(513, 407)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(107, 32)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'txtTaxExchangeRate
        '
        Me.txtTaxExchangeRate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxExchangeRate.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxExchangeRate.Location = New System.Drawing.Point(1051, 112)
        Me.txtTaxExchangeRate.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTaxExchangeRate.MaxLength = 18
        Me.txtTaxExchangeRate.Name = "txtTaxExchangeRate"
        Me.txtTaxExchangeRate.ReadOnly = True
        Me.txtTaxExchangeRate.Size = New System.Drawing.Size(153, 22)
        Me.txtTaxExchangeRate.TabIndex = 21
        Me.txtTaxExchangeRate.TabStop = False
        Me.txtTaxExchangeRate.Text = "1"
        Me.txtTaxExchangeRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(907, 117)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(131, 17)
        Me.Label17.TabIndex = 36
        Me.Label17.Text = "Tax Exchange Rate"
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Location = New System.Drawing.Point(488, 271)
        Me.btnFindCurrency.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(43, 25)
        Me.btnFindCurrency.TabIndex = 16
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "..."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        Me.btnFindCurrency.Visible = False
        '
        'txtExhangeRateToIDR
        '
        Me.txtExhangeRateToIDR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExhangeRateToIDR.Location = New System.Drawing.Point(308, 176)
        Me.txtExhangeRateToIDR.Margin = New System.Windows.Forms.Padding(4)
        Me.txtExhangeRateToIDR.MaxLength = 50
        Me.txtExhangeRateToIDR.Name = "txtExhangeRateToIDR"
        Me.txtExhangeRateToIDR.Size = New System.Drawing.Size(224, 22)
        Me.txtExhangeRateToIDR.TabIndex = 14
        Me.txtExhangeRateToIDR.TabStop = False
        Me.txtExhangeRateToIDR.Text = "1"
        Me.txtExhangeRateToIDR.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExhangeRateToIDR.Visible = False
        '
        'txtExchangeRateToUSD
        '
        Me.txtExchangeRateToUSD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExchangeRateToUSD.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtExchangeRateToUSD.Location = New System.Drawing.Point(308, 176)
        Me.txtExchangeRateToUSD.Margin = New System.Windows.Forms.Padding(4)
        Me.txtExchangeRateToUSD.MaxLength = 50
        Me.txtExchangeRateToUSD.Name = "txtExchangeRateToUSD"
        Me.txtExchangeRateToUSD.Size = New System.Drawing.Size(128, 22)
        Me.txtExchangeRateToUSD.TabIndex = 13
        Me.txtExchangeRateToUSD.TabStop = False
        Me.txtExchangeRateToUSD.Text = "1"
        Me.txtExchangeRateToUSD.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtExchangeRateToUSD.Visible = False
        '
        'txtTaxAmount
        '
        Me.txtTaxAmount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxAmount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxAmount.Location = New System.Drawing.Point(911, 144)
        Me.txtTaxAmount.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTaxAmount.MaxLength = 50
        Me.txtTaxAmount.Name = "txtTaxAmount"
        Me.txtTaxAmount.ReadOnly = True
        Me.txtTaxAmount.Size = New System.Drawing.Size(293, 22)
        Me.txtTaxAmount.TabIndex = 42
        Me.txtTaxAmount.TabStop = False
        Me.txtTaxAmount.Text = "0"
        Me.txtTaxAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(879, 117)
        Me.Label25.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(20, 17)
        Me.Label25.TabIndex = 35
        Me.Label25.Text = "%"
        '
        'chkTax
        '
        Me.chkTax.AutoSize = True
        Me.chkTax.Enabled = False
        Me.chkTax.Location = New System.Drawing.Point(612, 114)
        Me.chkTax.Margin = New System.Windows.Forms.Padding(4)
        Me.chkTax.Name = "chkTax"
        Me.chkTax.Size = New System.Drawing.Size(106, 21)
        Me.chkTax.TabIndex = 33
        Me.chkTax.Text = "Tax Exclude"
        Me.chkTax.UseVisualStyleBackColor = True
        '
        'txtSubTotal
        '
        Me.txtSubTotal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubTotal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSubTotal.Location = New System.Drawing.Point(767, 80)
        Me.txtSubTotal.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSubTotal.MaxLength = 18
        Me.txtSubTotal.Name = "txtSubTotal"
        Me.txtSubTotal.Size = New System.Drawing.Size(437, 22)
        Me.txtSubTotal.TabIndex = 19
        Me.txtSubTotal.Text = "0"
        Me.txtSubTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(309, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(309, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By: "
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(309, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By: "
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(309, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date: "
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 469)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1259, 25)
        Me.StatusStrip1.TabIndex = 47
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TLayoutPanel2)
        Me.gboDetail.Controls.Add(Me.txtDPPNilaiLain)
        Me.gboDetail.Controls.Add(Me.chkRevisedInvoice)
        Me.gboDetail.Controls.Add(Me.LlbViewDetailTax)
        Me.gboDetail.Controls.Add(Me.cboTaxtype)
        Me.gboDetail.Controls.Add(Me.cboNIKPASS)
        Me.gboDetail.Controls.Add(Me.txtNoNIKPass)
        Me.gboDetail.Controls.Add(Me.txtTaxAmount1)
        Me.gboDetail.Controls.Add(Me.LblPrintStatus)
        Me.gboDetail.Controls.Add(Me.txtPrimaryDPAccount)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.cboDPAccount)
        Me.gboDetail.Controls.Add(Me.cboPrimaryDPAccount)
        Me.gboDetail.Controls.Add(Me.lblCurrTax)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCurrTax)
        Me.gboDetail.Controls.Add(Me.cboCurrTax)
        Me.gboDetail.Controls.Add(Me.Label29)
        Me.gboDetail.Controls.Add(Me.Label18)
        Me.gboDetail.Controls.Add(Me.cboPrimaryCurrTax)
        Me.gboDetail.Controls.Add(Me.txtPrimaryAccount)
        Me.gboDetail.Controls.Add(Me.Label39)
        Me.gboDetail.Controls.Add(Me.cboAccount)
        Me.gboDetail.Controls.Add(Me.cboPrimaryAccount)
        Me.gboDetail.Controls.Add(Me.txtCurrencyCode)
        Me.gboDetail.Controls.Add(Me.Txt_E_Faktur_KodeJenisTransaksiPajak)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.TxtPrimarySalesPerson)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonCode)
        Me.gboDetail.Controls.Add(Me.TxtSalesPersonName)
        Me.gboDetail.Controls.Add(Me.Label19)
        Me.gboDetail.Controls.Add(Me.TxtPaymentTerm)
        Me.gboDetail.Controls.Add(Me.llbUpDateNPWP)
        Me.gboDetail.Controls.Add(Me.TxtRemarks)
        Me.gboDetail.Controls.Add(Me.Label16)
        Me.gboDetail.Controls.Add(Me.SplitContainer3)
        Me.gboDetail.Controls.Add(Me.Label14)
        Me.gboDetail.Controls.Add(Me.Label13)
        Me.gboDetail.Controls.Add(Me.llbEditNoInvoice)
        Me.gboDetail.Controls.Add(Me.TxtReadOnly)
        Me.gboDetail.Controls.Add(Me.Label10)
        Me.gboDetail.Controls.Add(Me.txtNoNPWP)
        Me.gboDetail.Controls.Add(Me.txtCustomerAddIdentifier)
        Me.gboDetail.Controls.Add(Me.Label9)
        Me.gboDetail.Controls.Add(Me.llbAddress)
        Me.gboDetail.Controls.Add(Me.cboCustomerAddress)
        Me.gboDetail.Controls.Add(Me.Label11)
        Me.gboDetail.Controls.Add(Me.CboFakturPajak)
        Me.gboDetail.Controls.Add(Me.Label12)
        Me.gboDetail.Controls.Add(Me.txtDutyStamp)
        Me.gboDetail.Controls.Add(Me.btnFindSalesOrder)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.txtPrimarySO)
        Me.gboDetail.Controls.Add(Me.txtSONo)
        Me.gboDetail.Controls.Add(Me.txtExhangeRateToIDR)
        Me.gboDetail.Controls.Add(Me.txtExchangeRateToUSD)
        Me.gboDetail.Controls.Add(Me.txtTaxExchangeRate)
        Me.gboDetail.Controls.Add(Me.Label17)
        Me.gboDetail.Controls.Add(Me.btnFindCurrency)
        Me.gboDetail.Controls.Add(Me.txtTaxAmount)
        Me.gboDetail.Controls.Add(Me.Label25)
        Me.gboDetail.Controls.Add(Me.chkTax)
        Me.gboDetail.Controls.Add(Me.txtTax)
        Me.gboDetail.Controls.Add(Me.Label24)
        Me.gboDetail.Controls.Add(Me.TxtTotal)
        Me.gboDetail.Controls.Add(Me.Label15)
        Me.gboDetail.Controls.Add(Me.txtSubTotal)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCustomer)
        Me.gboDetail.Controls.Add(Me.txtCustomerCode)
        Me.gboDetail.Controls.Add(Me.TxtCustomerName)
        Me.gboDetail.Controls.Add(Me.txtPrimaryCurrency)
        Me.gboDetail.Controls.Add(Me.Label22)
        Me.gboDetail.Controls.Add(Me.dtpDate)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryInvoice)
        Me.gboDetail.Controls.Add(Me.txtInvoiceNo)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.TextBox1)
        Me.gboDetail.Controls.Add(Me.txtNoFakturPajakSederhana)
        Me.gboDetail.Controls.Add(Me.cboPPNDipungut)
        Me.gboDetail.Controls.Add(Me.txtNoSeriFakturPajak)
        Me.gboDetail.Controls.Add(Me.llbEditNoFakturPajak)
        Me.gboDetail.Location = New System.Drawing.Point(11, 0)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(1213, 399)
        Me.gboDetail.TabIndex = 7
        Me.gboDetail.TabStop = False
        '
        'TLayoutPanel2
        '
        Me.TLayoutPanel2.ColumnCount = 3
        Me.TLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle)
        Me.TLayoutPanel2.Controls.Add(Me.txtRevisedInvoice, 0, 0)
        Me.TLayoutPanel2.Controls.Add(Me.Label7, 1, 0)
        Me.TLayoutPanel2.Controls.Add(Me.txtNoFakturPajak, 2, 0)
        Me.TLayoutPanel2.Location = New System.Drawing.Point(202, 236)
        Me.TLayoutPanel2.Name = "TLayoutPanel2"
        Me.TLayoutPanel2.RowCount = 1
        Me.TLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TLayoutPanel2.Size = New System.Drawing.Size(282, 29)
        Me.TLayoutPanel2.TabIndex = 12
        '
        'txtRevisedInvoice
        '
        Me.txtRevisedInvoice.Location = New System.Drawing.Point(4, 4)
        Me.txtRevisedInvoice.Margin = New System.Windows.Forms.Padding(4)
        Me.txtRevisedInvoice.MaxLength = 2
        Me.txtRevisedInvoice.Name = "txtRevisedInvoice"
        Me.txtRevisedInvoice.Size = New System.Drawing.Size(59, 22)
        Me.txtRevisedInvoice.TabIndex = 12
        Me.txtRevisedInvoice.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label7.Location = New System.Drawing.Point(71, 0)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(13, 29)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "-"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtNoFakturPajak
        '
        Me.txtNoFakturPajak.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNoFakturPajak.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNoFakturPajak.Location = New System.Drawing.Point(92, 4)
        Me.txtNoFakturPajak.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNoFakturPajak.MaxLength = 50
        Me.txtNoFakturPajak.Name = "txtNoFakturPajak"
        Me.txtNoFakturPajak.Size = New System.Drawing.Size(186, 22)
        Me.txtNoFakturPajak.TabIndex = 13
        '
        'txtDPPNilaiLain
        '
        Me.txtDPPNilaiLain.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDPPNilaiLain.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDPPNilaiLain.Location = New System.Drawing.Point(16, 303)
        Me.txtDPPNilaiLain.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDPPNilaiLain.MaxLength = 50
        Me.txtDPPNilaiLain.Name = "txtDPPNilaiLain"
        Me.txtDPPNilaiLain.ReadOnly = True
        Me.txtDPPNilaiLain.Size = New System.Drawing.Size(99, 22)
        Me.txtDPPNilaiLain.TabIndex = 319
        Me.txtDPPNilaiLain.TabStop = False
        Me.txtDPPNilaiLain.Text = "0"
        Me.txtDPPNilaiLain.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDPPNilaiLain.Visible = False
        '
        'chkRevisedInvoice
        '
        Me.chkRevisedInvoice.AutoSize = True
        Me.chkRevisedInvoice.Location = New System.Drawing.Point(17, 331)
        Me.chkRevisedInvoice.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkRevisedInvoice.Name = "chkRevisedInvoice"
        Me.chkRevisedInvoice.Size = New System.Drawing.Size(108, 21)
        Me.chkRevisedInvoice.TabIndex = 12
        Me.chkRevisedInvoice.Text = "Revised Tax"
        Me.chkRevisedInvoice.UseVisualStyleBackColor = True
        Me.chkRevisedInvoice.Visible = False
        '
        'LlbViewDetailTax
        '
        Me.LlbViewDetailTax.AutoSize = True
        Me.LlbViewDetailTax.Location = New System.Drawing.Point(491, 242)
        Me.LlbViewDetailTax.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LlbViewDetailTax.Name = "LlbViewDetailTax"
        Me.LlbViewDetailTax.Size = New System.Drawing.Size(37, 17)
        Me.LlbViewDetailTax.TabIndex = 14
        Me.LlbViewDetailTax.TabStop = True
        Me.LlbViewDetailTax.Text = "View"
        '
        'cboTaxtype
        '
        Me.cboTaxtype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTaxtype.FormattingEnabled = True
        Me.cboTaxtype.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10"})
        Me.cboTaxtype.Location = New System.Drawing.Point(140, 239)
        Me.cboTaxtype.Margin = New System.Windows.Forms.Padding(4)
        Me.cboTaxtype.Name = "cboTaxtype"
        Me.cboTaxtype.Size = New System.Drawing.Size(55, 24)
        Me.cboTaxtype.TabIndex = 11
        '
        'cboNIKPASS
        '
        Me.cboNIKPASS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboNIKPASS.FormattingEnabled = True
        Me.cboNIKPASS.Items.AddRange(New Object() {"NIK", "PASSPORT", "Other ID"})
        Me.cboNIKPASS.Location = New System.Drawing.Point(17, 204)
        Me.cboNIKPASS.Margin = New System.Windows.Forms.Padding(4)
        Me.cboNIKPASS.Name = "cboNIKPASS"
        Me.cboNIKPASS.Size = New System.Drawing.Size(112, 24)
        Me.cboNIKPASS.TabIndex = 9
        Me.cboNIKPASS.Visible = False
        '
        'txtNoNIKPass
        '
        Me.txtNoNIKPass.Location = New System.Drawing.Point(140, 208)
        Me.txtNoNIKPass.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNoNIKPass.MaxLength = 25
        Me.txtNoNIKPass.Name = "txtNoNIKPass"
        Me.txtNoNIKPass.ReadOnly = True
        Me.txtNoNIKPass.Size = New System.Drawing.Size(392, 22)
        Me.txtNoNIKPass.TabIndex = 10
        Me.txtNoNIKPass.TabStop = False
        Me.txtNoNIKPass.Visible = False
        '
        'txtTaxAmount1
        '
        Me.txtTaxAmount1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaxAmount1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTaxAmount1.Location = New System.Drawing.Point(911, 144)
        Me.txtTaxAmount1.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTaxAmount1.MaxLength = 50
        Me.txtTaxAmount1.Name = "txtTaxAmount1"
        Me.txtTaxAmount1.ReadOnly = True
        Me.txtTaxAmount1.Size = New System.Drawing.Size(293, 22)
        Me.txtTaxAmount1.TabIndex = 24
        Me.txtTaxAmount1.TabStop = False
        Me.txtTaxAmount1.Text = "0"
        Me.txtTaxAmount1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblPrintStatus
        '
        Me.LblPrintStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblPrintStatus.AutoSize = True
        Me.LblPrintStatus.Location = New System.Drawing.Point(8, 367)
        Me.LblPrintStatus.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LblPrintStatus.Name = "LblPrintStatus"
        Me.LblPrintStatus.Size = New System.Drawing.Size(94, 17)
        Me.LblPrintStatus.TabIndex = 316
        Me.LblPrintStatus.Text = "Print Status: -"
        '
        'txtPrimaryDPAccount
        '
        Me.txtPrimaryDPAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPrimaryDPAccount.Location = New System.Drawing.Point(689, 47)
        Me.txtPrimaryDPAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryDPAccount.MaxLength = 50
        Me.txtPrimaryDPAccount.Name = "txtPrimaryDPAccount"
        Me.txtPrimaryDPAccount.Size = New System.Drawing.Size(32, 22)
        Me.txtPrimaryDPAccount.TabIndex = 30
        Me.txtPrimaryDPAccount.TabStop = False
        Me.txtPrimaryDPAccount.Text = "-1"
        Me.txtPrimaryDPAccount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryDPAccount.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(612, 52)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 17)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "DP Account"
        '
        'cboDPAccount
        '
        Me.cboDPAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboDPAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDPAccount.FormattingEnabled = True
        Me.cboDPAccount.Location = New System.Drawing.Point(767, 47)
        Me.cboDPAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.cboDPAccount.Name = "cboDPAccount"
        Me.cboDPAccount.Size = New System.Drawing.Size(437, 24)
        Me.cboDPAccount.TabIndex = 18
        '
        'cboPrimaryDPAccount
        '
        Me.cboPrimaryDPAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryDPAccount.FormattingEnabled = True
        Me.cboPrimaryDPAccount.Location = New System.Drawing.Point(772, 47)
        Me.cboPrimaryDPAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.cboPrimaryDPAccount.Name = "cboPrimaryDPAccount"
        Me.cboPrimaryDPAccount.Size = New System.Drawing.Size(40, 24)
        Me.cboPrimaryDPAccount.TabIndex = 31
        '
        'lblCurrTax
        '
        Me.lblCurrTax.AutoSize = True
        Me.lblCurrTax.Location = New System.Drawing.Point(689, 180)
        Me.lblCurrTax.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCurrTax.Name = "lblCurrTax"
        Me.lblCurrTax.Size = New System.Drawing.Size(31, 17)
        Me.lblCurrTax.TabIndex = 44
        Me.lblCurrTax.Text = "IDR"
        '
        'txtPrimaryCurrTax
        '
        Me.txtPrimaryCurrTax.Location = New System.Drawing.Point(612, 145)
        Me.txtPrimaryCurrTax.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryCurrTax.Name = "txtPrimaryCurrTax"
        Me.txtPrimaryCurrTax.Size = New System.Drawing.Size(19, 22)
        Me.txtPrimaryCurrTax.TabIndex = 38
        Me.txtPrimaryCurrTax.TabStop = False
        Me.txtPrimaryCurrTax.Text = "-1"
        Me.txtPrimaryCurrTax.Visible = False
        '
        'cboCurrTax
        '
        Me.cboCurrTax.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCurrTax.FormattingEnabled = True
        Me.cboCurrTax.Location = New System.Drawing.Point(649, 145)
        Me.cboCurrTax.Margin = New System.Windows.Forms.Padding(4)
        Me.cboCurrTax.Name = "cboCurrTax"
        Me.cboCurrTax.Size = New System.Drawing.Size(97, 24)
        Me.cboCurrTax.TabIndex = 22
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.Location = New System.Drawing.Point(612, 180)
        Me.Label29.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(81, 17)
        Me.Label29.TabIndex = 43
        Me.Label29.Text = "Stamp Duty"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(612, 150)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(31, 17)
        Me.Label18.TabIndex = 40
        Me.Label18.Text = "Tax"
        '
        'cboPrimaryCurrTax
        '
        Me.cboPrimaryCurrTax.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryCurrTax.FormattingEnabled = True
        Me.cboPrimaryCurrTax.Location = New System.Drawing.Point(692, 145)
        Me.cboPrimaryCurrTax.Margin = New System.Windows.Forms.Padding(4)
        Me.cboPrimaryCurrTax.Name = "cboPrimaryCurrTax"
        Me.cboPrimaryCurrTax.Size = New System.Drawing.Size(55, 24)
        Me.cboPrimaryCurrTax.TabIndex = 40
        '
        'txtPrimaryAccount
        '
        Me.txtPrimaryAccount.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPrimaryAccount.Location = New System.Drawing.Point(689, 15)
        Me.txtPrimaryAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryAccount.MaxLength = 50
        Me.txtPrimaryAccount.Name = "txtPrimaryAccount"
        Me.txtPrimaryAccount.Size = New System.Drawing.Size(32, 22)
        Me.txtPrimaryAccount.TabIndex = 27
        Me.txtPrimaryAccount.TabStop = False
        Me.txtPrimaryAccount.Text = "-1"
        Me.txtPrimaryAccount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryAccount.Visible = False
        '
        'Label39
        '
        Me.Label39.AutoSize = True
        Me.Label39.Location = New System.Drawing.Point(612, 20)
        Me.Label39.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(82, 17)
        Me.Label39.TabIndex = 26
        Me.Label39.Text = "AR Account"
        '
        'cboAccount
        '
        Me.cboAccount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAccount.FormattingEnabled = True
        Me.cboAccount.Location = New System.Drawing.Point(767, 15)
        Me.cboAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.cboAccount.Name = "cboAccount"
        Me.cboAccount.Size = New System.Drawing.Size(437, 24)
        Me.cboAccount.TabIndex = 17
        '
        'cboPrimaryAccount
        '
        Me.cboPrimaryAccount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryAccount.FormattingEnabled = True
        Me.cboPrimaryAccount.Location = New System.Drawing.Point(772, 15)
        Me.cboPrimaryAccount.Margin = New System.Windows.Forms.Padding(4)
        Me.cboPrimaryAccount.Name = "cboPrimaryAccount"
        Me.cboPrimaryAccount.Size = New System.Drawing.Size(40, 24)
        Me.cboPrimaryAccount.TabIndex = 29
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCurrencyCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrencyCode.Location = New System.Drawing.Point(140, 273)
        Me.txtCurrencyCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCurrencyCode.MaxLength = 50
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.ReadOnly = True
        Me.txtCurrencyCode.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtCurrencyCode.Size = New System.Drawing.Size(339, 22)
        Me.txtCurrencyCode.TabIndex = 15
        Me.txtCurrencyCode.TabStop = False
        Me.txtCurrencyCode.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Txt_E_Faktur_KodeJenisTransaksiPajak
        '
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Location = New System.Drawing.Point(500, 144)
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Name = "Txt_E_Faktur_KodeJenisTransaksiPajak"
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Size = New System.Drawing.Size(32, 22)
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.TabIndex = 12
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.TabStop = False
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Text = "-1"
        Me.Txt_E_Faktur_KodeJenisTransaksiPajak.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 149)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 17)
        Me.Label2.TabIndex = 315
        Me.Label2.Text = "Sales Person"
        '
        'TxtPrimarySalesPerson
        '
        Me.TxtPrimarySalesPerson.Location = New System.Drawing.Point(140, 144)
        Me.TxtPrimarySalesPerson.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimarySalesPerson.Name = "TxtPrimarySalesPerson"
        Me.TxtPrimarySalesPerson.Size = New System.Drawing.Size(32, 22)
        Me.TxtPrimarySalesPerson.TabIndex = 10
        Me.TxtPrimarySalesPerson.TabStop = False
        Me.TxtPrimarySalesPerson.Visible = False
        '
        'TxtSalesPersonCode
        '
        Me.TxtSalesPersonCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesPersonCode.Location = New System.Drawing.Point(140, 144)
        Me.TxtSalesPersonCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonCode.MaxLength = 50
        Me.TxtSalesPersonCode.Name = "TxtSalesPersonCode"
        Me.TxtSalesPersonCode.ReadOnly = True
        Me.TxtSalesPersonCode.Size = New System.Drawing.Size(107, 22)
        Me.TxtSalesPersonCode.TabIndex = 6
        '
        'TxtSalesPersonName
        '
        Me.TxtSalesPersonName.Location = New System.Drawing.Point(256, 144)
        Me.TxtSalesPersonName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtSalesPersonName.MaxLength = 150
        Me.TxtSalesPersonName.Name = "TxtSalesPersonName"
        Me.TxtSalesPersonName.ReadOnly = True
        Me.TxtSalesPersonName.Size = New System.Drawing.Size(276, 22)
        Me.TxtSalesPersonName.TabIndex = 11
        Me.TxtSalesPersonName.TabStop = False
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(273, 276)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(104, 17)
        Me.Label19.TabIndex = 24
        Me.Label19.Text = "Exchange Rate"
        Me.Label19.Visible = False
        '
        'TxtPaymentTerm
        '
        Me.TxtPaymentTerm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPaymentTerm.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPaymentTerm.Location = New System.Drawing.Point(767, 241)
        Me.TxtPaymentTerm.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPaymentTerm.MaxLength = 18
        Me.TxtPaymentTerm.Name = "TxtPaymentTerm"
        Me.TxtPaymentTerm.Size = New System.Drawing.Size(388, 22)
        Me.TxtPaymentTerm.TabIndex = 27
        Me.TxtPaymentTerm.Text = "0"
        Me.TxtPaymentTerm.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'llbUpDateNPWP
        '
        Me.llbUpDateNPWP.AutoSize = True
        Me.llbUpDateNPWP.Location = New System.Drawing.Point(475, 213)
        Me.llbUpDateNPWP.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbUpDateNPWP.Name = "llbUpDateNPWP"
        Me.llbUpDateNPWP.Size = New System.Drawing.Size(54, 17)
        Me.llbUpDateNPWP.TabIndex = 17
        Me.llbUpDateNPWP.TabStop = True
        Me.llbUpDateNPWP.Text = "Update"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(612, 286)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(72, 17)
        Me.Label16.TabIndex = 53
        Me.Label16.Text = "Signed By"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer3.Location = New System.Drawing.Point(767, 271)
        Me.SplitContainer3.Margin = New System.Windows.Forms.Padding(4)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.TxtNamaSigned)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.TxtJabatanSigned)
        Me.SplitContainer3.Size = New System.Drawing.Size(439, 44)
        Me.SplitContainer3.SplitterDistance = 214
        Me.SplitContainer3.SplitterWidth = 5
        Me.SplitContainer3.TabIndex = 28
        Me.SplitContainer3.TabStop = False
        '
        'TxtNamaSigned
        '
        Me.TxtNamaSigned.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNamaSigned.Location = New System.Drawing.Point(4, 9)
        Me.TxtNamaSigned.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtNamaSigned.MaxLength = 50
        Me.TxtNamaSigned.Name = "TxtNamaSigned"
        Me.TxtNamaSigned.Size = New System.Drawing.Size(205, 22)
        Me.TxtNamaSigned.TabIndex = 0
        '
        'TxtJabatanSigned
        '
        Me.TxtJabatanSigned.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtJabatanSigned.Location = New System.Drawing.Point(5, 10)
        Me.TxtJabatanSigned.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtJabatanSigned.MaxLength = 50
        Me.TxtJabatanSigned.Name = "TxtJabatanSigned"
        Me.TxtJabatanSigned.Size = New System.Drawing.Size(208, 22)
        Me.TxtJabatanSigned.TabIndex = 0
        '
        'Label14
        '
        Me.Label14.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(1164, 246)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(40, 17)
        Me.Label14.TabIndex = 52
        Me.Label14.Text = "Days"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(612, 246)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(107, 17)
        Me.Label13.TabIndex = 50
        Me.Label13.Text = "Payment Terms"
        '
        'llbEditNoInvoice
        '
        Me.llbEditNoInvoice.AutoSize = True
        Me.llbEditNoInvoice.Location = New System.Drawing.Point(497, 20)
        Me.llbEditNoInvoice.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbEditNoInvoice.Name = "llbEditNoInvoice"
        Me.llbEditNoInvoice.Size = New System.Drawing.Size(32, 17)
        Me.llbEditNoInvoice.TabIndex = 2
        Me.llbEditNoInvoice.TabStop = True
        Me.llbEditNoInvoice.Text = "Edit"
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtReadOnly.Location = New System.Drawing.Point(140, 273)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtReadOnly.MaxLength = 50
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(89, 22)
        Me.TxtReadOnly.TabIndex = 21
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(13, 213)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(71, 17)
        Me.Label10.TabIndex = 303
        Me.Label10.Text = "No NPWP"
        '
        'txtNoNPWP
        '
        Me.txtNoNPWP.Location = New System.Drawing.Point(140, 208)
        Me.txtNoNPWP.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNoNPWP.MaxLength = 150
        Me.txtNoNPWP.Name = "txtNoNPWP"
        Me.txtNoNPWP.ReadOnly = True
        Me.txtNoNPWP.Size = New System.Drawing.Size(300, 22)
        Me.txtNoNPWP.TabIndex = 16
        Me.txtNoNPWP.TabStop = False
        '
        'txtCustomerAddIdentifier
        '
        Me.txtCustomerAddIdentifier.Location = New System.Drawing.Point(140, 176)
        Me.txtCustomerAddIdentifier.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCustomerAddIdentifier.Name = "txtCustomerAddIdentifier"
        Me.txtCustomerAddIdentifier.Size = New System.Drawing.Size(32, 22)
        Me.txtCustomerAddIdentifier.TabIndex = 13
        Me.txtCustomerAddIdentifier.TabStop = False
        Me.txtCustomerAddIdentifier.Text = "-1"
        Me.txtCustomerAddIdentifier.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(13, 181)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(97, 17)
        Me.Label9.TabIndex = 300
        Me.Label9.Text = "Address Code"
        '
        'llbAddress
        '
        Me.llbAddress.AutoSize = True
        Me.llbAddress.Location = New System.Drawing.Point(455, 181)
        Me.llbAddress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbAddress.Name = "llbAddress"
        Me.llbAddress.Size = New System.Drawing.Size(73, 17)
        Me.llbAddress.TabIndex = 8
        Me.llbAddress.TabStop = True
        Me.llbAddress.Text = "view detail"
        '
        'cboCustomerAddress
        '
        Me.cboCustomerAddress.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerAddress.FormattingEnabled = True
        Me.cboCustomerAddress.Location = New System.Drawing.Point(140, 176)
        Me.cboCustomerAddress.Margin = New System.Windows.Forms.Padding(4)
        Me.cboCustomerAddress.Name = "cboCustomerAddress"
        Me.cboCustomerAddress.Size = New System.Drawing.Size(300, 24)
        Me.cboCustomerAddress.TabIndex = 7
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(13, 245)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(109, 17)
        Me.Label11.TabIndex = 295
        Me.Label11.Text = "No Faktur Pajak"
        '
        'CboFakturPajak
        '
        Me.CboFakturPajak.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFakturPajak.FormattingEnabled = True
        Me.CboFakturPajak.Items.AddRange(New Object() {"Faktur Pajak Sederhana", "Faktur Pajak Standar"})
        Me.CboFakturPajak.Location = New System.Drawing.Point(140, 240)
        Me.CboFakturPajak.Margin = New System.Windows.Forms.Padding(4)
        Me.CboFakturPajak.Name = "CboFakturPajak"
        Me.CboFakturPajak.Size = New System.Drawing.Size(389, 24)
        Me.CboFakturPajak.TabIndex = 10
        Me.CboFakturPajak.TabStop = False
        Me.CboFakturPajak.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(13, 245)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(87, 17)
        Me.Label12.TabIndex = 293
        Me.Label12.Text = "Faktur Pajak"
        Me.Label12.Visible = False
        '
        'txtDutyStamp
        '
        Me.txtDutyStamp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDutyStamp.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDutyStamp.Location = New System.Drawing.Point(767, 177)
        Me.txtDutyStamp.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDutyStamp.MaxLength = 18
        Me.txtDutyStamp.Name = "txtDutyStamp"
        Me.txtDutyStamp.Size = New System.Drawing.Size(437, 22)
        Me.txtDutyStamp.TabIndex = 25
        Me.txtDutyStamp.Text = "0"
        Me.txtDutyStamp.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnFindSalesOrder
        '
        Me.btnFindSalesOrder.Location = New System.Drawing.Point(488, 79)
        Me.btnFindSalesOrder.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindSalesOrder.Name = "btnFindSalesOrder"
        Me.btnFindSalesOrder.Size = New System.Drawing.Size(43, 25)
        Me.btnFindSalesOrder.TabIndex = 7
        Me.btnFindSalesOrder.TabStop = False
        Me.btnFindSalesOrder.Text = "...."
        Me.btnFindSalesOrder.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(13, 84)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(106, 17)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Sales Order No"
        '
        'txtPrimarySO
        '
        Me.txtPrimarySO.Location = New System.Drawing.Point(140, 79)
        Me.txtPrimarySO.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimarySO.Name = "txtPrimarySO"
        Me.txtPrimarySO.Size = New System.Drawing.Size(32, 22)
        Me.txtPrimarySO.TabIndex = 6
        Me.txtPrimarySO.TabStop = False
        Me.txtPrimarySO.Visible = False
        '
        'txtSONo
        '
        Me.txtSONo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSONo.Location = New System.Drawing.Point(140, 79)
        Me.txtSONo.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSONo.MaxLength = 50
        Me.txtSONo.Name = "txtSONo"
        Me.txtSONo.Size = New System.Drawing.Size(339, 22)
        Me.txtSONo.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 116)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 17)
        Me.Label1.TabIndex = 132
        Me.Label1.Text = "Customer"
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(140, 111)
        Me.txtPrimaryCustomer.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(32, 22)
        Me.txtPrimaryCustomer.TabIndex = 8
        Me.txtPrimaryCustomer.TabStop = False
        Me.txtPrimaryCustomer.Visible = False
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(140, 111)
        Me.txtCustomerCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCustomerCode.MaxLength = 50
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.ReadOnly = True
        Me.txtCustomerCode.Size = New System.Drawing.Size(107, 22)
        Me.txtCustomerCode.TabIndex = 5
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Location = New System.Drawing.Point(256, 111)
        Me.TxtCustomerName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCustomerName.MaxLength = 150
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(276, 22)
        Me.TxtCustomerName.TabIndex = 9
        Me.TxtCustomerName.TabStop = False
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPrimaryCurrency.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(151, 273)
        Me.txtPrimaryCurrency.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(111, 22)
        Me.txtPrimaryCurrency.TabIndex = 23
        Me.txtPrimaryCurrency.TabStop = False
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPrimaryCurrency.Visible = False
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(13, 278)
        Me.Label22.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(102, 17)
        Me.Label22.TabIndex = 125
        Me.Label22.Text = "Currency Code"
        '
        'dtpDate
        '
        Me.dtpDate.CustomFormat = "ddd, dd MMM yyyy"
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.Location = New System.Drawing.Point(140, 47)
        Me.dtpDate.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(389, 22)
        Me.dtpDate.TabIndex = 3
        '
        'TxtPrimaryInvoice
        '
        Me.TxtPrimaryInvoice.Location = New System.Drawing.Point(140, 15)
        Me.TxtPrimaryInvoice.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryInvoice.Name = "TxtPrimaryInvoice"
        Me.TxtPrimaryInvoice.Size = New System.Drawing.Size(32, 22)
        Me.TxtPrimaryInvoice.TabIndex = 1
        Me.TxtPrimaryInvoice.TabStop = False
        Me.TxtPrimaryInvoice.Visible = False
        '
        'txtInvoiceNo
        '
        Me.txtInvoiceNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtInvoiceNo.Location = New System.Drawing.Point(140, 15)
        Me.txtInvoiceNo.Margin = New System.Windows.Forms.Padding(4)
        Me.txtInvoiceNo.MaxLength = 50
        Me.txtInvoiceNo.Name = "txtInvoiceNo"
        Me.txtInvoiceNo.ReadOnly = True
        Me.txtInvoiceNo.Size = New System.Drawing.Size(348, 22)
        Me.txtInvoiceNo.TabIndex = 0
        Me.txtInvoiceNo.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 52)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Date"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 20)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(74, 17)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Invoice No"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(615, 361)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(139, 32)
        Me.btnCancel.TabIndex = 31
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(460, 361)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(139, 32)
        Me.btnOK.TabIndex = 30
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox1.Location = New System.Drawing.Point(389, 271)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.TextBox1.MaxLength = 18
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(121, 22)
        Me.TextBox1.TabIndex = 25
        Me.TextBox1.TabStop = False
        Me.TextBox1.Text = "1.0000"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox1.Visible = False
        '
        'txtNoFakturPajakSederhana
        '
        Me.txtNoFakturPajakSederhana.Location = New System.Drawing.Point(140, 240)
        Me.txtNoFakturPajakSederhana.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNoFakturPajakSederhana.MaxLength = 50
        Me.txtNoFakturPajakSederhana.Name = "txtNoFakturPajakSederhana"
        Me.txtNoFakturPajakSederhana.Size = New System.Drawing.Size(389, 22)
        Me.txtNoFakturPajakSederhana.TabIndex = 12
        '
        'cboPPNDipungut
        '
        Me.cboPPNDipungut.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPPNDipungut.FormattingEnabled = True
        Me.cboPPNDipungut.Items.AddRange(New Object() {"Collect Tax", "Don't Collect Tax"})
        Me.cboPPNDipungut.Location = New System.Drawing.Point(767, 144)
        Me.cboPPNDipungut.Margin = New System.Windows.Forms.Padding(4)
        Me.cboPPNDipungut.Name = "cboPPNDipungut"
        Me.cboPPNDipungut.Size = New System.Drawing.Size(135, 24)
        Me.cboPPNDipungut.TabIndex = 23
        '
        'txtNoSeriFakturPajak
        '
        Me.txtNoSeriFakturPajak.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtNoSeriFakturPajak.Location = New System.Drawing.Point(140, 240)
        Me.txtNoSeriFakturPajak.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNoSeriFakturPajak.MaxLength = 50
        Me.txtNoSeriFakturPajak.Name = "txtNoSeriFakturPajak"
        Me.txtNoSeriFakturPajak.Size = New System.Drawing.Size(155, 22)
        Me.txtNoSeriFakturPajak.TabIndex = 18
        '
        'llbEditNoFakturPajak
        '
        Me.llbEditNoFakturPajak.AutoSize = True
        Me.llbEditNoFakturPajak.Location = New System.Drawing.Point(497, 245)
        Me.llbEditNoFakturPajak.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.llbEditNoFakturPajak.Name = "llbEditNoFakturPajak"
        Me.llbEditNoFakturPajak.Size = New System.Drawing.Size(32, 17)
        Me.llbEditNoFakturPajak.TabIndex = 21
        Me.llbEditNoFakturPajak.TabStop = True
        Me.llbEditNoFakturPajak.Text = "Edit"
        '
        'btnPrintTaxInvoice
        '
        Me.btnPrintTaxInvoice.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnPrintTaxInvoice.Location = New System.Drawing.Point(7, 407)
        Me.btnPrintTaxInvoice.Margin = New System.Windows.Forms.Padding(4)
        Me.btnPrintTaxInvoice.Name = "btnPrintTaxInvoice"
        Me.btnPrintTaxInvoice.Size = New System.Drawing.Size(107, 32)
        Me.btnPrintTaxInvoice.TabIndex = 4
        Me.btnPrintTaxInvoice.TabStop = False
        Me.btnPrintTaxInvoice.Text = "Tax Invoice"
        Me.btnPrintTaxInvoice.UseVisualStyleBackColor = True
        Me.btnPrintTaxInvoice.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'gbo_EFaktur
        '
        Me.gbo_EFaktur.Controls.Add(Me.Label41)
        Me.gbo_EFaktur.Controls.Add(Me.Txt_E_Faktur_Jalan)
        Me.gbo_EFaktur.Controls.Add(Me.TableLayoutPanel1)
        Me.gbo_EFaktur.Enabled = False
        Me.gbo_EFaktur.Location = New System.Drawing.Point(1007, 359)
        Me.gbo_EFaktur.Margin = New System.Windows.Forms.Padding(4)
        Me.gbo_EFaktur.Name = "gbo_EFaktur"
        Me.gbo_EFaktur.Padding = New System.Windows.Forms.Padding(4)
        Me.gbo_EFaktur.Size = New System.Drawing.Size(539, 203)
        Me.gbo_EFaktur.TabIndex = 61
        Me.gbo_EFaktur.TabStop = False
        Me.gbo_EFaktur.Text = "e-Faktur Data"
        Me.gbo_EFaktur.Visible = False
        '
        'Label41
        '
        Me.Label41.AutoSize = True
        Me.Label41.Location = New System.Drawing.Point(11, 23)
        Me.Label41.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(42, 17)
        Me.Label41.TabIndex = 78
        Me.Label41.Text = "Jalan"
        '
        'Txt_E_Faktur_Jalan
        '
        Me.Txt_E_Faktur_Jalan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Jalan.Location = New System.Drawing.Point(111, 20)
        Me.Txt_E_Faktur_Jalan.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_Jalan.MaxLength = 250
        Me.Txt_E_Faktur_Jalan.Name = "Txt_E_Faktur_Jalan"
        Me.Txt_E_Faktur_Jalan.Size = New System.Drawing.Size(424, 22)
        Me.Txt_E_Faktur_Jalan.TabIndex = 0
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 4
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 171.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 93.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 165.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label42, 2, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label43, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_NoTelepon, 3, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_KodePos, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_Propinsi, 4, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label44, 2, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_Kabupaten, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label45, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_Kecamatan, 3, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label46, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_Kelurahan, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label47, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_RW, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_RT, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label48, 2, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label49, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label50, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_No, 3, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Txt_E_Faktur_Blok, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label51, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(7, 50)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(529, 150)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'Label42
        '
        Me.Label42.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(275, 126)
        Me.Label42.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(82, 17)
        Me.Label42.TabIndex = 94
        Me.Label42.Text = "No Telepon"
        '
        'Label43
        '
        Me.Label43.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label43.AutoSize = True
        Me.Label43.Location = New System.Drawing.Point(4, 126)
        Me.Label43.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(69, 17)
        Me.Label43.TabIndex = 93
        Me.Label43.Text = "Kode Pos"
        '
        'Txt_E_Faktur_NoTelepon
        '
        Me.Txt_E_Faktur_NoTelepon.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_NoTelepon.Location = New System.Drawing.Point(368, 124)
        Me.Txt_E_Faktur_NoTelepon.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_NoTelepon.MaxLength = 250
        Me.Txt_E_Faktur_NoTelepon.Name = "Txt_E_Faktur_NoTelepon"
        Me.Txt_E_Faktur_NoTelepon.Size = New System.Drawing.Size(157, 22)
        Me.Txt_E_Faktur_NoTelepon.TabIndex = 9
        '
        'Txt_E_Faktur_KodePos
        '
        Me.Txt_E_Faktur_KodePos.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_KodePos.Location = New System.Drawing.Point(104, 124)
        Me.Txt_E_Faktur_KodePos.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_KodePos.MaxLength = 250
        Me.Txt_E_Faktur_KodePos.Name = "Txt_E_Faktur_KodePos"
        Me.Txt_E_Faktur_KodePos.Size = New System.Drawing.Size(163, 22)
        Me.Txt_E_Faktur_KodePos.TabIndex = 8
        '
        'Txt_E_Faktur_Propinsi
        '
        Me.Txt_E_Faktur_Propinsi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Propinsi.Location = New System.Drawing.Point(368, 94)
        Me.Txt_E_Faktur_Propinsi.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_Propinsi.MaxLength = 250
        Me.Txt_E_Faktur_Propinsi.Name = "Txt_E_Faktur_Propinsi"
        Me.Txt_E_Faktur_Propinsi.Size = New System.Drawing.Size(157, 22)
        Me.Txt_E_Faktur_Propinsi.TabIndex = 7
        '
        'Label44
        '
        Me.Label44.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label44.AutoSize = True
        Me.Label44.Location = New System.Drawing.Point(275, 96)
        Me.Label44.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(59, 17)
        Me.Label44.TabIndex = 89
        Me.Label44.Text = "Propinsi"
        '
        'Txt_E_Faktur_Kabupaten
        '
        Me.Txt_E_Faktur_Kabupaten.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Kabupaten.Location = New System.Drawing.Point(104, 94)
        Me.Txt_E_Faktur_Kabupaten.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_Kabupaten.MaxLength = 250
        Me.Txt_E_Faktur_Kabupaten.Name = "Txt_E_Faktur_Kabupaten"
        Me.Txt_E_Faktur_Kabupaten.Size = New System.Drawing.Size(163, 22)
        Me.Txt_E_Faktur_Kabupaten.TabIndex = 6
        '
        'Label45
        '
        Me.Label45.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label45.AutoSize = True
        Me.Label45.Location = New System.Drawing.Point(4, 96)
        Me.Label45.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(77, 17)
        Me.Label45.TabIndex = 87
        Me.Label45.Text = "Kabupaten"
        '
        'Txt_E_Faktur_Kecamatan
        '
        Me.Txt_E_Faktur_Kecamatan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Kecamatan.Location = New System.Drawing.Point(368, 64)
        Me.Txt_E_Faktur_Kecamatan.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_Kecamatan.MaxLength = 250
        Me.Txt_E_Faktur_Kecamatan.Name = "Txt_E_Faktur_Kecamatan"
        Me.Txt_E_Faktur_Kecamatan.Size = New System.Drawing.Size(157, 22)
        Me.Txt_E_Faktur_Kecamatan.TabIndex = 5
        '
        'Label46
        '
        Me.Label46.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label46.AutoSize = True
        Me.Label46.Location = New System.Drawing.Point(275, 66)
        Me.Label46.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(79, 17)
        Me.Label46.TabIndex = 85
        Me.Label46.Text = "Kecamatan"
        '
        'Txt_E_Faktur_Kelurahan
        '
        Me.Txt_E_Faktur_Kelurahan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Kelurahan.Location = New System.Drawing.Point(104, 64)
        Me.Txt_E_Faktur_Kelurahan.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_Kelurahan.MaxLength = 250
        Me.Txt_E_Faktur_Kelurahan.Name = "Txt_E_Faktur_Kelurahan"
        Me.Txt_E_Faktur_Kelurahan.Size = New System.Drawing.Size(163, 22)
        Me.Txt_E_Faktur_Kelurahan.TabIndex = 4
        '
        'Label47
        '
        Me.Label47.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label47.AutoSize = True
        Me.Label47.Location = New System.Drawing.Point(4, 66)
        Me.Label47.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(73, 17)
        Me.Label47.TabIndex = 83
        Me.Label47.Text = "Kelurahan"
        '
        'Txt_E_Faktur_RW
        '
        Me.Txt_E_Faktur_RW.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_RW.Location = New System.Drawing.Point(368, 34)
        Me.Txt_E_Faktur_RW.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_RW.MaxLength = 250
        Me.Txt_E_Faktur_RW.Name = "Txt_E_Faktur_RW"
        Me.Txt_E_Faktur_RW.Size = New System.Drawing.Size(157, 22)
        Me.Txt_E_Faktur_RW.TabIndex = 3
        '
        'Txt_E_Faktur_RT
        '
        Me.Txt_E_Faktur_RT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_RT.Location = New System.Drawing.Point(104, 34)
        Me.Txt_E_Faktur_RT.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_RT.MaxLength = 250
        Me.Txt_E_Faktur_RT.Name = "Txt_E_Faktur_RT"
        Me.Txt_E_Faktur_RT.Size = New System.Drawing.Size(163, 22)
        Me.Txt_E_Faktur_RT.TabIndex = 2
        '
        'Label48
        '
        Me.Label48.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label48.AutoSize = True
        Me.Label48.Location = New System.Drawing.Point(275, 36)
        Me.Label48.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(31, 17)
        Me.Label48.TabIndex = 80
        Me.Label48.Text = "RW"
        '
        'Label49
        '
        Me.Label49.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label49.AutoSize = True
        Me.Label49.Location = New System.Drawing.Point(4, 36)
        Me.Label49.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(27, 17)
        Me.Label49.TabIndex = 79
        Me.Label49.Text = "RT"
        '
        'Label50
        '
        Me.Label50.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label50.AutoSize = True
        Me.Label50.Location = New System.Drawing.Point(275, 6)
        Me.Label50.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(26, 17)
        Me.Label50.TabIndex = 78
        Me.Label50.Text = "No"
        '
        'Txt_E_Faktur_No
        '
        Me.Txt_E_Faktur_No.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_No.Location = New System.Drawing.Point(368, 4)
        Me.Txt_E_Faktur_No.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_No.MaxLength = 250
        Me.Txt_E_Faktur_No.Name = "Txt_E_Faktur_No"
        Me.Txt_E_Faktur_No.Size = New System.Drawing.Size(157, 22)
        Me.Txt_E_Faktur_No.TabIndex = 1
        '
        'Txt_E_Faktur_Blok
        '
        Me.Txt_E_Faktur_Blok.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Txt_E_Faktur_Blok.Location = New System.Drawing.Point(104, 4)
        Me.Txt_E_Faktur_Blok.Margin = New System.Windows.Forms.Padding(4)
        Me.Txt_E_Faktur_Blok.MaxLength = 250
        Me.Txt_E_Faktur_Blok.Name = "Txt_E_Faktur_Blok"
        Me.Txt_E_Faktur_Blok.Size = New System.Drawing.Size(163, 22)
        Me.Txt_E_Faktur_Blok.TabIndex = 0
        '
        'Label51
        '
        Me.Label51.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label51.AutoSize = True
        Me.Label51.Location = New System.Drawing.Point(4, 6)
        Me.Label51.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(35, 17)
        Me.Label51.TabIndex = 75
        Me.Label51.Text = "Blok"
        '
        'FrmTransaksiSalesInvoiceDP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(1259, 494)
        Me.Controls.Add(Me.gbo_EFaktur)
        Me.Controls.Add(Me.btnPrintTaxInvoice)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.gboDetail)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(1259, 505)
        Me.Name = "FrmTransaksiSalesInvoiceDP"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sales Invoice - Down Payment"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.TLayoutPanel2.ResumeLayout(False)
        Me.TLayoutPanel2.PerformLayout()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        Me.SplitContainer3.Panel2.PerformLayout()
        Me.SplitContainer3.ResumeLayout(False)
        Me.gbo_EFaktur.ResumeLayout(False)
        Me.gbo_EFaktur.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents txtTax As System.Windows.Forms.TextBox
    Friend WithEvents TxtTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents txtTaxExchangeRate As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents txtExhangeRateToIDR As System.Windows.Forms.TextBox
    Friend WithEvents txtExchangeRateToUSD As System.Windows.Forms.TextBox
    Friend WithEvents txtTaxAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents chkTax As System.Windows.Forms.CheckBox
    Friend WithEvents txtSubTotal As System.Windows.Forms.TextBox
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents txtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents TxtPrimaryInvoice As System.Windows.Forms.TextBox
    Friend WithEvents txtInvoiceNo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnFindSalesOrder As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtPrimarySO As System.Windows.Forms.TextBox
    Friend WithEvents txtSONo As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtDutyStamp As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtNoNPWP As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerAddIdentifier As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents llbAddress As System.Windows.Forms.LinkLabel
    Friend WithEvents cboCustomerAddress As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtNoFakturPajak As System.Windows.Forms.TextBox
    Friend WithEvents txtNoSeriFakturPajak As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents CboFakturPajak As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents txtNoFakturPajakSederhana As System.Windows.Forms.TextBox
    Friend WithEvents llbEditNoFakturPajak As System.Windows.Forms.LinkLabel
    Friend WithEvents llbEditNoInvoice As System.Windows.Forms.LinkLabel
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents btnPrintTaxInvoice As System.Windows.Forms.Button
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TxtNamaSigned As System.Windows.Forms.TextBox
    Friend WithEvents TxtJabatanSigned As System.Windows.Forms.TextBox
    Friend WithEvents llbUpDateNPWP As System.Windows.Forms.LinkLabel
    Friend WithEvents TxtPaymentTerm As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents cboPPNDipungut As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimarySalesPerson As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalesPersonCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtSalesPersonName As System.Windows.Forms.TextBox
    Friend WithEvents gbo_EFaktur As System.Windows.Forms.GroupBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Jalan As System.Windows.Forms.TextBox
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_NoTelepon As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_KodePos As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_Propinsi As System.Windows.Forms.TextBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Kabupaten As System.Windows.Forms.TextBox
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Kecamatan As System.Windows.Forms.TextBox
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_Kelurahan As System.Windows.Forms.TextBox
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_RW As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_RT As System.Windows.Forms.TextBox
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_No As System.Windows.Forms.TextBox
    Friend WithEvents Txt_E_Faktur_Blok As System.Windows.Forms.TextBox
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents Txt_E_Faktur_KodeJenisTransaksiPajak As System.Windows.Forms.TextBox
    Friend WithEvents txtPrimaryAccount As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents cboAccount As System.Windows.Forms.ComboBox
    Friend WithEvents cboPrimaryAccount As System.Windows.Forms.ComboBox
    Friend WithEvents txtPrimaryDPAccount As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cboDPAccount As System.Windows.Forms.ComboBox
    Friend WithEvents cboPrimaryDPAccount As System.Windows.Forms.ComboBox
    Friend WithEvents lblCurrTax As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCurrTax As System.Windows.Forms.TextBox
    Friend WithEvents cboCurrTax As System.Windows.Forms.ComboBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents cboPrimaryCurrTax As System.Windows.Forms.ComboBox
    Friend WithEvents LblPrintStatus As System.Windows.Forms.Label
    Friend WithEvents txtTaxAmount1 As System.Windows.Forms.TextBox
    Friend WithEvents cboNIKPASS As System.Windows.Forms.ComboBox
    Friend WithEvents txtNoNIKPass As System.Windows.Forms.TextBox
    Friend WithEvents chkRevisedInvoice As System.Windows.Forms.CheckBox
    Friend WithEvents LlbViewDetailTax As System.Windows.Forms.LinkLabel
    Friend WithEvents cboTaxtype As System.Windows.Forms.ComboBox
    Friend WithEvents txtDPPNilaiLain As System.Windows.Forms.TextBox
    Friend WithEvents TLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Public WithEvents txtRevisedInvoice As System.Windows.Forms.TextBox
End Class
