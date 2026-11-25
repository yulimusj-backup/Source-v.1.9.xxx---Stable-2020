<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanVAT2025
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
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.btnImportData = New System.Windows.Forms.Button
        Me.btnCreateDataImpor = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Llbcheck = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.btnCreate = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnShow = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.btnCreateCsv2015 = New System.Windows.Forms.Button
        Me.ChkShowDiscount = New System.Windows.Forms.CheckBox
        Me.BtnCreateCoreTax = New System.Windows.Forms.Button
        Me.btnImportTaxNumber = New System.Windows.Forms.Button
        Me.ColExport = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.colEmail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIDTKU = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExportStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeTransaksiString = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWPNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNik = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPassport = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOtherID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceNumber = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColTaxInvoiceNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNoNotaRetur = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPurchaseSales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVAT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDocAcuan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEnumType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEnumInOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxAddress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTanggalTransaksiRef = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNoFakturPajakRef = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNamaJalan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBlok = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNomor = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRW = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKelurahan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKecamatan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKabupaten = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPropinsi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTelepon = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKodePos = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlamatPajak1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAlamatPajak2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKodeJenisTransaksi = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsDP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSPPBNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colNPWPType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colAdditionalInfo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFacilityStamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colTaxRevised = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colTaxType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colSPPBDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtpTo
        '
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(220, 39)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(144, 20)
        Me.dtpTo.TabIndex = 2
        Me.dtpTo.Value = New Date(2024, 12, 31, 0, 0, 0, 0)
        '
        'btnImportData
        '
        Me.btnImportData.Location = New System.Drawing.Point(588, 37)
        Me.btnImportData.Name = "btnImportData"
        Me.btnImportData.Size = New System.Drawing.Size(130, 23)
        Me.btnImportData.TabIndex = 38
        Me.btnImportData.TabStop = False
        Me.btnImportData.Text = "Import Data"
        Me.btnImportData.UseVisualStyleBackColor = True
        Me.btnImportData.Visible = False
        '
        'btnCreateDataImpor
        '
        Me.btnCreateDataImpor.Location = New System.Drawing.Point(588, 39)
        Me.btnCreateDataImpor.Name = "btnCreateDataImpor"
        Me.btnCreateDataImpor.Size = New System.Drawing.Size(109, 23)
        Me.btnCreateDataImpor.TabIndex = 37
        Me.btnCreateDataImpor.TabStop = False
        Me.btnCreateDataImpor.Text = "Create Data Import"
        Me.btnCreateDataImpor.UseVisualStyleBackColor = True
        Me.btnCreateDataImpor.Visible = False
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "csv"
        Me.SaveFileDialog1.Filter = "csv files|*.csv"
        Me.SaveFileDialog1.Title = "Save e-SPT data"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.crv1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(713, 359)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Printable Version"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.DisplayStatusBar = False
        Me.crv1.DisplayToolbar = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.Location = New System.Drawing.Point(3, 3)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(707, 353)
        Me.crv1.TabIndex = 6
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(16, 65)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(721, 385)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Llbcheck)
        Me.TabPage1.Controls.Add(Me.LinkLabel1)
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(713, 359)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Data"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Llbcheck
        '
        Me.Llbcheck.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Llbcheck.AutoSize = True
        Me.Llbcheck.Location = New System.Drawing.Point(660, 3)
        Me.Llbcheck.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Llbcheck.Name = "Llbcheck"
        Me.Llbcheck.Size = New System.Drawing.Size(52, 13)
        Me.Llbcheck.TabIndex = 4
        Me.Llbcheck.TabStop = True
        Me.Llbcheck.Text = "Check All"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(594, 3)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel1.TabIndex = 3
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Uncheck All"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColExport, Me.colEmail, Me.ColIDTKU, Me.ColExportStatus, Me.ColTipeTransaksiString, Me.ColSuppCust, Me.ColNPWPNumber, Me.ColNik, Me.ColPassport, Me.ColOtherID, Me.ColInvoiceNumber, Me.ColTaxInvoiceNumber, Me.ColNoNotaRetur, Me.ColDate, Me.ColPurchaseSales, Me.ColVAT, Me.ColPrimaryDocAcuan, Me.ColEnumType, Me.ColEnumInOut, Me.ColTaxAddress, Me.ColTanggalTransaksiRef, Me.ColNoFakturPajakRef, Me.ColNamaJalan, Me.ColBlok, Me.ColNomor, Me.ColRT, Me.ColRW, Me.ColKelurahan, Me.ColKecamatan, Me.ColKabupaten, Me.ColPropinsi, Me.ColTelepon, Me.ColKodePos, Me.ColAlamatPajak1, Me.ColAlamatPajak2, Me.ColKodeJenisTransaksi, Me.ColIsDP, Me.ColSPPBNo, Me.colNPWPType, Me.colAdditionalInfo, Me.ColFacilityStamp, Me.colTaxRevised, Me.colTaxType, Me.colSPPBDate})
        Me.dgvItem.Location = New System.Drawing.Point(3, 20)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(707, 336)
        Me.dgvItem.TabIndex = 0
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(598, 41)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(109, 23)
        Me.btnCreate.TabIndex = 33
        Me.btnCreate.TabStop = False
        Me.btnCreate.Text = "Create csv data-obselete"
        Me.btnCreate.UseVisualStyleBackColor = True
        Me.btnCreate.Visible = False
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"VAT Out: Sales", "VAT Out: Sales DP"})
        Me.ComboBox1.Location = New System.Drawing.Point(16, 15)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(162, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.dtpFrom.Location = New System.Drawing.Point(220, 13)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(144, 20)
        Me.dtpFrom.TabIndex = 1
        Me.dtpFrom.Value = New Date(2024, 12, 31, 0, 0, 0, 0)
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(184, 41)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "To"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(184, 15)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(30, 13)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "From"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(370, 37)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(76, 23)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(370, 13)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(76, 23)
        Me.btnShow.TabIndex = 3
        Me.btnShow.Text = "Show"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "csv files|*.csv"
        Me.OpenFileDialog1.Title = "Import Tax Invoice Number"
        '
        'btnCreateCsv2015
        '
        Me.btnCreateCsv2015.Location = New System.Drawing.Point(588, 15)
        Me.btnCreateCsv2015.Name = "btnCreateCsv2015"
        Me.btnCreateCsv2015.Size = New System.Drawing.Size(119, 23)
        Me.btnCreateCsv2015.TabIndex = 4
        Me.btnCreateCsv2015.Text = "Create e-Faktur data"
        Me.btnCreateCsv2015.UseVisualStyleBackColor = True
        Me.btnCreateCsv2015.Visible = False
        '
        'ChkShowDiscount
        '
        Me.ChkShowDiscount.AutoSize = True
        Me.ChkShowDiscount.Location = New System.Drawing.Point(16, 42)
        Me.ChkShowDiscount.Name = "ChkShowDiscount"
        Me.ChkShowDiscount.Size = New System.Drawing.Size(98, 17)
        Me.ChkShowDiscount.TabIndex = 39
        Me.ChkShowDiscount.Text = "Show Discount"
        Me.ChkShowDiscount.UseVisualStyleBackColor = True
        '
        'BtnCreateCoreTax
        '
        Me.BtnCreateCoreTax.Location = New System.Drawing.Point(452, 13)
        Me.BtnCreateCoreTax.Name = "BtnCreateCoreTax"
        Me.BtnCreateCoreTax.Size = New System.Drawing.Size(115, 23)
        Me.BtnCreateCoreTax.TabIndex = 40
        Me.BtnCreateCoreTax.Text = "Data for Converter"
        Me.BtnCreateCoreTax.UseVisualStyleBackColor = True
        Me.BtnCreateCoreTax.Visible = False
        '
        'btnImportTaxNumber
        '
        Me.btnImportTaxNumber.Location = New System.Drawing.Point(452, 36)
        Me.btnImportTaxNumber.Name = "btnImportTaxNumber"
        Me.btnImportTaxNumber.Size = New System.Drawing.Size(115, 23)
        Me.btnImportTaxNumber.TabIndex = 41
        Me.btnImportTaxNumber.Text = "Import Tax Number"
        Me.btnImportTaxNumber.UseVisualStyleBackColor = True
        '
        'ColExport
        '
        Me.ColExport.HeaderText = "Export"
        Me.ColExport.Name = "ColExport"
        Me.ColExport.ReadOnly = True
        Me.ColExport.Width = 43
        '
        'colEmail
        '
        Me.colEmail.HeaderText = "Email"
        Me.colEmail.Name = "colEmail"
        Me.colEmail.ReadOnly = True
        Me.colEmail.Visible = False
        Me.colEmail.Width = 57
        '
        'ColIDTKU
        '
        Me.ColIDTKU.HeaderText = "ID TKU"
        Me.ColIDTKU.Name = "ColIDTKU"
        Me.ColIDTKU.ReadOnly = True
        Me.ColIDTKU.Visible = False
        Me.ColIDTKU.Width = 68
        '
        'ColExportStatus
        '
        Me.ColExportStatus.HeaderText = "Export Status"
        Me.ColExportStatus.Name = "ColExportStatus"
        Me.ColExportStatus.ReadOnly = True
        Me.ColExportStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColExportStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExportStatus.Width = 76
        '
        'ColTipeTransaksiString
        '
        Me.ColTipeTransaksiString.HeaderText = "Type"
        Me.ColTipeTransaksiString.Name = "ColTipeTransaksiString"
        Me.ColTipeTransaksiString.ReadOnly = True
        Me.ColTipeTransaksiString.Width = 56
        '
        'ColSuppCust
        '
        Me.ColSuppCust.HeaderText = "Supplier/Customer"
        Me.ColSuppCust.Name = "ColSuppCust"
        Me.ColSuppCust.ReadOnly = True
        Me.ColSuppCust.Width = 119
        '
        'ColNPWPNumber
        '
        Me.ColNPWPNumber.HeaderText = "NPWP"
        Me.ColNPWPNumber.Name = "ColNPWPNumber"
        Me.ColNPWPNumber.ReadOnly = True
        Me.ColNPWPNumber.Width = 65
        '
        'ColNik
        '
        Me.ColNik.HeaderText = "NIK"
        Me.ColNik.Name = "ColNik"
        Me.ColNik.ReadOnly = True
        Me.ColNik.Width = 50
        '
        'ColPassport
        '
        Me.ColPassport.HeaderText = "PASSPORT"
        Me.ColPassport.Name = "ColPassport"
        Me.ColPassport.ReadOnly = True
        Me.ColPassport.Width = 90
        '
        'ColOtherID
        '
        Me.ColOtherID.HeaderText = "OTHER ID"
        Me.ColOtherID.Name = "ColOtherID"
        Me.ColOtherID.ReadOnly = True
        Me.ColOtherID.Width = 84
        '
        'ColInvoiceNumber
        '
        Me.ColInvoiceNumber.ActiveLinkColor = System.Drawing.Color.Blue
        Me.ColInvoiceNumber.HeaderText = "Invoice Number"
        Me.ColInvoiceNumber.Name = "ColInvoiceNumber"
        Me.ColInvoiceNumber.ReadOnly = True
        Me.ColInvoiceNumber.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColInvoiceNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColInvoiceNumber.VisitedLinkColor = System.Drawing.Color.Blue
        Me.ColInvoiceNumber.Width = 107
        '
        'ColTaxInvoiceNumber
        '
        Me.ColTaxInvoiceNumber.HeaderText = "Tax Invoice Number"
        Me.ColTaxInvoiceNumber.Name = "ColTaxInvoiceNumber"
        Me.ColTaxInvoiceNumber.ReadOnly = True
        Me.ColTaxInvoiceNumber.Width = 128
        '
        'ColNoNotaRetur
        '
        Me.ColNoNotaRetur.HeaderText = "Return Notes Number"
        Me.ColNoNotaRetur.Name = "ColNoNotaRetur"
        Me.ColNoNotaRetur.ReadOnly = True
        Me.ColNoNotaRetur.Width = 135
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 55
        '
        'ColPurchaseSales
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPurchaseSales.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColPurchaseSales.HeaderText = "Purchase/Sales"
        Me.ColPurchaseSales.Name = "ColPurchaseSales"
        Me.ColPurchaseSales.ReadOnly = True
        Me.ColPurchaseSales.Width = 108
        '
        'ColVAT
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColVAT.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColVAT.HeaderText = "VAT"
        Me.ColVAT.Name = "ColVAT"
        Me.ColVAT.ReadOnly = True
        Me.ColVAT.Width = 53
        '
        'ColPrimaryDocAcuan
        '
        Me.ColPrimaryDocAcuan.HeaderText = "Primary Doc Acuan"
        Me.ColPrimaryDocAcuan.Name = "ColPrimaryDocAcuan"
        Me.ColPrimaryDocAcuan.ReadOnly = True
        Me.ColPrimaryDocAcuan.Visible = False
        Me.ColPrimaryDocAcuan.Width = 123
        '
        'ColEnumType
        '
        Me.ColEnumType.HeaderText = "Enum Type"
        Me.ColEnumType.Name = "ColEnumType"
        Me.ColEnumType.ReadOnly = True
        Me.ColEnumType.Visible = False
        Me.ColEnumType.Width = 86
        '
        'ColEnumInOut
        '
        Me.ColEnumInOut.HeaderText = "Enum In Out"
        Me.ColEnumInOut.Name = "ColEnumInOut"
        Me.ColEnumInOut.ReadOnly = True
        Me.ColEnumInOut.Visible = False
        Me.ColEnumInOut.Width = 91
        '
        'ColTaxAddress
        '
        Me.ColTaxAddress.HeaderText = "Tax Address"
        Me.ColTaxAddress.Name = "ColTaxAddress"
        Me.ColTaxAddress.ReadOnly = True
        Me.ColTaxAddress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxAddress.Visible = False
        Me.ColTaxAddress.Width = 72
        '
        'ColTanggalTransaksiRef
        '
        Me.ColTanggalTransaksiRef.HeaderText = "Tanggal Transaksi Ref"
        Me.ColTanggalTransaksiRef.Name = "ColTanggalTransaksiRef"
        Me.ColTanggalTransaksiRef.ReadOnly = True
        Me.ColTanggalTransaksiRef.Visible = False
        Me.ColTanggalTransaksiRef.Width = 140
        '
        'ColNoFakturPajakRef
        '
        Me.ColNoFakturPajakRef.HeaderText = "No Faktur Pajak Ref"
        Me.ColNoFakturPajakRef.Name = "ColNoFakturPajakRef"
        Me.ColNoFakturPajakRef.ReadOnly = True
        Me.ColNoFakturPajakRef.Visible = False
        Me.ColNoFakturPajakRef.Width = 129
        '
        'ColNamaJalan
        '
        Me.ColNamaJalan.HeaderText = "Nama Jalan"
        Me.ColNamaJalan.Name = "ColNamaJalan"
        Me.ColNamaJalan.ReadOnly = True
        Me.ColNamaJalan.Visible = False
        Me.ColNamaJalan.Width = 88
        '
        'ColBlok
        '
        Me.ColBlok.HeaderText = "Blok"
        Me.ColBlok.Name = "ColBlok"
        Me.ColBlok.ReadOnly = True
        Me.ColBlok.Visible = False
        Me.ColBlok.Width = 53
        '
        'ColNomor
        '
        Me.ColNomor.HeaderText = "Nomor"
        Me.ColNomor.Name = "ColNomor"
        Me.ColNomor.ReadOnly = True
        Me.ColNomor.Visible = False
        Me.ColNomor.Width = 63
        '
        'ColRT
        '
        Me.ColRT.HeaderText = "RT"
        Me.ColRT.Name = "ColRT"
        Me.ColRT.ReadOnly = True
        Me.ColRT.Visible = False
        Me.ColRT.Width = 47
        '
        'ColRW
        '
        Me.ColRW.HeaderText = "RW"
        Me.ColRW.Name = "ColRW"
        Me.ColRW.ReadOnly = True
        Me.ColRW.Visible = False
        Me.ColRW.Width = 51
        '
        'ColKelurahan
        '
        Me.ColKelurahan.HeaderText = "Kelurahan"
        Me.ColKelurahan.Name = "ColKelurahan"
        Me.ColKelurahan.ReadOnly = True
        Me.ColKelurahan.Visible = False
        Me.ColKelurahan.Width = 80
        '
        'ColKecamatan
        '
        Me.ColKecamatan.HeaderText = "Kecamatan"
        Me.ColKecamatan.Name = "ColKecamatan"
        Me.ColKecamatan.ReadOnly = True
        Me.ColKecamatan.Visible = False
        Me.ColKecamatan.Width = 86
        '
        'ColKabupaten
        '
        Me.ColKabupaten.HeaderText = "Kabupaten"
        Me.ColKabupaten.Name = "ColKabupaten"
        Me.ColKabupaten.ReadOnly = True
        Me.ColKabupaten.Visible = False
        Me.ColKabupaten.Width = 84
        '
        'ColPropinsi
        '
        Me.ColPropinsi.HeaderText = "Propinsi"
        Me.ColPropinsi.Name = "ColPropinsi"
        Me.ColPropinsi.ReadOnly = True
        Me.ColPropinsi.Visible = False
        Me.ColPropinsi.Width = 69
        '
        'ColTelepon
        '
        Me.ColTelepon.HeaderText = "Telepon"
        Me.ColTelepon.Name = "ColTelepon"
        Me.ColTelepon.ReadOnly = True
        Me.ColTelepon.Visible = False
        Me.ColTelepon.Width = 71
        '
        'ColKodePos
        '
        Me.ColKodePos.HeaderText = "Kode Pos"
        Me.ColKodePos.Name = "ColKodePos"
        Me.ColKodePos.ReadOnly = True
        Me.ColKodePos.Visible = False
        Me.ColKodePos.Width = 78
        '
        'ColAlamatPajak1
        '
        Me.ColAlamatPajak1.HeaderText = "Alamat Pajak 1"
        Me.ColAlamatPajak1.Name = "ColAlamatPajak1"
        Me.ColAlamatPajak1.ReadOnly = True
        Me.ColAlamatPajak1.Visible = False
        Me.ColAlamatPajak1.Width = 103
        '
        'ColAlamatPajak2
        '
        Me.ColAlamatPajak2.HeaderText = "Alamat Pajak 2"
        Me.ColAlamatPajak2.Name = "ColAlamatPajak2"
        Me.ColAlamatPajak2.ReadOnly = True
        Me.ColAlamatPajak2.Visible = False
        Me.ColAlamatPajak2.Width = 103
        '
        'ColKodeJenisTransaksi
        '
        Me.ColKodeJenisTransaksi.HeaderText = "Kode Jenis Transaksi"
        Me.ColKodeJenisTransaksi.Name = "ColKodeJenisTransaksi"
        Me.ColKodeJenisTransaksi.ReadOnly = True
        Me.ColKodeJenisTransaksi.Visible = False
        Me.ColKodeJenisTransaksi.Width = 133
        '
        'ColIsDP
        '
        Me.ColIsDP.HeaderText = "Is DP"
        Me.ColIsDP.Name = "ColIsDP"
        Me.ColIsDP.ReadOnly = True
        Me.ColIsDP.Visible = False
        Me.ColIsDP.Width = 58
        '
        'ColSPPBNo
        '
        Me.ColSPPBNo.HeaderText = "SPPB No"
        Me.ColSPPBNo.Name = "ColSPPBNo"
        Me.ColSPPBNo.ReadOnly = True
        Me.ColSPPBNo.Visible = False
        Me.ColSPPBNo.Width = 77
        '
        'colNPWPType
        '
        Me.colNPWPType.HeaderText = "NPWP Type"
        Me.colNPWPType.Name = "colNPWPType"
        Me.colNPWPType.ReadOnly = True
        Me.colNPWPType.Visible = False
        Me.colNPWPType.Width = 92
        '
        'colAdditionalInfo
        '
        Me.colAdditionalInfo.HeaderText = "Additional Info"
        Me.colAdditionalInfo.Name = "colAdditionalInfo"
        Me.colAdditionalInfo.ReadOnly = True
        Me.colAdditionalInfo.Visible = False
        Me.colAdditionalInfo.Width = 99
        '
        'ColFacilityStamp
        '
        Me.ColFacilityStamp.HeaderText = "Facility Stamp"
        Me.ColFacilityStamp.Name = "ColFacilityStamp"
        Me.ColFacilityStamp.ReadOnly = True
        Me.ColFacilityStamp.Visible = False
        Me.ColFacilityStamp.Width = 97
        '
        'colTaxRevised
        '
        Me.colTaxRevised.HeaderText = "Tax Revised"
        Me.colTaxRevised.Name = "colTaxRevised"
        Me.colTaxRevised.ReadOnly = True
        Me.colTaxRevised.Visible = False
        Me.colTaxRevised.Width = 92
        '
        'colTaxType
        '
        Me.colTaxType.HeaderText = "Tax Type"
        Me.colTaxType.Name = "colTaxType"
        Me.colTaxType.ReadOnly = True
        Me.colTaxType.Visible = False
        Me.colTaxType.Width = 77
        '
        'colSPPBDate
        '
        Me.colSPPBDate.HeaderText = "SPPB Date"
        Me.colSPPBDate.Name = "colSPPBDate"
        Me.colSPPBDate.ReadOnly = True
        Me.colSPPBDate.Visible = False
        Me.colSPPBDate.Width = 86
        '
        'FrmLaporanVAT2025
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(748, 474)
        Me.Controls.Add(Me.btnImportTaxNumber)
        Me.Controls.Add(Me.BtnCreateCoreTax)
        Me.Controls.Add(Me.ChkShowDiscount)
        Me.Controls.Add(Me.btnCreateCsv2015)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.btnImportData)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnShow)
        Me.Controls.Add(Me.btnCreateDataImpor)
        Me.Controls.Add(Me.btnCreate)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(756, 501)
        Me.Name = "FrmLaporanVAT2025"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VAT Report CoreTax 2025"
        Me.TabPage2.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnImportData As System.Windows.Forms.Button
    Friend WithEvents btnCreateDataImpor As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnCreateCsv2015 As System.Windows.Forms.Button
    Friend WithEvents ChkShowDiscount As System.Windows.Forms.CheckBox
    Friend WithEvents Llbcheck As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents BtnCreateCoreTax As System.Windows.Forms.Button
    Friend WithEvents btnImportTaxNumber As System.Windows.Forms.Button
    Friend WithEvents ColExport As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents colEmail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIDTKU As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExportStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeTransaksiString As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWPNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNik As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPassport As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOtherID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceNumber As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColTaxInvoiceNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoNotaRetur As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPurchaseSales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVAT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDocAcuan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEnumType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEnumInOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTanggalTransaksiRef As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoFakturPajakRef As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNamaJalan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBlok As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNomor As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRW As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKelurahan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKecamatan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKabupaten As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPropinsi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTelepon As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKodePos As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlamatPajak1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAlamatPajak2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKodeJenisTransaksi As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsDP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSPPBNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNPWPType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colAdditionalInfo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFacilityStamp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTaxRevised As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTaxType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colSPPBDate As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
