<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanVAT2_Custom
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
        Me.btnImportCheckData = New System.Windows.Forms.Button
        Me.ColTipeTransaksiString = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSuppCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWPNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNik = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPassport = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.ColNpwpName = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.dtpTo.Location = New System.Drawing.Point(293, 48)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(191, 22)
        Me.dtpTo.TabIndex = 2
        '
        'btnImportData
        '
        Me.btnImportData.Location = New System.Drawing.Point(757, 46)
        Me.btnImportData.Margin = New System.Windows.Forms.Padding(4)
        Me.btnImportData.Name = "btnImportData"
        Me.btnImportData.Size = New System.Drawing.Size(145, 28)
        Me.btnImportData.TabIndex = 38
        Me.btnImportData.TabStop = False
        Me.btnImportData.Text = "Import Data"
        Me.btnImportData.UseVisualStyleBackColor = True
        Me.btnImportData.Visible = False
        '
        'btnCreateDataImpor
        '
        Me.btnCreateDataImpor.Location = New System.Drawing.Point(603, 46)
        Me.btnCreateDataImpor.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCreateDataImpor.Name = "btnCreateDataImpor"
        Me.btnCreateDataImpor.Size = New System.Drawing.Size(145, 28)
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
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(4)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(4)
        Me.TabPage2.Size = New System.Drawing.Size(933, 445)
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
        Me.crv1.Location = New System.Drawing.Point(4, 4)
        Me.crv1.Margin = New System.Windows.Forms.Padding(4)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(925, 437)
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
        Me.TabControl1.Location = New System.Drawing.Point(21, 80)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(941, 474)
        Me.TabControl1.TabIndex = 6
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvItem)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(4)
        Me.TabPage1.Size = New System.Drawing.Size(933, 445)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Data"
        Me.TabPage1.UseVisualStyleBackColor = True
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTipeTransaksiString, Me.ColSuppCust, Me.ColNPWPNumber, Me.ColNik, Me.ColPassport, Me.ColInvoiceNumber, Me.ColTaxInvoiceNumber, Me.ColNoNotaRetur, Me.ColDate, Me.ColPurchaseSales, Me.ColVAT, Me.ColPrimaryDocAcuan, Me.ColEnumType, Me.ColEnumInOut, Me.ColTaxAddress, Me.ColTanggalTransaksiRef, Me.ColNoFakturPajakRef, Me.ColNamaJalan, Me.ColBlok, Me.ColNomor, Me.ColRT, Me.ColRW, Me.ColKelurahan, Me.ColKecamatan, Me.ColKabupaten, Me.ColPropinsi, Me.ColTelepon, Me.ColKodePos, Me.ColAlamatPajak1, Me.ColAlamatPajak2, Me.ColKodeJenisTransaksi, Me.ColIsDP, Me.ColNpwpName})
        Me.dgvItem.Location = New System.Drawing.Point(4, 4)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(923, 434)
        Me.dgvItem.TabIndex = 0
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(603, 16)
        Me.btnCreate.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(145, 28)
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
        Me.ComboBox1.Items.AddRange(New Object() {"VAT In", "VAT OUT", "Both", "VAT In: Purchase", "VAT In: Sales Return", "VAT Out: Sales", "VAT Out: Purchase Return"})
        Me.ComboBox1.Location = New System.Drawing.Point(21, 18)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(215, 24)
        Me.ComboBox1.TabIndex = 0
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(293, 16)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(4)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(191, 22)
        Me.dtpFrom.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(245, 50)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(25, 17)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "To"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(245, 18)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 17)
        Me.Label4.TabIndex = 34
        Me.Label4.Text = "From"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(493, 46)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(101, 28)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnShow
        '
        Me.btnShow.Location = New System.Drawing.Point(493, 16)
        Me.btnShow.Margin = New System.Windows.Forms.Padding(4)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(101, 28)
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
        Me.btnCreateCsv2015.Location = New System.Drawing.Point(603, 16)
        Me.btnCreateCsv2015.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCreateCsv2015.Name = "btnCreateCsv2015"
        Me.btnCreateCsv2015.Size = New System.Drawing.Size(159, 28)
        Me.btnCreateCsv2015.TabIndex = 4
        Me.btnCreateCsv2015.Text = "Create e-Faktur data"
        Me.btnCreateCsv2015.UseVisualStyleBackColor = True
        '
        'ChkShowDiscount
        '
        Me.ChkShowDiscount.AutoSize = True
        Me.ChkShowDiscount.Location = New System.Drawing.Point(21, 52)
        Me.ChkShowDiscount.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkShowDiscount.Name = "ChkShowDiscount"
        Me.ChkShowDiscount.Size = New System.Drawing.Size(120, 21)
        Me.ChkShowDiscount.TabIndex = 39
        Me.ChkShowDiscount.Text = "Show Discount"
        Me.ChkShowDiscount.UseVisualStyleBackColor = True
        '
        'btnImportCheckData
        '
        Me.btnImportCheckData.Location = New System.Drawing.Point(603, 46)
        Me.btnImportCheckData.Margin = New System.Windows.Forms.Padding(4)
        Me.btnImportCheckData.Name = "btnImportCheckData"
        Me.btnImportCheckData.Size = New System.Drawing.Size(159, 28)
        Me.btnImportCheckData.TabIndex = 40
        Me.btnImportCheckData.Text = "Import Check data"
        Me.btnImportCheckData.UseVisualStyleBackColor = True
        '
        'ColTipeTransaksiString
        '
        Me.ColTipeTransaksiString.HeaderText = "Type"
        Me.ColTipeTransaksiString.Name = "ColTipeTransaksiString"
        Me.ColTipeTransaksiString.ReadOnly = True
        Me.ColTipeTransaksiString.Width = 65
        '
        'ColSuppCust
        '
        Me.ColSuppCust.HeaderText = "Supplier/Customer"
        Me.ColSuppCust.Name = "ColSuppCust"
        Me.ColSuppCust.ReadOnly = True
        Me.ColSuppCust.Width = 149
        '
        'ColNPWPNumber
        '
        Me.ColNPWPNumber.HeaderText = "NPWP"
        Me.ColNPWPNumber.Name = "ColNPWPNumber"
        Me.ColNPWPNumber.ReadOnly = True
        Me.ColNPWPNumber.Width = 74
        '
        'ColNik
        '
        Me.ColNik.HeaderText = "NIK"
        Me.ColNik.Name = "ColNik"
        Me.ColNik.ReadOnly = True
        Me.ColNik.Width = 55
        '
        'ColPassport
        '
        Me.ColPassport.HeaderText = "PASSPORT"
        Me.ColPassport.Name = "ColPassport"
        Me.ColPassport.ReadOnly = True
        Me.ColPassport.Width = 108
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
        Me.ColInvoiceNumber.Width = 131
        '
        'ColTaxInvoiceNumber
        '
        Me.ColTaxInvoiceNumber.HeaderText = "Tax Invoice Number"
        Me.ColTaxInvoiceNumber.Name = "ColTaxInvoiceNumber"
        Me.ColTaxInvoiceNumber.ReadOnly = True
        Me.ColTaxInvoiceNumber.Width = 158
        '
        'ColNoNotaRetur
        '
        Me.ColNoNotaRetur.HeaderText = "Return Notes Number"
        Me.ColNoNotaRetur.Name = "ColNoNotaRetur"
        Me.ColNoNotaRetur.ReadOnly = True
        Me.ColNoNotaRetur.Width = 171
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 63
        '
        'ColPurchaseSales
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPurchaseSales.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColPurchaseSales.HeaderText = "Purchase/Sales"
        Me.ColPurchaseSales.Name = "ColPurchaseSales"
        Me.ColPurchaseSales.ReadOnly = True
        Me.ColPurchaseSales.Width = 132
        '
        'ColVAT
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColVAT.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColVAT.HeaderText = "VAT"
        Me.ColVAT.Name = "ColVAT"
        Me.ColVAT.ReadOnly = True
        Me.ColVAT.Width = 60
        '
        'ColPrimaryDocAcuan
        '
        Me.ColPrimaryDocAcuan.HeaderText = "Primary Doc Acuan"
        Me.ColPrimaryDocAcuan.Name = "ColPrimaryDocAcuan"
        Me.ColPrimaryDocAcuan.ReadOnly = True
        Me.ColPrimaryDocAcuan.Visible = False
        Me.ColPrimaryDocAcuan.Width = 154
        '
        'ColEnumType
        '
        Me.ColEnumType.HeaderText = "Enum Type"
        Me.ColEnumType.Name = "ColEnumType"
        Me.ColEnumType.ReadOnly = True
        Me.ColEnumType.Visible = False
        Me.ColEnumType.Width = 105
        '
        'ColEnumInOut
        '
        Me.ColEnumInOut.HeaderText = "Enum In Out"
        Me.ColEnumInOut.Name = "ColEnumInOut"
        Me.ColEnumInOut.ReadOnly = True
        Me.ColEnumInOut.Visible = False
        Me.ColEnumInOut.Width = 111
        '
        'ColTaxAddress
        '
        Me.ColTaxAddress.HeaderText = "Tax Address"
        Me.ColTaxAddress.Name = "ColTaxAddress"
        Me.ColTaxAddress.ReadOnly = True
        Me.ColTaxAddress.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxAddress.Visible = False
        Me.ColTaxAddress.Width = 93
        '
        'ColTanggalTransaksiRef
        '
        Me.ColTanggalTransaksiRef.HeaderText = "Tanggal Transaksi Ref"
        Me.ColTanggalTransaksiRef.Name = "ColTanggalTransaksiRef"
        Me.ColTanggalTransaksiRef.ReadOnly = True
        Me.ColTanggalTransaksiRef.Visible = False
        Me.ColTanggalTransaksiRef.Width = 177
        '
        'ColNoFakturPajakRef
        '
        Me.ColNoFakturPajakRef.HeaderText = "No Faktur Pajak Ref"
        Me.ColNoFakturPajakRef.Name = "ColNoFakturPajakRef"
        Me.ColNoFakturPajakRef.ReadOnly = True
        Me.ColNoFakturPajakRef.Visible = False
        Me.ColNoFakturPajakRef.Width = 160
        '
        'ColNamaJalan
        '
        Me.ColNamaJalan.HeaderText = "Nama Jalan"
        Me.ColNamaJalan.Name = "ColNamaJalan"
        Me.ColNamaJalan.ReadOnly = True
        Me.ColNamaJalan.Visible = False
        Me.ColNamaJalan.Width = 108
        '
        'ColBlok
        '
        Me.ColBlok.HeaderText = "Blok"
        Me.ColBlok.Name = "ColBlok"
        Me.ColBlok.ReadOnly = True
        Me.ColBlok.Visible = False
        Me.ColBlok.Width = 60
        '
        'ColNomor
        '
        Me.ColNomor.HeaderText = "Nomor"
        Me.ColNomor.Name = "ColNomor"
        Me.ColNomor.ReadOnly = True
        Me.ColNomor.Visible = False
        Me.ColNomor.Width = 75
        '
        'ColRT
        '
        Me.ColRT.HeaderText = "RT"
        Me.ColRT.Name = "ColRT"
        Me.ColRT.ReadOnly = True
        Me.ColRT.Visible = False
        Me.ColRT.Width = 52
        '
        'ColRW
        '
        Me.ColRW.HeaderText = "RW"
        Me.ColRW.Name = "ColRW"
        Me.ColRW.ReadOnly = True
        Me.ColRW.Visible = False
        Me.ColRW.Width = 56
        '
        'ColKelurahan
        '
        Me.ColKelurahan.HeaderText = "Kelurahan"
        Me.ColKelurahan.Name = "ColKelurahan"
        Me.ColKelurahan.ReadOnly = True
        Me.ColKelurahan.Visible = False
        Me.ColKelurahan.Width = 98
        '
        'ColKecamatan
        '
        Me.ColKecamatan.HeaderText = "Kecamatan"
        Me.ColKecamatan.Name = "ColKecamatan"
        Me.ColKecamatan.ReadOnly = True
        Me.ColKecamatan.Visible = False
        Me.ColKecamatan.Width = 104
        '
        'ColKabupaten
        '
        Me.ColKabupaten.HeaderText = "Kabupaten"
        Me.ColKabupaten.Name = "ColKabupaten"
        Me.ColKabupaten.ReadOnly = True
        Me.ColKabupaten.Visible = False
        Me.ColKabupaten.Width = 102
        '
        'ColPropinsi
        '
        Me.ColPropinsi.HeaderText = "Propinsi"
        Me.ColPropinsi.Name = "ColPropinsi"
        Me.ColPropinsi.ReadOnly = True
        Me.ColPropinsi.Visible = False
        Me.ColPropinsi.Width = 84
        '
        'ColTelepon
        '
        Me.ColTelepon.HeaderText = "Telepon"
        Me.ColTelepon.Name = "ColTelepon"
        Me.ColTelepon.ReadOnly = True
        Me.ColTelepon.Visible = False
        Me.ColTelepon.Width = 85
        '
        'ColKodePos
        '
        Me.ColKodePos.HeaderText = "Kode Pos"
        Me.ColKodePos.Name = "ColKodePos"
        Me.ColKodePos.ReadOnly = True
        Me.ColKodePos.Visible = False
        Me.ColKodePos.Width = 94
        '
        'ColAlamatPajak1
        '
        Me.ColAlamatPajak1.HeaderText = "Alamat Pajak 1"
        Me.ColAlamatPajak1.Name = "ColAlamatPajak1"
        Me.ColAlamatPajak1.ReadOnly = True
        Me.ColAlamatPajak1.Visible = False
        Me.ColAlamatPajak1.Width = 127
        '
        'ColAlamatPajak2
        '
        Me.ColAlamatPajak2.HeaderText = "Alamat Pajak 2"
        Me.ColAlamatPajak2.Name = "ColAlamatPajak2"
        Me.ColAlamatPajak2.ReadOnly = True
        Me.ColAlamatPajak2.Visible = False
        Me.ColAlamatPajak2.Width = 127
        '
        'ColKodeJenisTransaksi
        '
        Me.ColKodeJenisTransaksi.HeaderText = "Kode Jenis Transaksi"
        Me.ColKodeJenisTransaksi.Name = "ColKodeJenisTransaksi"
        Me.ColKodeJenisTransaksi.ReadOnly = True
        Me.ColKodeJenisTransaksi.Visible = False
        Me.ColKodeJenisTransaksi.Width = 169
        '
        'ColIsDP
        '
        Me.ColIsDP.HeaderText = "Is DP"
        Me.ColIsDP.Name = "ColIsDP"
        Me.ColIsDP.ReadOnly = True
        Me.ColIsDP.Visible = False
        Me.ColIsDP.Width = 66
        '
        'ColNpwpName
        '
        Me.ColNpwpName.HeaderText = "NPWP Name"
        Me.ColNpwpName.Name = "ColNpwpName"
        Me.ColNpwpName.ReadOnly = True
        Me.ColNpwpName.Visible = False
        Me.ColNpwpName.Width = 115
        '
        'FrmLaporanVAT2_Custom
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(991, 570)
        Me.Controls.Add(Me.btnImportCheckData)
        Me.Controls.Add(Me.ChkShowDiscount)
        Me.Controls.Add(Me.btnCreateCsv2015)
        Me.Controls.Add(Me.dtpTo)
        Me.Controls.Add(Me.btnImportData)
        Me.Controls.Add(Me.btnCreateDataImpor)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.dtpFrom)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnShow)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimumSize = New System.Drawing.Size(999, 605)
        Me.Name = "FrmLaporanVAT2_Custom"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VAT Report - e-Faktur 2015"
        Me.TabPage2.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
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
    Friend WithEvents btnImportCheckData As System.Windows.Forms.Button
    Friend WithEvents ColTipeTransaksiString As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSuppCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWPNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNik As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPassport As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents ColNpwpName As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
