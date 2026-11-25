<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterSupplier
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
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtEmail = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtTKUID = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.LlbEditNPWPLama = New System.Windows.Forms.LinkLabel
        Me.CboTipeNPWP = New System.Windows.Forms.ComboBox
        Me.rboNo = New System.Windows.Forms.RadioButton
        Me.rboYes = New System.Windows.Forms.RadioButton
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvAddress = New System.Windows.Forms.DataGridView
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column9 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierAdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.dgvCP = New System.Windows.Forms.DataGridView
        Me.ColPersonCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPersonName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPhoneNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFax = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column18 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.dgvBank = New System.Windows.Forms.DataGridView
        Me.Column12 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNamaBank = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccountNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAddress1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBeneficiaryName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column17 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column19 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSwiftCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBankBranch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TxtCompanyWebsite = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnFindCurrency = New System.Windows.Forms.Button
        Me.TxtCurrencyCode = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtNPWP = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtLeadTime = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.TxtCountry = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.TxtPaymentTerm = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.CboPriceType = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.TxtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.txtsuppliername = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtsuppliercode = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtprimarysupplier = New System.Windows.Forms.TextBox
        Me.TxtNPWPLama = New System.Windows.Forms.TextBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.gboDetail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvCP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.dgvBank, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TxtEmail)
        Me.gboDetail.Controls.Add(Me.Label6)
        Me.gboDetail.Controls.Add(Me.TxtTKUID)
        Me.gboDetail.Controls.Add(Me.Label5)
        Me.gboDetail.Controls.Add(Me.LlbEditNPWPLama)
        Me.gboDetail.Controls.Add(Me.CboTipeNPWP)
        Me.gboDetail.Controls.Add(Me.rboNo)
        Me.gboDetail.Controls.Add(Me.rboYes)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.TxtCompanyWebsite)
        Me.gboDetail.Controls.Add(Me.Label4)
        Me.gboDetail.Controls.Add(Me.btnFindCurrency)
        Me.gboDetail.Controls.Add(Me.TxtCurrencyCode)
        Me.gboDetail.Controls.Add(Me.Label3)
        Me.gboDetail.Controls.Add(Me.TxtNPWP)
        Me.gboDetail.Controls.Add(Me.Label16)
        Me.gboDetail.Controls.Add(Me.txtLeadTime)
        Me.gboDetail.Controls.Add(Me.Label15)
        Me.gboDetail.Controls.Add(Me.Label14)
        Me.gboDetail.Controls.Add(Me.TxtCountry)
        Me.gboDetail.Controls.Add(Me.Label13)
        Me.gboDetail.Controls.Add(Me.TxtPaymentTerm)
        Me.gboDetail.Controls.Add(Me.Label12)
        Me.gboDetail.Controls.Add(Me.CboPriceType)
        Me.gboDetail.Controls.Add(Me.Label11)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryCurrency)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.txtsuppliername)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.txtsuppliercode)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.txtprimarysupplier)
        Me.gboDetail.Controls.Add(Me.TxtNPWPLama)
        Me.gboDetail.Location = New System.Drawing.Point(7, 4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(872, 483)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'TxtEmail
        '
        Me.TxtEmail.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtEmail.Location = New System.Drawing.Point(560, 152)
        Me.TxtEmail.MaxLength = 100
        Me.TxtEmail.Name = "TxtEmail"
        Me.TxtEmail.Size = New System.Drawing.Size(303, 20)
        Me.TxtEmail.TabIndex = 14
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(460, 154)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 13)
        Me.Label6.TabIndex = 50
        Me.Label6.Text = "Email"
        '
        'TxtTKUID
        '
        Me.TxtTKUID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTKUID.Location = New System.Drawing.Point(560, 129)
        Me.TxtTKUID.MaxLength = 6
        Me.TxtTKUID.Name = "TxtTKUID"
        Me.TxtTKUID.Size = New System.Drawing.Size(303, 20)
        Me.TxtTKUID.TabIndex = 13
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(460, 132)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(43, 13)
        Me.Label5.TabIndex = 49
        Me.Label5.Text = "TKU ID"
        '
        'LlbEditNPWPLama
        '
        Me.LlbEditNPWPLama.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbEditNPWPLama.AutoSize = True
        Me.LlbEditNPWPLama.Location = New System.Drawing.Point(782, 107)
        Me.LlbEditNPWPLama.Name = "LlbEditNPWPLama"
        Me.LlbEditNPWPLama.Size = New System.Drawing.Size(80, 13)
        Me.LlbEditNPWPLama.TabIndex = 12
        Me.LlbEditNPWPLama.TabStop = True
        Me.LlbEditNPWPLama.Text = "Edit Old NPWP"
        '
        'CboTipeNPWP
        '
        Me.CboTipeNPWP.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboTipeNPWP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTipeNPWP.FormattingEnabled = True
        Me.CboTipeNPWP.Items.AddRange(New Object() {"Company", "Individual"})
        Me.CboTipeNPWP.Location = New System.Drawing.Point(699, 104)
        Me.CboTipeNPWP.Name = "CboTipeNPWP"
        Me.CboTipeNPWP.Size = New System.Drawing.Size(77, 21)
        Me.CboTipeNPWP.TabIndex = 11
        '
        'rboNo
        '
        Me.rboNo.AutoSize = True
        Me.rboNo.Checked = True
        Me.rboNo.Location = New System.Drawing.Point(673, 78)
        Me.rboNo.Name = "rboNo"
        Me.rboNo.Size = New System.Drawing.Size(39, 17)
        Me.rboNo.TabIndex = 9
        Me.rboNo.TabStop = True
        Me.rboNo.Text = "No"
        Me.rboNo.UseVisualStyleBackColor = True
        '
        'rboYes
        '
        Me.rboYes.AutoSize = True
        Me.rboYes.Location = New System.Drawing.Point(560, 78)
        Me.rboYes.Name = "rboYes"
        Me.rboYes.Size = New System.Drawing.Size(43, 17)
        Me.rboYes.TabIndex = 8
        Me.rboYes.Text = "Yes"
        Me.rboYes.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(6, 180)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(856, 265)
        Me.TabControl1.TabIndex = 15
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvAddress)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(848, 239)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Address"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvAddress
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvAddress.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvAddress.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAddress.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column11, Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column6, Me.Column8, Me.Column9, Me.Column7, Me.ColIdentifierAdd})
        Me.dgvAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAddress.Location = New System.Drawing.Point(3, 3)
        Me.dgvAddress.MultiSelect = False
        Me.dgvAddress.Name = "dgvAddress"
        Me.dgvAddress.RowTemplate.Height = 24
        Me.dgvAddress.Size = New System.Drawing.Size(842, 233)
        Me.dgvAddress.TabIndex = 0
        '
        'Column11
        '
        Me.Column11.HeaderText = "#"
        Me.Column11.Name = "Column11"
        Me.Column11.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column11.Width = 21
        '
        'Column1
        '
        Me.Column1.HeaderText = "Address Code"
        Me.Column1.Name = "Column1"
        Me.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column1.Width = 79
        '
        'Column2
        '
        Me.Column2.HeaderText = "Line 1"
        Me.Column2.Name = "Column2"
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column2.Width = 42
        '
        'Column3
        '
        Me.Column3.HeaderText = "Line 2"
        Me.Column3.Name = "Column3"
        Me.Column3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column3.Width = 42
        '
        'Column4
        '
        Me.Column4.HeaderText = "Line 3"
        Me.Column4.Name = "Column4"
        Me.Column4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column4.Width = 42
        '
        'Column5
        '
        Me.Column5.HeaderText = "Line 4"
        Me.Column5.Name = "Column5"
        Me.Column5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column5.Width = 42
        '
        'Column6
        '
        Me.Column6.HeaderText = "Line 5"
        Me.Column6.Name = "Column6"
        Me.Column6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column6.Width = 42
        '
        'Column8
        '
        Me.Column8.HeaderText = "Phone Number"
        Me.Column8.Name = "Column8"
        Me.Column8.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column8.Width = 84
        '
        'Column9
        '
        Me.Column9.HeaderText = "Fax Number"
        Me.Column9.Name = "Column9"
        Me.Column9.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column9.Width = 70
        '
        'Column7
        '
        Me.Column7.HeaderText = "Remarks"
        Me.Column7.Name = "Column7"
        Me.Column7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column7.Width = 55
        '
        'ColIdentifierAdd
        '
        Me.ColIdentifierAdd.HeaderText = "Identifier"
        Me.ColIdentifierAdd.Name = "ColIdentifierAdd"
        Me.ColIdentifierAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColIdentifierAdd.Visible = False
        Me.ColIdentifierAdd.Width = 53
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvCP)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(848, 239)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Contact Person"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'dgvCP
        '
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvCP.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvCP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPersonCode, Me.ColPersonName, Me.Column10, Me.ColPhoneNumber, Me.ColFax, Me.ColEmail, Me.Column13, Me.Column18})
        Me.dgvCP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCP.Location = New System.Drawing.Point(3, 3)
        Me.dgvCP.MultiSelect = False
        Me.dgvCP.Name = "dgvCP"
        Me.dgvCP.RowTemplate.Height = 24
        Me.dgvCP.Size = New System.Drawing.Size(842, 233)
        Me.dgvCP.TabIndex = 59
        '
        'ColPersonCode
        '
        Me.ColPersonCode.HeaderText = "#"
        Me.ColPersonCode.Name = "ColPersonCode"
        Me.ColPersonCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonCode.Width = 21
        '
        'ColPersonName
        '
        Me.ColPersonName.HeaderText = "Contact Person Name"
        Me.ColPersonName.Name = "ColPersonName"
        Me.ColPersonName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonName.Width = 117
        '
        'Column10
        '
        Me.Column10.HeaderText = "Position"
        Me.Column10.Name = "Column10"
        Me.Column10.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column10.Width = 50
        '
        'ColPhoneNumber
        '
        Me.ColPhoneNumber.HeaderText = "Phone Number"
        Me.ColPhoneNumber.Name = "ColPhoneNumber"
        Me.ColPhoneNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPhoneNumber.Width = 84
        '
        'ColFax
        '
        Me.ColFax.HeaderText = "Fax Number"
        Me.ColFax.Name = "ColFax"
        Me.ColFax.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFax.Width = 70
        '
        'ColEmail
        '
        Me.ColEmail.HeaderText = "Email Address"
        Me.ColEmail.Name = "ColEmail"
        Me.ColEmail.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmail.Width = 79
        '
        'Column13
        '
        Me.Column13.HeaderText = "Remarks"
        Me.Column13.Name = "Column13"
        Me.Column13.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column13.Width = 55
        '
        'Column18
        '
        Me.Column18.HeaderText = "Identifier"
        Me.Column18.Name = "Column18"
        Me.Column18.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column18.Visible = False
        Me.Column18.Width = 53
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.dgvBank)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(848, 239)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Bank Account"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvBank
        '
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvBank.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvBank.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvBank.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column12, Me.ColNamaBank, Me.ColAccountNumber, Me.Column14, Me.ColAddress1, Me.ColBeneficiaryName, Me.Column15, Me.Column17, Me.ColPrimaryCurrency, Me.Column19, Me.ColSwiftCode, Me.ColBankBranch})
        Me.dgvBank.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvBank.Location = New System.Drawing.Point(3, 3)
        Me.dgvBank.MultiSelect = False
        Me.dgvBank.Name = "dgvBank"
        Me.dgvBank.RowTemplate.Height = 24
        Me.dgvBank.Size = New System.Drawing.Size(842, 233)
        Me.dgvBank.TabIndex = 47
        '
        'Column12
        '
        Me.Column12.HeaderText = "#"
        Me.Column12.Name = "Column12"
        Me.Column12.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column12.Width = 21
        '
        'ColNamaBank
        '
        Me.ColNamaBank.HeaderText = "Bank's Name"
        Me.ColNamaBank.Name = "ColNamaBank"
        Me.ColNamaBank.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNamaBank.Width = 76
        '
        'ColAccountNumber
        '
        Me.ColAccountNumber.HeaderText = "Account Number"
        Me.ColAccountNumber.Name = "ColAccountNumber"
        Me.ColAccountNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAccountNumber.Width = 93
        '
        'Column14
        '
        Me.Column14.HeaderText = "Currency"
        Me.Column14.Name = "Column14"
        Me.Column14.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column14.Width = 55
        '
        'ColAddress1
        '
        Me.ColAddress1.HeaderText = "Address"
        Me.ColAddress1.Name = "ColAddress1"
        Me.ColAddress1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAddress1.Width = 51
        '
        'ColBeneficiaryName
        '
        Me.ColBeneficiaryName.HeaderText = "Beneficiary Name"
        Me.ColBeneficiaryName.Name = "ColBeneficiaryName"
        Me.ColBeneficiaryName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBeneficiaryName.Width = 96
        '
        'Column15
        '
        Me.Column15.HeaderText = "Phone No"
        Me.Column15.Name = "Column15"
        Me.Column15.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column15.Width = 61
        '
        'Column17
        '
        Me.Column17.HeaderText = "Remarks"
        Me.Column17.Name = "Column17"
        Me.Column17.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column17.Width = 55
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 92
        '
        'Column19
        '
        Me.Column19.HeaderText = "Identifier"
        Me.Column19.Name = "Column19"
        Me.Column19.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column19.Visible = False
        Me.Column19.Width = 53
        '
        'ColSwiftCode
        '
        Me.ColSwiftCode.HeaderText = "Swift Code"
        Me.ColSwiftCode.Name = "ColSwiftCode"
        Me.ColSwiftCode.Width = 83
        '
        'ColBankBranch
        '
        Me.ColBankBranch.HeaderText = "Bank Branch"
        Me.ColBankBranch.Name = "ColBankBranch"
        Me.ColBankBranch.Width = 94
        '
        'TxtCompanyWebsite
        '
        Me.TxtCompanyWebsite.Location = New System.Drawing.Point(111, 155)
        Me.TxtCompanyWebsite.MaxLength = 100
        Me.TxtCompanyWebsite.Name = "TxtCompanyWebsite"
        Me.TxtCompanyWebsite.Size = New System.Drawing.Size(304, 20)
        Me.TxtCompanyWebsite.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 158)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 13)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Company Website"
        '
        'btnFindCurrency
        '
        Me.btnFindCurrency.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCurrency.Location = New System.Drawing.Point(820, 50)
        Me.btnFindCurrency.Name = "btnFindCurrency"
        Me.btnFindCurrency.Size = New System.Drawing.Size(42, 23)
        Me.btnFindCurrency.TabIndex = 7
        Me.btnFindCurrency.TabStop = False
        Me.btnFindCurrency.Text = "...."
        Me.btnFindCurrency.UseVisualStyleBackColor = True
        '
        'TxtCurrencyCode
        '
        Me.TxtCurrencyCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCurrencyCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCurrencyCode.Location = New System.Drawing.Point(560, 52)
        Me.TxtCurrencyCode.MaxLength = 25
        Me.TxtCurrencyCode.Name = "TxtCurrencyCode"
        Me.TxtCurrencyCode.Size = New System.Drawing.Size(254, 20)
        Me.TxtCurrencyCode.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(460, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 13)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "Default Currency"
        '
        'TxtNPWP
        '
        Me.TxtNPWP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWP.Location = New System.Drawing.Point(560, 104)
        Me.TxtNPWP.MaxLength = 25
        Me.TxtNPWP.Name = "TxtNPWP"
        Me.TxtNPWP.Size = New System.Drawing.Size(133, 20)
        Me.TxtNPWP.TabIndex = 10
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(458, 104)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(80, 13)
        Me.Label16.TabIndex = 34
        Me.Label16.Text = "NPWP Number"
        '
        'txtLeadTime
        '
        Me.txtLeadTime.Location = New System.Drawing.Point(111, 132)
        Me.txtLeadTime.MaxLength = 3
        Me.txtLeadTime.Name = "txtLeadTime"
        Me.txtLeadTime.Size = New System.Drawing.Size(304, 20)
        Me.txtLeadTime.TabIndex = 4
        Me.txtLeadTime.Text = "0"
        Me.txtLeadTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(10, 134)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(88, 13)
        Me.Label15.TabIndex = 32
        Me.Label15.Text = "Lead Time (days)"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(458, 78)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(83, 13)
        Me.Label14.TabIndex = 30
        Me.Label14.Text = "Foreign Supplier"
        '
        'TxtCountry
        '
        Me.TxtCountry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCountry.Location = New System.Drawing.Point(560, 26)
        Me.TxtCountry.MaxLength = 50
        Me.TxtCountry.Name = "TxtCountry"
        Me.TxtCountry.Size = New System.Drawing.Size(302, 20)
        Me.TxtCountry.TabIndex = 6
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(460, 31)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(43, 13)
        Me.Label13.TabIndex = 28
        Me.Label13.Text = "Country"
        '
        'TxtPaymentTerm
        '
        Me.TxtPaymentTerm.Location = New System.Drawing.Point(111, 106)
        Me.TxtPaymentTerm.MaxLength = 100
        Me.TxtPaymentTerm.Name = "TxtPaymentTerm"
        Me.TxtPaymentTerm.Size = New System.Drawing.Size(304, 20)
        Me.TxtPaymentTerm.TabIndex = 3
        Me.TxtPaymentTerm.Text = "0"
        Me.TxtPaymentTerm.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(10, 111)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(75, 13)
        Me.Label12.TabIndex = 26
        Me.Label12.Text = "Payment Term"
        '
        'CboPriceType
        '
        Me.CboPriceType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPriceType.FormattingEnabled = True
        Me.CboPriceType.Items.AddRange(New Object() {"exclude", "include"})
        Me.CboPriceType.Location = New System.Drawing.Point(111, 79)
        Me.CboPriceType.Name = "CboPriceType"
        Me.CboPriceType.Size = New System.Drawing.Size(304, 21)
        Me.CboPriceType.TabIndex = 2
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(10, 84)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(58, 13)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Price Type"
        '
        'TxtPrimaryCurrency
        '
        Me.TxtPrimaryCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryCurrency.Location = New System.Drawing.Point(536, 52)
        Me.TxtPrimaryCurrency.Name = "TxtPrimaryCurrency"
        Me.TxtPrimaryCurrency.Size = New System.Drawing.Size(16, 20)
        Me.TxtPrimaryCurrency.TabIndex = 23
        Me.TxtPrimaryCurrency.TabStop = False
        Me.TxtPrimaryCurrency.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(439, 451)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 17
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(323, 451)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 16
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'txtsuppliername
        '
        Me.txtsuppliername.Location = New System.Drawing.Point(111, 53)
        Me.txtsuppliername.MaxLength = 150
        Me.txtsuppliername.Name = "txtsuppliername"
        Me.txtsuppliername.Size = New System.Drawing.Size(304, 20)
        Me.txtsuppliername.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Supplier Name"
        '
        'txtsuppliercode
        '
        Me.txtsuppliercode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtsuppliercode.Location = New System.Drawing.Point(111, 27)
        Me.txtsuppliercode.MaxLength = 50
        Me.txtsuppliercode.Name = "txtsuppliercode"
        Me.txtsuppliercode.Size = New System.Drawing.Size(304, 20)
        Me.txtsuppliercode.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Supplier Code"
        '
        'txtprimarysupplier
        '
        Me.txtprimarysupplier.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtprimarysupplier.Location = New System.Drawing.Point(89, 26)
        Me.txtprimarysupplier.Name = "txtprimarysupplier"
        Me.txtprimarysupplier.Size = New System.Drawing.Size(16, 20)
        Me.txtprimarysupplier.TabIndex = 0
        Me.txtprimarysupplier.TabStop = False
        Me.txtprimarysupplier.Visible = False
        '
        'TxtNPWPLama
        '
        Me.TxtNPWPLama.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWPLama.Location = New System.Drawing.Point(560, 104)
        Me.TxtNPWPLama.MaxLength = 25
        Me.TxtNPWPLama.Name = "TxtNPWPLama"
        Me.TxtNPWPLama.Size = New System.Drawing.Size(90, 20)
        Me.TxtNPWPLama.TabIndex = 42
        Me.TxtNPWPLama.TabStop = False
        Me.TxtNPWPLama.Visible = False
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(627, 493)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(110, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(511, 493)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(110, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(395, 493)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(110, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(279, 493)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(110, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(163, 493)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(110, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 531)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(890, 22)
        Me.StatusStrip1.TabIndex = 17
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(218, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(218, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(218, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(218, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'FrmMasterSupplier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(890, 553)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(898, 576)
        Me.Name = "FrmMasterSupplier"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Supplier Data"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvAddress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvCP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvBank, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents txtprimarysupplier As System.Windows.Forms.TextBox
    Friend WithEvents txtsuppliercode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtsuppliername As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents TxtPaymentTerm As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents CboPriceType As System.Windows.Forms.ComboBox
    Friend WithEvents TxtCountry As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents txtLeadTime As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TxtNPWP As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents btnFindCurrency As System.Windows.Forms.Button
    Friend WithEvents TxtCurrencyCode As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtCompanyWebsite As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents dgvAddress As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents dgvCP As System.Windows.Forms.DataGridView
    Friend WithEvents ColPersonCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPhoneNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column18 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvBank As System.Windows.Forms.DataGridView
    Friend WithEvents rboNo As System.Windows.Forms.RadioButton
    Friend WithEvents rboYes As System.Windows.Forms.RadioButton
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierAdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNamaBank As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccountNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAddress1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBeneficiaryName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column15 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column17 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column19 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSwiftCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBankBranch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LlbEditNPWPLama As System.Windows.Forms.LinkLabel
    Friend WithEvents CboTipeNPWP As System.Windows.Forms.ComboBox
    Friend WithEvents TxtNPWPLama As System.Windows.Forms.TextBox
    Friend WithEvents TxtEmail As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtTKUID As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
