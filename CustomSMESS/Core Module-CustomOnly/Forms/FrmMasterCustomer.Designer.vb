<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterCustomer
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
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.dgvAddress = New System.Windows.Forms.DataGridView
        Me.Column11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_Line5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPhoneAdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColFaxAdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarksAdd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAddressIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.gboGrouping = New System.Windows.Forms.GroupBox
        Me.TxtCustomerGroupCode = New System.Windows.Forms.TextBox
        Me.BtnFindCustomertype = New System.Windows.Forms.Button
        Me.TxtCustomerType = New System.Windows.Forms.TextBox
        Me.BtnFindCustomerGroup = New System.Windows.Forms.Button
        Me.BtnFindSalesArea = New System.Windows.Forms.Button
        Me.Label25 = New System.Windows.Forms.Label
        Me.TxtSalesArea = New System.Windows.Forms.TextBox
        Me.TxtCustomerSegmentation = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnFindCustSegmentation = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtPrimaryCustomerGroup = New System.Windows.Forms.TextBox
        Me.TxtPrimarySalesArea = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtPrimaryCustomerSegmentation = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCustomerType = New System.Windows.Forms.TextBox
        Me.gboTax = New System.Windows.Forms.GroupBox
        Me.TxtNPWP = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.TxtTaxCode = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.TxtAddressPajak1 = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtAddressPajak2 = New System.Windows.Forms.TextBox
        Me.txtpaymentterm = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.TxtCreditLimit = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.TxtZipCode = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TxtCountry = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.TxtCity = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.TxtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
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
        Me.gboGrouping.SuspendLayout()
        Me.gboTax.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.gboGrouping)
        Me.gboDetail.Controls.Add(Me.gboTax)
        Me.gboDetail.Controls.Add(Me.txtpaymentterm)
        Me.gboDetail.Controls.Add(Me.Label24)
        Me.gboDetail.Controls.Add(Me.TxtCreditLimit)
        Me.gboDetail.Controls.Add(Me.Label23)
        Me.gboDetail.Controls.Add(Me.TxtZipCode)
        Me.gboDetail.Controls.Add(Me.Label10)
        Me.gboDetail.Controls.Add(Me.TxtCountry)
        Me.gboDetail.Controls.Add(Me.Label9)
        Me.gboDetail.Controls.Add(Me.TxtCity)
        Me.gboDetail.Controls.Add(Me.Label8)
        Me.gboDetail.Controls.Add(Me.TxtCustomerName)
        Me.gboDetail.Controls.Add(Me.TxtCustomerCode)
        Me.gboDetail.Controls.Add(Me.TxtPrimaryCustomer)
        Me.gboDetail.Controls.Add(Me.Label2)
        Me.gboDetail.Controls.Add(Me.Label1)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Location = New System.Drawing.Point(17, 5)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Size = New System.Drawing.Size(852, 485)
        Me.gboDetail.TabIndex = 5
        Me.gboDetail.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(6, 239)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(840, 208)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvAddress)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(832, 182)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Address"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'dgvAddress
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvAddress.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvAddress.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAddress.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column11, Me.Column1, Me.Col_Line1, Me.Col_Line2, Me.Col_Line3, Me.Col_Line4, Me.Col_Line5, Me.ColPhoneAdd, Me.ColFaxAdd, Me.ColRemarksAdd, Me.ColAddressIdentifier})
        Me.dgvAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvAddress.Location = New System.Drawing.Point(3, 3)
        Me.dgvAddress.MultiSelect = False
        Me.dgvAddress.Name = "dgvAddress"
        Me.dgvAddress.Size = New System.Drawing.Size(826, 176)
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
        'Col_Line1
        '
        Me.Col_Line1.HeaderText = "Line 1"
        Me.Col_Line1.Name = "Col_Line1"
        Me.Col_Line1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line1.Width = 42
        '
        'Col_Line2
        '
        Me.Col_Line2.HeaderText = "Line 2"
        Me.Col_Line2.Name = "Col_Line2"
        Me.Col_Line2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line2.Width = 42
        '
        'Col_Line3
        '
        Me.Col_Line3.HeaderText = "Line 3"
        Me.Col_Line3.Name = "Col_Line3"
        Me.Col_Line3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line3.Width = 42
        '
        'Col_Line4
        '
        Me.Col_Line4.HeaderText = "Line 4"
        Me.Col_Line4.Name = "Col_Line4"
        Me.Col_Line4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line4.Width = 42
        '
        'Col_Line5
        '
        Me.Col_Line5.HeaderText = "Line 5"
        Me.Col_Line5.Name = "Col_Line5"
        Me.Col_Line5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line5.Width = 42
        '
        'ColPhoneAdd
        '
        Me.ColPhoneAdd.HeaderText = "Phone Number"
        Me.ColPhoneAdd.Name = "ColPhoneAdd"
        Me.ColPhoneAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPhoneAdd.Width = 84
        '
        'ColFaxAdd
        '
        Me.ColFaxAdd.HeaderText = "Fax Number"
        Me.ColFaxAdd.Name = "ColFaxAdd"
        Me.ColFaxAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFaxAdd.Width = 70
        '
        'ColRemarksAdd
        '
        Me.ColRemarksAdd.HeaderText = "Remarks"
        Me.ColRemarksAdd.Name = "ColRemarksAdd"
        Me.ColRemarksAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarksAdd.Width = 55
        '
        'ColAddressIdentifier
        '
        Me.ColAddressIdentifier.HeaderText = "Identifier"
        Me.ColAddressIdentifier.Name = "ColAddressIdentifier"
        Me.ColAddressIdentifier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAddressIdentifier.Visible = False
        Me.ColAddressIdentifier.Width = 53
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgvCP)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(832, 182)
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
        Me.dgvCP.Size = New System.Drawing.Size(826, 176)
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
        Me.TabPage2.Size = New System.Drawing.Size(832, 182)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Bank Account"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'dgvBank
        '
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvBank.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvBank.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvBank.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column12, Me.ColNamaBank, Me.ColAccountNumber, Me.Column14, Me.ColAddress1, Me.ColBeneficiaryName, Me.Column15, Me.Column17, Me.ColPrimaryCurrency, Me.Column19})
        Me.dgvBank.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvBank.Location = New System.Drawing.Point(3, 3)
        Me.dgvBank.MultiSelect = False
        Me.dgvBank.Name = "dgvBank"
        Me.dgvBank.Size = New System.Drawing.Size(826, 176)
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
        Me.ColPrimaryCurrency.Width = 83
        '
        'Column19
        '
        Me.Column19.HeaderText = "Identifier"
        Me.Column19.Name = "Column19"
        Me.Column19.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column19.Visible = False
        Me.Column19.Width = 53
        '
        'gboGrouping
        '
        Me.gboGrouping.Controls.Add(Me.TxtCustomerGroupCode)
        Me.gboGrouping.Controls.Add(Me.BtnFindCustomertype)
        Me.gboGrouping.Controls.Add(Me.TxtCustomerType)
        Me.gboGrouping.Controls.Add(Me.BtnFindCustomerGroup)
        Me.gboGrouping.Controls.Add(Me.BtnFindSalesArea)
        Me.gboGrouping.Controls.Add(Me.Label25)
        Me.gboGrouping.Controls.Add(Me.TxtSalesArea)
        Me.gboGrouping.Controls.Add(Me.TxtCustomerSegmentation)
        Me.gboGrouping.Controls.Add(Me.Label4)
        Me.gboGrouping.Controls.Add(Me.btnFindCustSegmentation)
        Me.gboGrouping.Controls.Add(Me.Label3)
        Me.gboGrouping.Controls.Add(Me.TxtPrimaryCustomerGroup)
        Me.gboGrouping.Controls.Add(Me.TxtPrimarySalesArea)
        Me.gboGrouping.Controls.Add(Me.Label5)
        Me.gboGrouping.Controls.Add(Me.TxtPrimaryCustomerSegmentation)
        Me.gboGrouping.Controls.Add(Me.TxtPrimaryCustomerType)
        Me.gboGrouping.Location = New System.Drawing.Point(6, 108)
        Me.gboGrouping.Name = "gboGrouping"
        Me.gboGrouping.Size = New System.Drawing.Size(417, 125)
        Me.gboGrouping.TabIndex = 7
        Me.gboGrouping.TabStop = False
        Me.gboGrouping.Text = "Grouping"
        '
        'TxtCustomerGroupCode
        '
        Me.TxtCustomerGroupCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerGroupCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerGroupCode.Location = New System.Drawing.Point(140, 15)
        Me.TxtCustomerGroupCode.MaxLength = 50
        Me.TxtCustomerGroupCode.Name = "TxtCustomerGroupCode"
        Me.TxtCustomerGroupCode.Size = New System.Drawing.Size(227, 20)
        Me.TxtCustomerGroupCode.TabIndex = 0
        '
        'BtnFindCustomertype
        '
        Me.BtnFindCustomertype.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCustomertype.Location = New System.Drawing.Point(373, 94)
        Me.BtnFindCustomertype.Name = "BtnFindCustomertype"
        Me.BtnFindCustomertype.Size = New System.Drawing.Size(38, 20)
        Me.BtnFindCustomertype.TabIndex = 7
        Me.BtnFindCustomertype.TabStop = False
        Me.BtnFindCustomertype.Text = "...."
        Me.BtnFindCustomertype.UseVisualStyleBackColor = True
        '
        'TxtCustomerType
        '
        Me.TxtCustomerType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerType.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerType.Location = New System.Drawing.Point(140, 97)
        Me.TxtCustomerType.MaxLength = 50
        Me.TxtCustomerType.Name = "TxtCustomerType"
        Me.TxtCustomerType.Size = New System.Drawing.Size(227, 20)
        Me.TxtCustomerType.TabIndex = 6
        '
        'BtnFindCustomerGroup
        '
        Me.BtnFindCustomerGroup.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindCustomerGroup.Location = New System.Drawing.Point(373, 19)
        Me.BtnFindCustomerGroup.Name = "BtnFindCustomerGroup"
        Me.BtnFindCustomerGroup.Size = New System.Drawing.Size(38, 20)
        Me.BtnFindCustomerGroup.TabIndex = 1
        Me.BtnFindCustomerGroup.TabStop = False
        Me.BtnFindCustomerGroup.Text = "...."
        Me.BtnFindCustomerGroup.UseVisualStyleBackColor = True
        '
        'BtnFindSalesArea
        '
        Me.BtnFindSalesArea.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFindSalesArea.Location = New System.Drawing.Point(373, 45)
        Me.BtnFindSalesArea.Name = "BtnFindSalesArea"
        Me.BtnFindSalesArea.Size = New System.Drawing.Size(38, 20)
        Me.BtnFindSalesArea.TabIndex = 3
        Me.BtnFindSalesArea.TabStop = False
        Me.BtnFindSalesArea.Text = "...."
        Me.BtnFindSalesArea.UseVisualStyleBackColor = True
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(9, 71)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(119, 13)
        Me.Label25.TabIndex = 74
        Me.Label25.Text = "Customer Segmentation"
        '
        'TxtSalesArea
        '
        Me.TxtSalesArea.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtSalesArea.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtSalesArea.Location = New System.Drawing.Point(140, 42)
        Me.TxtSalesArea.MaxLength = 50
        Me.TxtSalesArea.Name = "TxtSalesArea"
        Me.TxtSalesArea.Size = New System.Drawing.Size(227, 20)
        Me.TxtSalesArea.TabIndex = 2
        '
        'TxtCustomerSegmentation
        '
        Me.TxtCustomerSegmentation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerSegmentation.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerSegmentation.Location = New System.Drawing.Point(140, 68)
        Me.TxtCustomerSegmentation.MaxLength = 50
        Me.TxtCustomerSegmentation.Name = "TxtCustomerSegmentation"
        Me.TxtCustomerSegmentation.Size = New System.Drawing.Size(227, 20)
        Me.TxtCustomerSegmentation.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 49)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 81
        Me.Label4.Text = "Sales Area"
        '
        'btnFindCustSegmentation
        '
        Me.btnFindCustSegmentation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindCustSegmentation.Location = New System.Drawing.Point(373, 68)
        Me.btnFindCustSegmentation.Name = "btnFindCustSegmentation"
        Me.btnFindCustSegmentation.Size = New System.Drawing.Size(38, 20)
        Me.btnFindCustSegmentation.TabIndex = 5
        Me.btnFindCustSegmentation.TabStop = False
        Me.btnFindCustSegmentation.Text = "...."
        Me.btnFindCustSegmentation.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 13)
        Me.Label3.TabIndex = 79
        Me.Label3.Text = "Customer Group"
        '
        'TxtPrimaryCustomerGroup
        '
        Me.TxtPrimaryCustomerGroup.Location = New System.Drawing.Point(36, 19)
        Me.TxtPrimaryCustomerGroup.Name = "TxtPrimaryCustomerGroup"
        Me.TxtPrimaryCustomerGroup.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCustomerGroup.TabIndex = 26
        Me.TxtPrimaryCustomerGroup.TabStop = False
        Me.TxtPrimaryCustomerGroup.Visible = False
        '
        'TxtPrimarySalesArea
        '
        Me.TxtPrimarySalesArea.Location = New System.Drawing.Point(36, 45)
        Me.TxtPrimarySalesArea.Name = "TxtPrimarySalesArea"
        Me.TxtPrimarySalesArea.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimarySalesArea.TabIndex = 27
        Me.TxtPrimarySalesArea.TabStop = False
        Me.TxtPrimarySalesArea.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(9, 97)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(78, 13)
        Me.Label5.TabIndex = 85
        Me.Label5.Text = "Customer Type"
        '
        'TxtPrimaryCustomerSegmentation
        '
        Me.TxtPrimaryCustomerSegmentation.Location = New System.Drawing.Point(36, 71)
        Me.TxtPrimaryCustomerSegmentation.Name = "TxtPrimaryCustomerSegmentation"
        Me.TxtPrimaryCustomerSegmentation.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCustomerSegmentation.TabIndex = 28
        Me.TxtPrimaryCustomerSegmentation.Visible = False
        '
        'TxtPrimaryCustomerType
        '
        Me.TxtPrimaryCustomerType.Location = New System.Drawing.Point(36, 98)
        Me.TxtPrimaryCustomerType.Name = "TxtPrimaryCustomerType"
        Me.TxtPrimaryCustomerType.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCustomerType.TabIndex = 29
        Me.TxtPrimaryCustomerType.TabStop = False
        Me.TxtPrimaryCustomerType.Visible = False
        '
        'gboTax
        '
        Me.gboTax.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboTax.Controls.Add(Me.TxtNPWP)
        Me.gboTax.Controls.Add(Me.Label19)
        Me.gboTax.Controls.Add(Me.Label20)
        Me.gboTax.Controls.Add(Me.TxtTaxCode)
        Me.gboTax.Controls.Add(Me.Label22)
        Me.gboTax.Controls.Add(Me.TxtAddressPajak1)
        Me.gboTax.Controls.Add(Me.Label21)
        Me.gboTax.Controls.Add(Me.TxtAddressPajak2)
        Me.gboTax.Location = New System.Drawing.Point(429, 108)
        Me.gboTax.Name = "gboTax"
        Me.gboTax.Size = New System.Drawing.Size(417, 124)
        Me.gboTax.TabIndex = 8
        Me.gboTax.TabStop = False
        Me.gboTax.Text = "Tax Data"
        '
        'TxtNPWP
        '
        Me.TxtNPWP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWP.Location = New System.Drawing.Point(127, 19)
        Me.TxtNPWP.MaxLength = 20
        Me.TxtNPWP.Name = "TxtNPWP"
        Me.TxtNPWP.Size = New System.Drawing.Size(284, 20)
        Me.TxtNPWP.TabIndex = 1
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(16, 22)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(57, 13)
        Me.Label19.TabIndex = 62
        Me.Label19.Text = "NPWP No"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(16, 48)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(90, 13)
        Me.Label20.TabIndex = 64
        Me.Label20.Text = "Default Tax Code"
        '
        'TxtTaxCode
        '
        Me.TxtTaxCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxCode.Location = New System.Drawing.Point(127, 45)
        Me.TxtTaxCode.MaxLength = 7
        Me.TxtTaxCode.Name = "TxtTaxCode"
        Me.TxtTaxCode.Size = New System.Drawing.Size(284, 20)
        Me.TxtTaxCode.TabIndex = 2
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(16, 71)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(90, 13)
        Me.Label22.TabIndex = 66
        Me.Label22.Text = "Tax Addres Line1"
        '
        'TxtAddressPajak1
        '
        Me.TxtAddressPajak1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddressPajak1.Location = New System.Drawing.Point(127, 71)
        Me.TxtAddressPajak1.MaxLength = 100
        Me.TxtAddressPajak1.Name = "TxtAddressPajak1"
        Me.TxtAddressPajak1.Size = New System.Drawing.Size(284, 20)
        Me.TxtAddressPajak1.TabIndex = 3
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(16, 97)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(95, 13)
        Me.Label21.TabIndex = 68
        Me.Label21.Text = "Tax Address Line2"
        '
        'TxtAddressPajak2
        '
        Me.TxtAddressPajak2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddressPajak2.Location = New System.Drawing.Point(127, 97)
        Me.TxtAddressPajak2.MaxLength = 100
        Me.TxtAddressPajak2.Name = "TxtAddressPajak2"
        Me.TxtAddressPajak2.Size = New System.Drawing.Size(284, 20)
        Me.TxtAddressPajak2.TabIndex = 4
        '
        'txtpaymentterm
        '
        Me.txtpaymentterm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtpaymentterm.Location = New System.Drawing.Point(569, 82)
        Me.txtpaymentterm.MaxLength = 3
        Me.txtpaymentterm.Name = "txtpaymentterm"
        Me.txtpaymentterm.Size = New System.Drawing.Size(277, 20)
        Me.txtpaymentterm.TabIndex = 6
        Me.txtpaymentterm.Text = "0"
        Me.txtpaymentterm.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label24
        '
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(439, 85)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(103, 13)
        Me.Label24.TabIndex = 72
        Me.Label24.Text = "PaymentTerm (days)"
        '
        'TxtCreditLimit
        '
        Me.TxtCreditLimit.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCreditLimit.Location = New System.Drawing.Point(569, 56)
        Me.TxtCreditLimit.MaxLength = 25
        Me.TxtCreditLimit.Name = "TxtCreditLimit"
        Me.TxtCreditLimit.Size = New System.Drawing.Size(277, 20)
        Me.TxtCreditLimit.TabIndex = 5
        Me.TxtCreditLimit.Text = "0"
        Me.TxtCreditLimit.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(439, 59)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(86, 13)
        Me.Label23.TabIndex = 70
        Me.Label23.Text = "Credit Limit (IDR)"
        '
        'TxtZipCode
        '
        Me.TxtZipCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtZipCode.Location = New System.Drawing.Point(569, 30)
        Me.TxtZipCode.MaxLength = 50
        Me.TxtZipCode.Name = "TxtZipCode"
        Me.TxtZipCode.Size = New System.Drawing.Size(277, 20)
        Me.TxtZipCode.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(439, 30)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(50, 13)
        Me.Label10.TabIndex = 44
        Me.Label10.Text = "Zip Code"
        '
        'TxtCountry
        '
        Me.TxtCountry.Location = New System.Drawing.Point(205, 82)
        Me.TxtCountry.MaxLength = 50
        Me.TxtCountry.Name = "TxtCountry"
        Me.TxtCountry.Size = New System.Drawing.Size(218, 20)
        Me.TxtCountry.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(161, 85)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(43, 13)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "Country"
        '
        'TxtCity
        '
        Me.TxtCity.Location = New System.Drawing.Point(51, 82)
        Me.TxtCity.MaxLength = 50
        Me.TxtCity.Name = "TxtCity"
        Me.TxtCity.Size = New System.Drawing.Size(104, 20)
        Me.TxtCity.TabIndex = 2
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 85)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(24, 13)
        Me.Label8.TabIndex = 40
        Me.Label8.Text = "City"
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Location = New System.Drawing.Point(146, 56)
        Me.TxtCustomerName.MaxLength = 150
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.Size = New System.Drawing.Size(277, 20)
        Me.TxtCustomerName.TabIndex = 1
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCode.Location = New System.Drawing.Point(146, 30)
        Me.TxtCustomerCode.MaxLength = 50
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.Size = New System.Drawing.Size(277, 20)
        Me.TxtCustomerCode.TabIndex = 0
        '
        'TxtPrimaryCustomer
        '
        Me.TxtPrimaryCustomer.Location = New System.Drawing.Point(110, 30)
        Me.TxtPrimaryCustomer.Name = "TxtPrimaryCustomer"
        Me.TxtPrimaryCustomer.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimaryCustomer.TabIndex = 23
        Me.TxtPrimaryCustomer.TabStop = False
        Me.TxtPrimaryCustomer.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Customer Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Customer Code"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(414, 453)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 26)
        Me.btnCancel.TabIndex = 11
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(298, 453)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 26)
        Me.btnOK.TabIndex = 10
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(598, 498)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(110, 26)
        Me.btnExit.TabIndex = 4
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(482, 498)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(110, 26)
        Me.btnDelete.TabIndex = 3
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(250, 498)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(110, 26)
        Me.btnNew.TabIndex = 1
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(366, 498)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(110, 26)
        Me.btnEdit.TabIndex = 2
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(134, 498)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(110, 26)
        Me.btnFind.TabIndex = 0
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 527)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(884, 22)
        Me.StatusStrip1.TabIndex = 6
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(217, 17)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'FrmMasterCustomer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(884, 549)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.Controls.Add(Me.gboDetail)
        Me.MinimumSize = New System.Drawing.Size(900, 585)
        Me.Name = "FrmMasterCustomer"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customer Data"
        Me.gboDetail.ResumeLayout(False)
        Me.gboDetail.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvAddress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvCP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        CType(Me.dgvBank, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboGrouping.ResumeLayout(False)
        Me.gboGrouping.PerformLayout()
        Me.gboTax.ResumeLayout(False)
        Me.gboTax.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimaryCustomerGroup As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCustomerType As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCustomerSegmentation As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimarySalesArea As System.Windows.Forms.TextBox
    Friend WithEvents TxtZipCode As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtCountry As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtCity As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtTaxCode As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TxtNPWP As System.Windows.Forms.TextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TxtAddressPajak2 As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtAddressPajak1 As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TxtCreditLimit As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents txtpaymentterm As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents btnFindCustSegmentation As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerSegmentation As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents gboTax As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCustomerGroup As System.Windows.Forms.Button
    Friend WithEvents BtnFindSalesArea As System.Windows.Forms.Button
    Friend WithEvents TxtSalesArea As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BtnFindCustomertype As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerType As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents gboGrouping As System.Windows.Forms.GroupBox
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
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvBank As System.Windows.Forms.DataGridView
    Friend WithEvents TxtCustomerGroupCode As System.Windows.Forms.TextBox
    Friend WithEvents ColPersonCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPersonName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPhoneNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFax As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column18 As System.Windows.Forms.DataGridViewTextBoxColumn
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
    Friend WithEvents Column11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_Line5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPhoneAdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColFaxAdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarksAdd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAddressIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
