<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterProspect
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
        Me.TxtNPWP = New System.Windows.Forms.TextBox
        Me.gboDetail = New System.Windows.Forms.GroupBox
        Me.TxtRemarks = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
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
        Me.gboTax = New System.Windows.Forms.GroupBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.TxtTaxCode = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.TxtAddressPajak1 = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.TxtAddressPajak2 = New System.Windows.Forms.TextBox
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
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.tslCreatedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslCreatedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedBy = New System.Windows.Forms.ToolStripStatusLabel
        Me.tslModifiedDate = New System.Windows.Forms.ToolStripStatusLabel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.btnFind = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.gboDetail.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.dgvAddress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgvCP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboTax.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TxtNPWP
        '
        Me.TxtNPWP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtNPWP.Location = New System.Drawing.Point(169, 23)
        Me.TxtNPWP.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtNPWP.MaxLength = 20
        Me.TxtNPWP.Name = "TxtNPWP"
        Me.TxtNPWP.Size = New System.Drawing.Size(640, 22)
        Me.TxtNPWP.TabIndex = 1
        '
        'gboDetail
        '
        Me.gboDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboDetail.Controls.Add(Me.TabControl1)
        Me.gboDetail.Controls.Add(Me.gboTax)
        Me.gboDetail.Controls.Add(Me.btnCancel)
        Me.gboDetail.Controls.Add(Me.btnOK)
        Me.gboDetail.Controls.Add(Me.SplitContainer1)
        Me.gboDetail.Location = New System.Drawing.Point(13, 3)
        Me.gboDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Name = "gboDetail"
        Me.gboDetail.Padding = New System.Windows.Forms.Padding(4)
        Me.gboDetail.Size = New System.Drawing.Size(835, 597)
        Me.gboDetail.TabIndex = 12
        Me.gboDetail.TabStop = False
        '
        'TxtRemarks
        '
        Me.TxtRemarks.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRemarks.Location = New System.Drawing.Point(94, 32)
        Me.TxtRemarks.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtRemarks.MaxLength = 50
        Me.TxtRemarks.Multiline = True
        Me.TxtRemarks.Name = "TxtRemarks"
        Me.TxtRemarks.Size = New System.Drawing.Size(312, 54)
        Me.TxtRemarks.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(22, 35)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 17)
        Me.Label6.TabIndex = 74
        Me.Label6.Text = "Remarks"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(8, 294)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(819, 256)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.dgvAddress)
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(4)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(4)
        Me.TabPage3.Size = New System.Drawing.Size(811, 227)
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
        Me.dgvAddress.Location = New System.Drawing.Point(4, 4)
        Me.dgvAddress.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvAddress.MultiSelect = False
        Me.dgvAddress.Name = "dgvAddress"
        Me.dgvAddress.RowTemplate.Height = 24
        Me.dgvAddress.Size = New System.Drawing.Size(803, 219)
        Me.dgvAddress.TabIndex = 0
        '
        'Column11
        '
        Me.Column11.HeaderText = "#"
        Me.Column11.Name = "Column11"
        Me.Column11.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column11.Width = 22
        '
        'Column1
        '
        Me.Column1.HeaderText = "Address Code"
        Me.Column1.Name = "Column1"
        Me.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column1.Width = 103
        '
        'Col_Line1
        '
        Me.Col_Line1.HeaderText = "Line 1"
        Me.Col_Line1.Name = "Col_Line1"
        Me.Col_Line1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line1.Width = 53
        '
        'Col_Line2
        '
        Me.Col_Line2.HeaderText = "Line 2"
        Me.Col_Line2.Name = "Col_Line2"
        Me.Col_Line2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line2.Width = 53
        '
        'Col_Line3
        '
        Me.Col_Line3.HeaderText = "Line 3"
        Me.Col_Line3.Name = "Col_Line3"
        Me.Col_Line3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line3.Width = 53
        '
        'Col_Line4
        '
        Me.Col_Line4.HeaderText = "Line 4"
        Me.Col_Line4.Name = "Col_Line4"
        Me.Col_Line4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line4.Width = 53
        '
        'Col_Line5
        '
        Me.Col_Line5.HeaderText = "Line 5"
        Me.Col_Line5.Name = "Col_Line5"
        Me.Col_Line5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Col_Line5.Width = 53
        '
        'ColPhoneAdd
        '
        Me.ColPhoneAdd.HeaderText = "Phone Number"
        Me.ColPhoneAdd.Name = "ColPhoneAdd"
        Me.ColPhoneAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPhoneAdd.Width = 109
        '
        'ColFaxAdd
        '
        Me.ColFaxAdd.HeaderText = "Fax Number"
        Me.ColFaxAdd.Name = "ColFaxAdd"
        Me.ColFaxAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFaxAdd.Width = 90
        '
        'ColRemarksAdd
        '
        Me.ColRemarksAdd.HeaderText = "Remarks"
        Me.ColRemarksAdd.Name = "ColRemarksAdd"
        Me.ColRemarksAdd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarksAdd.Width = 70
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
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(4)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(4)
        Me.TabPage1.Size = New System.Drawing.Size(811, 227)
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
        Me.dgvCP.Location = New System.Drawing.Point(4, 4)
        Me.dgvCP.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvCP.MultiSelect = False
        Me.dgvCP.Name = "dgvCP"
        Me.dgvCP.RowTemplate.Height = 24
        Me.dgvCP.Size = New System.Drawing.Size(803, 219)
        Me.dgvCP.TabIndex = 59
        '
        'ColPersonCode
        '
        Me.ColPersonCode.HeaderText = "#"
        Me.ColPersonCode.Name = "ColPersonCode"
        Me.ColPersonCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonCode.Width = 22
        '
        'ColPersonName
        '
        Me.ColPersonName.HeaderText = "Contact Person Name"
        Me.ColPersonName.Name = "ColPersonName"
        Me.ColPersonName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPersonName.Width = 152
        '
        'Column10
        '
        Me.Column10.HeaderText = "Position"
        Me.Column10.Name = "Column10"
        Me.Column10.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column10.Width = 64
        '
        'ColPhoneNumber
        '
        Me.ColPhoneNumber.HeaderText = "Phone Number"
        Me.ColPhoneNumber.Name = "ColPhoneNumber"
        Me.ColPhoneNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPhoneNumber.Width = 109
        '
        'ColFax
        '
        Me.ColFax.HeaderText = "Fax Number"
        Me.ColFax.Name = "ColFax"
        Me.ColFax.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColFax.Width = 90
        '
        'ColEmail
        '
        Me.ColEmail.HeaderText = "Email Address"
        Me.ColEmail.Name = "ColEmail"
        Me.ColEmail.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmail.Width = 104
        '
        'Column13
        '
        Me.Column13.HeaderText = "Remarks"
        Me.Column13.Name = "Column13"
        Me.Column13.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column13.Width = 70
        '
        'Column18
        '
        Me.Column18.HeaderText = "Identifier"
        Me.Column18.Name = "Column18"
        Me.Column18.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column18.Visible = False
        Me.Column18.Width = 53
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
        Me.gboTax.Location = New System.Drawing.Point(8, 133)
        Me.gboTax.Margin = New System.Windows.Forms.Padding(4)
        Me.gboTax.Name = "gboTax"
        Me.gboTax.Padding = New System.Windows.Forms.Padding(4)
        Me.gboTax.Size = New System.Drawing.Size(819, 153)
        Me.gboTax.TabIndex = 9
        Me.gboTax.TabStop = False
        Me.gboTax.Text = "Tax Data"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(21, 27)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(71, 17)
        Me.Label19.TabIndex = 62
        Me.Label19.Text = "NPWP No"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(21, 59)
        Me.Label20.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(117, 17)
        Me.Label20.TabIndex = 64
        Me.Label20.Text = "Default Tax Code"
        '
        'TxtTaxCode
        '
        Me.TxtTaxCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtTaxCode.Location = New System.Drawing.Point(169, 55)
        Me.TxtTaxCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTaxCode.MaxLength = 7
        Me.TxtTaxCode.Name = "TxtTaxCode"
        Me.TxtTaxCode.Size = New System.Drawing.Size(640, 22)
        Me.TxtTaxCode.TabIndex = 2
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(21, 87)
        Me.Label22.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(119, 17)
        Me.Label22.TabIndex = 66
        Me.Label22.Text = "Tax Addres Line1"
        '
        'TxtAddressPajak1
        '
        Me.TxtAddressPajak1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddressPajak1.Location = New System.Drawing.Point(169, 87)
        Me.TxtAddressPajak1.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAddressPajak1.MaxLength = 100
        Me.TxtAddressPajak1.Name = "TxtAddressPajak1"
        Me.TxtAddressPajak1.Size = New System.Drawing.Size(640, 22)
        Me.TxtAddressPajak1.TabIndex = 3
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(21, 119)
        Me.Label21.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(126, 17)
        Me.Label21.TabIndex = 68
        Me.Label21.Text = "Tax Address Line2"
        '
        'TxtAddressPajak2
        '
        Me.TxtAddressPajak2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtAddressPajak2.Location = New System.Drawing.Point(169, 119)
        Me.TxtAddressPajak2.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtAddressPajak2.MaxLength = 100
        Me.TxtAddressPajak2.Name = "TxtAddressPajak2"
        Me.TxtAddressPajak2.Size = New System.Drawing.Size(640, 22)
        Me.TxtAddressPajak2.TabIndex = 4
        '
        'TxtZipCode
        '
        Me.TxtZipCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtZipCode.Location = New System.Drawing.Point(94, 0)
        Me.TxtZipCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtZipCode.MaxLength = 50
        Me.TxtZipCode.Name = "TxtZipCode"
        Me.TxtZipCode.Size = New System.Drawing.Size(312, 22)
        Me.TxtZipCode.TabIndex = 4
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(22, 4)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(65, 17)
        Me.Label10.TabIndex = 44
        Me.Label10.Text = "Zip Code"
        '
        'TxtCountry
        '
        Me.TxtCountry.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCountry.Location = New System.Drawing.Point(266, 65)
        Me.TxtCountry.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCountry.MaxLength = 50
        Me.TxtCountry.Name = "TxtCountry"
        Me.TxtCountry.Size = New System.Drawing.Size(138, 22)
        Me.TxtCountry.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(208, 69)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(57, 17)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "Country"
        '
        'TxtCity
        '
        Me.TxtCity.Location = New System.Drawing.Point(61, 65)
        Me.TxtCity.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCity.MaxLength = 50
        Me.TxtCity.Name = "TxtCity"
        Me.TxtCity.Size = New System.Drawing.Size(137, 22)
        Me.TxtCity.TabIndex = 2
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(14, 69)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(31, 17)
        Me.Label8.TabIndex = 40
        Me.Label8.Text = "City"
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerName.Location = New System.Drawing.Point(108, 33)
        Me.TxtCustomerName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCustomerName.MaxLength = 150
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.Size = New System.Drawing.Size(297, 22)
        Me.TxtCustomerName.TabIndex = 1
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCode.Location = New System.Drawing.Point(108, 1)
        Me.TxtCustomerCode.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCustomerCode.MaxLength = 50
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.Size = New System.Drawing.Size(297, 22)
        Me.TxtCustomerCode.TabIndex = 0
        '
        'TxtPrimaryCustomer
        '
        Me.TxtPrimaryCustomer.Location = New System.Drawing.Point(61, -2)
        Me.TxtPrimaryCustomer.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtPrimaryCustomer.Name = "TxtPrimaryCustomer"
        Me.TxtPrimaryCustomer.Size = New System.Drawing.Size(39, 22)
        Me.TxtPrimaryCustomer.TabIndex = 23
        Me.TxtPrimaryCustomer.TabStop = False
        Me.TxtPrimaryCustomer.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 33)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 17)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 1)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 17)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Code"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(421, 558)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(147, 32)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.Location = New System.Drawing.Point(266, 558)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(147, 32)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslCreatedBy, Me.tslCreatedDate, Me.tslModifiedBy, Me.tslModifiedDate})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 651)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(861, 25)
        Me.StatusStrip1.TabIndex = 13
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tslCreatedBy
        '
        Me.tslCreatedBy.Name = "tslCreatedBy"
        Me.tslCreatedBy.Size = New System.Drawing.Size(210, 20)
        Me.tslCreatedBy.Spring = True
        Me.tslCreatedBy.Text = "Created By:"
        '
        'tslCreatedDate
        '
        Me.tslCreatedDate.Name = "tslCreatedDate"
        Me.tslCreatedDate.Size = New System.Drawing.Size(210, 20)
        Me.tslCreatedDate.Spring = True
        Me.tslCreatedDate.Text = "Created Date:"
        '
        'tslModifiedBy
        '
        Me.tslModifiedBy.Name = "tslModifiedBy"
        Me.tslModifiedBy.Size = New System.Drawing.Size(210, 20)
        Me.tslModifiedBy.Spring = True
        Me.tslModifiedBy.Text = "Modified By:"
        '
        'tslModifiedDate
        '
        Me.tslModifiedDate.Name = "tslModifiedDate"
        Me.tslModifiedDate.Size = New System.Drawing.Size(210, 20)
        Me.tslModifiedDate.Spring = True
        Me.tslModifiedDate.Text = "Modified Date:"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnDelete.Location = New System.Drawing.Point(512, 610)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(147, 32)
        Me.btnDelete.TabIndex = 10
        Me.btnDelete.Text = "&Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(666, 610)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(147, 32)
        Me.btnExit.TabIndex = 11
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnNew.Location = New System.Drawing.Point(202, 610)
        Me.btnNew.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(147, 32)
        Me.btnNew.TabIndex = 8
        Me.btnNew.Text = "&New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnEdit.Location = New System.Drawing.Point(357, 610)
        Me.btnEdit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(147, 32)
        Me.btnEdit.TabIndex = 9
        Me.btnEdit.Text = "&Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnFind
        '
        Me.btnFind.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnFind.Location = New System.Drawing.Point(48, 610)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(147, 32)
        Me.btnFind.TabIndex = 7
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(8, 22)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerCode)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCountry)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryCustomer)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCity)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtCustomerName)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtRemarks)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtZipCode)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label10)
        Me.SplitContainer1.Size = New System.Drawing.Size(819, 104)
        Me.SplitContainer1.SplitterDistance = 409
        Me.SplitContainer1.TabIndex = 75
        '
        'FrmMasterProspect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(861, 676)
        Me.Controls.Add(Me.gboDetail)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnFind)
        Me.MinimumSize = New System.Drawing.Size(869, 713)
        Me.Name = "FrmMasterProspect"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Prospect Data"
        Me.gboDetail.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        CType(Me.dgvAddress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        CType(Me.dgvCP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboTax.ResumeLayout(False)
        Me.gboTax.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtNPWP As System.Windows.Forms.TextBox
    Friend WithEvents gboDetail As System.Windows.Forms.GroupBox
    Friend WithEvents TxtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents dgvAddress As System.Windows.Forms.DataGridView
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
    Friend WithEvents gboTax As System.Windows.Forms.GroupBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents TxtTaxCode As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents TxtAddressPajak1 As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents TxtAddressPajak2 As System.Windows.Forms.TextBox
    Friend WithEvents TxtZipCode As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TxtCountry As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtCity As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tslCreatedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslCreatedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedBy As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslModifiedDate As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnFind As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
End Class
