<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMasterProduct_PriceList
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
        Me.BtnFilter = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnPrintCust = New System.Windows.Forms.Button
        Me.BtnPrintCusGrouop = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.cboCustomerGroup = New System.Windows.Forms.ComboBox
        Me.TxtKodeProduk = New System.Windows.Forms.TextBox
        Me.btnDelete_ByCustomer = New System.Windows.Forms.Button
        Me.btnDelete_ByGroup = New System.Windows.Forms.Button
        Me.ChkUpdateCust = New System.Windows.Forms.CheckBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.ChkUpdateCustGroup = New System.Windows.Forms.CheckBox
        Me.txtprimaryCust = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCustGroup = New System.Windows.Forms.TextBox
        Me.TxtIndexCustType = New System.Windows.Forms.TextBox
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.TxtCustomerCode = New System.Windows.Forms.TextBox
        Me.rboCustomer = New System.Windows.Forms.RadioButton
        Me.cboCustomerCode = New System.Windows.Forms.ComboBox
        Me.rboCustomerGroup = New System.Windows.Forms.RadioButton
        Me.rboGeneral = New System.Windows.Forms.RadioButton
        Me.txtPrimaryCurrency = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.CboCurrency = New System.Windows.Forms.ComboBox
        Me.cmdKeluar = New System.Windows.Forms.Button
        Me.DGVHarga = New System.Windows.Forms.DataGridView
        Me.ColProductCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMaxQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefrec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnPrint = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ChkGeneralAll = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.DGVHarga, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnFilter
        '
        Me.BtnFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnFilter.Location = New System.Drawing.Point(853, 14)
        Me.BtnFilter.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnFilter.Name = "BtnFilter"
        Me.BtnFilter.Size = New System.Drawing.Size(91, 32)
        Me.BtnFilter.TabIndex = 2
        Me.BtnFilter.TabStop = False
        Me.BtnFilter.Text = "Filter"
        Me.BtnFilter.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.BtnPrintCust)
        Me.GroupBox1.Controls.Add(Me.BtnPrintCusGrouop)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.cboCustomerGroup)
        Me.GroupBox1.Controls.Add(Me.TxtKodeProduk)
        Me.GroupBox1.Controls.Add(Me.btnDelete_ByCustomer)
        Me.GroupBox1.Controls.Add(Me.btnDelete_ByGroup)
        Me.GroupBox1.Controls.Add(Me.ChkUpdateCust)
        Me.GroupBox1.Controls.Add(Me.TxtReadOnly)
        Me.GroupBox1.Controls.Add(Me.ChkUpdateCustGroup)
        Me.GroupBox1.Controls.Add(Me.txtprimaryCust)
        Me.GroupBox1.Controls.Add(Me.TxtPrimaryCustGroup)
        Me.GroupBox1.Controls.Add(Me.TxtIndexCustType)
        Me.GroupBox1.Controls.Add(Me.TxtCustomerName)
        Me.GroupBox1.Controls.Add(Me.btnFindCustomer)
        Me.GroupBox1.Controls.Add(Me.TxtCustomerCode)
        Me.GroupBox1.Controls.Add(Me.rboCustomer)
        Me.GroupBox1.Controls.Add(Me.cboCustomerCode)
        Me.GroupBox1.Controls.Add(Me.rboCustomerGroup)
        Me.GroupBox1.Controls.Add(Me.rboGeneral)
        Me.GroupBox1.Controls.Add(Me.txtPrimaryCurrency)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.CboCurrency)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 7)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(838, 159)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Filter"
        '
        'BtnPrintCust
        '
        Me.BtnPrintCust.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPrintCust.Location = New System.Drawing.Point(776, 129)
        Me.BtnPrintCust.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnPrintCust.Name = "BtnPrintCust"
        Me.BtnPrintCust.Size = New System.Drawing.Size(51, 25)
        Me.BtnPrintCust.TabIndex = 229
        Me.BtnPrintCust.TabStop = False
        Me.BtnPrintCust.Text = "Print"
        Me.BtnPrintCust.UseVisualStyleBackColor = True
        '
        'BtnPrintCusGrouop
        '
        Me.BtnPrintCusGrouop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnPrintCusGrouop.Location = New System.Drawing.Point(776, 101)
        Me.BtnPrintCusGrouop.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.BtnPrintCusGrouop.Name = "BtnPrintCusGrouop"
        Me.BtnPrintCusGrouop.Size = New System.Drawing.Size(51, 25)
        Me.BtnPrintCusGrouop.TabIndex = 5
        Me.BtnPrintCusGrouop.TabStop = False
        Me.BtnPrintCusGrouop.Text = "Print"
        Me.BtnPrintCusGrouop.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Location = New System.Drawing.Point(685, 129)
        Me.Button2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(85, 25)
        Me.Button2.TabIndex = 228
        Me.Button2.TabStop = False
        Me.Button2.Text = "Delete All"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(685, 101)
        Me.Button1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(85, 25)
        Me.Button1.TabIndex = 227
        Me.Button1.TabStop = False
        Me.Button1.Text = "Delete All"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'cboCustomerGroup
        '
        Me.cboCustomerGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerGroup.FormattingEnabled = True
        Me.cboCustomerGroup.Items.AddRange(New Object() {"Customer Group", "Customer Sales Area", "Customer Segmentation", "Customer Type"})
        Me.cboCustomerGroup.Location = New System.Drawing.Point(149, 102)
        Me.cboCustomerGroup.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboCustomerGroup.Name = "cboCustomerGroup"
        Me.cboCustomerGroup.Size = New System.Drawing.Size(208, 24)
        Me.cboCustomerGroup.TabIndex = 6
        '
        'TxtKodeProduk
        '
        Me.TxtKodeProduk.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtKodeProduk.BackColor = System.Drawing.SystemColors.Window
        Me.TxtKodeProduk.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtKodeProduk.Location = New System.Drawing.Point(149, 18)
        Me.TxtKodeProduk.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtKodeProduk.Name = "TxtKodeProduk"
        Me.TxtKodeProduk.Size = New System.Drawing.Size(610, 22)
        Me.TxtKodeProduk.TabIndex = 0
        '
        'btnDelete_ByCustomer
        '
        Me.btnDelete_ByCustomer.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete_ByCustomer.Location = New System.Drawing.Point(602, 130)
        Me.btnDelete_ByCustomer.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnDelete_ByCustomer.Name = "btnDelete_ByCustomer"
        Me.btnDelete_ByCustomer.Size = New System.Drawing.Size(76, 25)
        Me.btnDelete_ByCustomer.TabIndex = 13
        Me.btnDelete_ByCustomer.TabStop = False
        Me.btnDelete_ByCustomer.Text = "Delete"
        Me.btnDelete_ByCustomer.UseVisualStyleBackColor = True
        '
        'btnDelete_ByGroup
        '
        Me.btnDelete_ByGroup.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete_ByGroup.Location = New System.Drawing.Point(602, 101)
        Me.btnDelete_ByGroup.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnDelete_ByGroup.Name = "btnDelete_ByGroup"
        Me.btnDelete_ByGroup.Size = New System.Drawing.Size(76, 25)
        Me.btnDelete_ByGroup.TabIndex = 8
        Me.btnDelete_ByGroup.TabStop = False
        Me.btnDelete_ByGroup.Text = "Delete"
        Me.btnDelete_ByGroup.UseVisualStyleBackColor = True
        '
        'ChkUpdateCust
        '
        Me.ChkUpdateCust.AutoSize = True
        Me.ChkUpdateCust.Location = New System.Drawing.Point(411, 78)
        Me.ChkUpdateCust.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkUpdateCust.Name = "ChkUpdateCust"
        Me.ChkUpdateCust.Size = New System.Drawing.Size(215, 21)
        Me.ChkUpdateCust.TabIndex = 4
        Me.ChkUpdateCust.Text = "Update price list for customer"
        Me.ChkUpdateCust.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(115, 18)
        Me.TxtReadOnly.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.TxtReadOnly.Size = New System.Drawing.Size(28, 22)
        Me.TxtReadOnly.TabIndex = 220
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'ChkUpdateCustGroup
        '
        Me.ChkUpdateCustGroup.AutoSize = True
        Me.ChkUpdateCustGroup.Location = New System.Drawing.Point(149, 78)
        Me.ChkUpdateCustGroup.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkUpdateCustGroup.Name = "ChkUpdateCustGroup"
        Me.ChkUpdateCustGroup.Size = New System.Drawing.Size(256, 21)
        Me.ChkUpdateCustGroup.TabIndex = 3
        Me.ChkUpdateCustGroup.Text = "Update price list for customer group"
        Me.ChkUpdateCustGroup.UseVisualStyleBackColor = True
        '
        'txtprimaryCust
        '
        Me.txtprimaryCust.Location = New System.Drawing.Point(116, 128)
        Me.txtprimaryCust.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtprimaryCust.Name = "txtprimaryCust"
        Me.txtprimaryCust.Size = New System.Drawing.Size(28, 22)
        Me.txtprimaryCust.TabIndex = 226
        Me.txtprimaryCust.TabStop = False
        Me.txtprimaryCust.Text = "-1"
        Me.txtprimaryCust.Visible = False
        '
        'TxtPrimaryCustGroup
        '
        Me.TxtPrimaryCustGroup.Location = New System.Drawing.Point(419, 103)
        Me.TxtPrimaryCustGroup.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtPrimaryCustGroup.Name = "TxtPrimaryCustGroup"
        Me.TxtPrimaryCustGroup.Size = New System.Drawing.Size(28, 22)
        Me.TxtPrimaryCustGroup.TabIndex = 225
        Me.TxtPrimaryCustGroup.TabStop = False
        Me.TxtPrimaryCustGroup.Text = "-1"
        Me.TxtPrimaryCustGroup.Visible = False
        '
        'TxtIndexCustType
        '
        Me.TxtIndexCustType.Location = New System.Drawing.Point(272, 105)
        Me.TxtIndexCustType.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtIndexCustType.Name = "TxtIndexCustType"
        Me.TxtIndexCustType.Size = New System.Drawing.Size(28, 22)
        Me.TxtIndexCustType.TabIndex = 224
        Me.TxtIndexCustType.TabStop = False
        Me.TxtIndexCustType.Text = "-1"
        Me.TxtIndexCustType.Visible = False
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerName.BackColor = System.Drawing.SystemColors.Window
        Me.TxtCustomerName.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerName.Enabled = False
        Me.TxtCustomerName.Location = New System.Drawing.Point(363, 130)
        Me.TxtCustomerName.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(233, 22)
        Me.TxtCustomerName.TabIndex = 12
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Location = New System.Drawing.Point(306, 130)
        Me.btnFindCustomer.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(51, 25)
        Me.btnFindCustomer.TabIndex = 11
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "Find"
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'TxtCustomerCode
        '
        Me.TxtCustomerCode.BackColor = System.Drawing.SystemColors.Window
        Me.TxtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtCustomerCode.Location = New System.Drawing.Point(149, 130)
        Me.TxtCustomerCode.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TxtCustomerCode.Name = "TxtCustomerCode"
        Me.TxtCustomerCode.Size = New System.Drawing.Size(151, 22)
        Me.TxtCustomerCode.TabIndex = 10
        '
        'rboCustomer
        '
        Me.rboCustomer.AutoSize = True
        Me.rboCustomer.Location = New System.Drawing.Point(9, 129)
        Me.rboCustomer.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rboCustomer.Name = "rboCustomer"
        Me.rboCustomer.Size = New System.Drawing.Size(89, 21)
        Me.rboCustomer.TabIndex = 9
        Me.rboCustomer.Text = "Customer"
        Me.rboCustomer.UseVisualStyleBackColor = True
        '
        'cboCustomerCode
        '
        Me.cboCustomerCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboCustomerCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCustomerCode.FormattingEnabled = True
        Me.cboCustomerCode.Location = New System.Drawing.Point(363, 102)
        Me.cboCustomerCode.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cboCustomerCode.Name = "cboCustomerCode"
        Me.cboCustomerCode.Size = New System.Drawing.Size(233, 24)
        Me.cboCustomerCode.TabIndex = 7
        '
        'rboCustomerGroup
        '
        Me.rboCustomerGroup.AutoSize = True
        Me.rboCustomerGroup.Location = New System.Drawing.Point(11, 102)
        Me.rboCustomerGroup.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rboCustomerGroup.Name = "rboCustomerGroup"
        Me.rboCustomerGroup.Size = New System.Drawing.Size(133, 21)
        Me.rboCustomerGroup.TabIndex = 5
        Me.rboCustomerGroup.Text = "Customer Group"
        Me.rboCustomerGroup.UseVisualStyleBackColor = True
        '
        'rboGeneral
        '
        Me.rboGeneral.AutoSize = True
        Me.rboGeneral.Checked = True
        Me.rboGeneral.Location = New System.Drawing.Point(11, 75)
        Me.rboGeneral.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rboGeneral.Name = "rboGeneral"
        Me.rboGeneral.Size = New System.Drawing.Size(80, 21)
        Me.rboGeneral.TabIndex = 2
        Me.rboGeneral.TabStop = True
        Me.rboGeneral.Text = "General"
        Me.rboGeneral.UseVisualStyleBackColor = True
        '
        'txtPrimaryCurrency
        '
        Me.txtPrimaryCurrency.Location = New System.Drawing.Point(117, 46)
        Me.txtPrimaryCurrency.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtPrimaryCurrency.Name = "txtPrimaryCurrency"
        Me.txtPrimaryCurrency.Size = New System.Drawing.Size(28, 22)
        Me.txtPrimaryCurrency.TabIndex = 214
        Me.txtPrimaryCurrency.TabStop = False
        Me.txtPrimaryCurrency.Text = "-1"
        Me.txtPrimaryCurrency.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 17)
        Me.Label1.TabIndex = 40
        Me.Label1.Text = "Code No"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 49)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(65, 17)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Currency"
        '
        'CboCurrency
        '
        Me.CboCurrency.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CboCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboCurrency.FormattingEnabled = True
        Me.CboCurrency.Location = New System.Drawing.Point(149, 46)
        Me.CboCurrency.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CboCurrency.Name = "CboCurrency"
        Me.CboCurrency.Size = New System.Drawing.Size(610, 24)
        Me.CboCurrency.TabIndex = 1
        '
        'cmdKeluar
        '
        Me.cmdKeluar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdKeluar.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdKeluar.Location = New System.Drawing.Point(876, 516)
        Me.cmdKeluar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmdKeluar.Name = "cmdKeluar"
        Me.cmdKeluar.Size = New System.Drawing.Size(67, 28)
        Me.cmdKeluar.TabIndex = 4
        Me.cmdKeluar.TabStop = False
        Me.cmdKeluar.Text = "&Close"
        Me.cmdKeluar.UseVisualStyleBackColor = True
        '
        'DGVHarga
        '
        Me.DGVHarga.AllowUserToAddRows = False
        Me.DGVHarga.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGVHarga.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGVHarga.ColumnHeadersHeight = 25
        Me.DGVHarga.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColProductCode, Me.ColDescription, Me.ColUnit, Me.ColUnitPrice, Me.ColDiscount, Me.ColMaxQty, Me.ColRemarks, Me.ColPrimaryUnit, Me.ColPrimaryCodeNo, Me.ColPrimary, Me.ColDefrec, Me.ColDecimal})
        Me.DGVHarga.Location = New System.Drawing.Point(9, 171)
        Me.DGVHarga.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DGVHarga.MultiSelect = False
        Me.DGVHarga.Name = "DGVHarga"
        Me.DGVHarga.RowTemplate.Height = 24
        Me.DGVHarga.Size = New System.Drawing.Size(933, 373)
        Me.DGVHarga.TabIndex = 2
        Me.DGVHarga.TabStop = False
        '
        'ColProductCode
        '
        Me.ColProductCode.HeaderText = "Code No"
        Me.ColProductCode.Name = "ColProductCode"
        Me.ColProductCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColProductCode.Width = 69
        '
        'ColDescription
        '
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDescription.Width = 85
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 39
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.MaxInputLength = 12
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 75
        '
        'ColDiscount
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColDiscount.HeaderText = "Discount"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscount.Width = 69
        '
        'ColMaxQty
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMaxQty.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColMaxQty.HeaderText = "Max Qty"
        Me.ColMaxQty.MaxInputLength = 12
        Me.ColMaxQty.Name = "ColMaxQty"
        Me.ColMaxQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColMaxQty.Width = 65
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 70
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Unit"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 110
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 132
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 81
        '
        'ColDefrec
        '
        Me.ColDefrec.HeaderText = "DefPrec"
        Me.ColDefrec.Name = "ColDefrec"
        Me.ColDefrec.Visible = False
        Me.ColDefrec.Width = 84
        '
        'ColDecimal
        '
        Me.ColDecimal.HeaderText = "Decimal"
        Me.ColDecimal.Name = "ColDecimal"
        Me.ColDecimal.Visible = False
        Me.ColDecimal.Width = 83
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Location = New System.Drawing.Point(852, 50)
        Me.btnPrint.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(91, 32)
        Me.btnPrint.TabIndex = 3
        Me.btnPrint.TabStop = False
        Me.btnPrint.Text = "Print"
        Me.btnPrint.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'ChkGeneralAll
        '
        Me.ChkGeneralAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ChkGeneralAll.Location = New System.Drawing.Point(853, 89)
        Me.ChkGeneralAll.Margin = New System.Windows.Forms.Padding(4)
        Me.ChkGeneralAll.Name = "ChkGeneralAll"
        Me.ChkGeneralAll.Size = New System.Drawing.Size(89, 21)
        Me.ChkGeneralAll.TabIndex = 5
        Me.ChkGeneralAll.Text = "Show All"
        Me.ChkGeneralAll.UseVisualStyleBackColor = True
        '
        'frmMasterProduct_PriceList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdKeluar
        Me.ClientSize = New System.Drawing.Size(952, 551)
        Me.Controls.Add(Me.DGVHarga)
        Me.Controls.Add(Me.ChkGeneralAll)
        Me.Controls.Add(Me.btnPrint)
        Me.Controls.Add(Me.BtnFilter)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdKeluar)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimumSize = New System.Drawing.Size(960, 588)
        Me.Name = "frmMasterProduct_PriceList"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Price List"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DGVHarga, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnFilter As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtPrimaryCurrency As System.Windows.Forms.TextBox
    Friend WithEvents TxtKodeProduk As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents CboCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents cmdKeluar As System.Windows.Forms.Button
    Friend WithEvents DGVHarga As System.Windows.Forms.DataGridView
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents TxtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents rboCustomer As System.Windows.Forms.RadioButton
    Friend WithEvents cboCustomerCode As System.Windows.Forms.ComboBox
    Friend WithEvents cboCustomerGroup As System.Windows.Forms.ComboBox
    Friend WithEvents rboCustomerGroup As System.Windows.Forms.RadioButton
    Friend WithEvents rboGeneral As System.Windows.Forms.RadioButton
    Friend WithEvents txtprimaryCust As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCustGroup As System.Windows.Forms.TextBox
    Friend WithEvents TxtIndexCustType As System.Windows.Forms.TextBox
    Friend WithEvents ChkUpdateCust As System.Windows.Forms.CheckBox
    Friend WithEvents ChkUpdateCustGroup As System.Windows.Forms.CheckBox
    Friend WithEvents btnDelete_ByCustomer As System.Windows.Forms.Button
    Friend WithEvents btnDelete_ByGroup As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColProductCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMaxQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefrec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents BtnPrintCust As System.Windows.Forms.Button
    Friend WithEvents BtnPrintCusGrouop As System.Windows.Forms.Button
    Friend WithEvents ChkGeneralAll As System.Windows.Forms.CheckBox
End Class
