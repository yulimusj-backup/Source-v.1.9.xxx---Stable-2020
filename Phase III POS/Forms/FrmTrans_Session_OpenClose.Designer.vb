<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Session_OpenClose
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
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.TxtPrimarySession = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.TxtIdNumber = New System.Windows.Forms.TextBox
        Me.TxtPrimaryMember = New System.Windows.Forms.TextBox
        Me.TxtMember = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.btnPayment = New System.Windows.Forms.Button
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDisc_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSubTotal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAutoNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.TxtDisplay = New System.Windows.Forms.TextBox
        Me.BtnOK = New System.Windows.Forms.Button
        Me.btnVoid = New System.Windows.Forms.Button
        Me.btnVoidAll = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.btnVoidAll)
        Me.GroupBox1.Controls.Add(Me.btnVoid)
        Me.GroupBox1.Controls.Add(Me.SplitContainer1)
        Me.GroupBox1.Controls.Add(Me.ListBox1)
        Me.GroupBox1.Controls.Add(Me.btnPayment)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(724, 137)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 97)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimarySession)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimary)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtIdNumber)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtPrimaryMember)
        Me.SplitContainer1.Panel2.Controls.Add(Me.TxtMember)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
        Me.SplitContainer1.Size = New System.Drawing.Size(712, 36)
        Me.SplitContainer1.SplitterDistance = 353
        Me.SplitContainer1.TabIndex = 1
        '
        'TxtPrimarySession
        '
        Me.TxtPrimarySession.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimarySession.Location = New System.Drawing.Point(51, 9)
        Me.TxtPrimarySession.Name = "TxtPrimarySession"
        Me.TxtPrimarySession.Size = New System.Drawing.Size(30, 20)
        Me.TxtPrimarySession.TabIndex = 20
        Me.TxtPrimarySession.TabStop = False
        Me.TxtPrimarySession.Text = "-1"
        Me.TxtPrimarySession.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 13)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "ID Number"
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimary.Location = New System.Drawing.Point(87, 9)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimary.TabIndex = 19
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Text = "-1"
        Me.TxtPrimary.Visible = False
        '
        'TxtIdNumber
        '
        Me.TxtIdNumber.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtIdNumber.Location = New System.Drawing.Point(87, 9)
        Me.TxtIdNumber.Name = "TxtIdNumber"
        Me.TxtIdNumber.ReadOnly = True
        Me.TxtIdNumber.Size = New System.Drawing.Size(263, 20)
        Me.TxtIdNumber.TabIndex = 0
        Me.TxtIdNumber.TabStop = False
        '
        'TxtPrimaryMember
        '
        Me.TxtPrimaryMember.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtPrimaryMember.Location = New System.Drawing.Point(54, 9)
        Me.TxtPrimaryMember.Name = "TxtPrimaryMember"
        Me.TxtPrimaryMember.Size = New System.Drawing.Size(32, 20)
        Me.TxtPrimaryMember.TabIndex = 19
        Me.TxtPrimaryMember.TabStop = False
        Me.TxtPrimaryMember.Text = "-1"
        Me.TxtPrimaryMember.Visible = False
        '
        'TxtMember
        '
        Me.TxtMember.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtMember.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtMember.Location = New System.Drawing.Point(92, 8)
        Me.TxtMember.Name = "TxtMember"
        Me.TxtMember.ReadOnly = True
        Me.TxtMember.Size = New System.Drawing.Size(260, 20)
        Me.TxtMember.TabIndex = 0
        Me.TxtMember.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Member"
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(6, 9)
        Me.ListBox1.MultiColumn = True
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(636, 82)
        Me.ListBox1.TabIndex = 0
        '
        'btnPayment
        '
        Me.btnPayment.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPayment.Location = New System.Drawing.Point(648, 9)
        Me.btnPayment.Name = "btnPayment"
        Me.btnPayment.Size = New System.Drawing.Size(69, 23)
        Me.btnPayment.TabIndex = 2
        Me.btnPayment.TabStop = False
        Me.btnPayment.Text = "Pa&yment"
        Me.btnPayment.UseVisualStyleBackColor = True
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtReadOnly.Location = New System.Drawing.Point(596, 454)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(43, 20)
        Me.TxtReadOnly.TabIndex = 26
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDesc, Me.ColLocation, Me.ColBatchNumber, Me.ColExpiry, Me.ColQty, Me.ColUnit, Me.ColUnitPrice, Me.ColDiscString, Me.ColDisc_Amount, Me.ColNettPrice, Me.ColSubTotal, Me.ColPrimaryCodeNo, Me.ColPrimaryUnit, Me.ColPrimaryDisc1, Me.ColDiscAmount1, Me.ColPrimaryDisc2, Me.ColDiscAmount2, Me.ColPrimaryDisc3, Me.ColDiscAmount3, Me.ColPrimaryLocation, Me.ColDiscString1, Me.ColDiscString2, Me.ColDiscString3, Me.ColPrimaryAutoNumber})
        Me.dgvItem.Location = New System.Drawing.Point(6, 10)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(712, 188)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 55
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.ReadOnly = True
        Me.ColLocation.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocation.Visible = False
        Me.ColLocation.Width = 54
        '
        'ColBatchNumber
        '
        Me.ColBatchNumber.HeaderText = "Lot No"
        Me.ColBatchNumber.Name = "ColBatchNumber"
        Me.ColBatchNumber.ReadOnly = True
        Me.ColBatchNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBatchNumber.Visible = False
        Me.ColBatchNumber.Width = 45
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.ReadOnly = True
        Me.ColExpiry.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExpiry.Visible = False
        Me.ColExpiry.Width = 41
        '
        'ColQty
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 29
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 32
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 59
        '
        'ColDiscString
        '
        Me.ColDiscString.HeaderText = "Discount String"
        Me.ColDiscString.Name = "ColDiscString"
        Me.ColDiscString.ReadOnly = True
        Me.ColDiscString.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscString.Visible = False
        Me.ColDiscString.Width = 85
        '
        'ColDisc_Amount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDisc_Amount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDisc_Amount.HeaderText = "Discount"
        Me.ColDisc_Amount.Name = "ColDisc_Amount"
        Me.ColDisc_Amount.ReadOnly = True
        Me.ColDisc_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDisc_Amount.Width = 55
        '
        'ColNettPrice
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNettPrice.Width = 60
        '
        'ColSubTotal
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSubTotal.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColSubTotal.HeaderText = "Sub Total"
        Me.ColSubTotal.Name = "ColSubTotal"
        Me.ColSubTotal.ReadOnly = True
        Me.ColSubTotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSubTotal.Width = 59
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 111
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Unit"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.ReadOnly = True
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 88
        '
        'ColPrimaryDisc1
        '
        Me.ColPrimaryDisc1.HeaderText = "PrimaryDisc1"
        Me.ColPrimaryDisc1.Name = "ColPrimaryDisc1"
        Me.ColPrimaryDisc1.ReadOnly = True
        Me.ColPrimaryDisc1.Visible = False
        Me.ColPrimaryDisc1.Width = 93
        '
        'ColDiscAmount1
        '
        Me.ColDiscAmount1.HeaderText = "DiscAmount1"
        Me.ColDiscAmount1.Name = "ColDiscAmount1"
        Me.ColDiscAmount1.ReadOnly = True
        Me.ColDiscAmount1.Visible = False
        Me.ColDiscAmount1.Width = 95
        '
        'ColPrimaryDisc2
        '
        Me.ColPrimaryDisc2.HeaderText = "Primary Disc 2"
        Me.ColPrimaryDisc2.Name = "ColPrimaryDisc2"
        Me.ColPrimaryDisc2.ReadOnly = True
        Me.ColPrimaryDisc2.Visible = False
        Me.ColPrimaryDisc2.Width = 99
        '
        'ColDiscAmount2
        '
        Me.ColDiscAmount2.HeaderText = "Disc Amount2"
        Me.ColDiscAmount2.Name = "ColDiscAmount2"
        Me.ColDiscAmount2.ReadOnly = True
        Me.ColDiscAmount2.Visible = False
        Me.ColDiscAmount2.Width = 98
        '
        'ColPrimaryDisc3
        '
        Me.ColPrimaryDisc3.HeaderText = "Primary Disc 3"
        Me.ColPrimaryDisc3.Name = "ColPrimaryDisc3"
        Me.ColPrimaryDisc3.ReadOnly = True
        Me.ColPrimaryDisc3.Visible = False
        Me.ColPrimaryDisc3.Width = 99
        '
        'ColDiscAmount3
        '
        Me.ColDiscAmount3.HeaderText = "Disc Amount 3"
        Me.ColDiscAmount3.Name = "ColDiscAmount3"
        Me.ColDiscAmount3.ReadOnly = True
        Me.ColDiscAmount3.Visible = False
        Me.ColDiscAmount3.Width = 101
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.ReadOnly = True
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 110
        '
        'ColDiscString1
        '
        Me.ColDiscString1.HeaderText = "DiscString1"
        Me.ColDiscString1.Name = "ColDiscString1"
        Me.ColDiscString1.ReadOnly = True
        Me.ColDiscString1.Visible = False
        Me.ColDiscString1.Width = 86
        '
        'ColDiscString2
        '
        Me.ColDiscString2.HeaderText = "DiscString2"
        Me.ColDiscString2.Name = "ColDiscString2"
        Me.ColDiscString2.ReadOnly = True
        Me.ColDiscString2.Visible = False
        Me.ColDiscString2.Width = 86
        '
        'ColDiscString3
        '
        Me.ColDiscString3.HeaderText = "DiscString3"
        Me.ColDiscString3.Name = "ColDiscString3"
        Me.ColDiscString3.ReadOnly = True
        Me.ColDiscString3.Visible = False
        Me.ColDiscString3.Width = 86
        '
        'ColPrimaryAutoNumber
        '
        Me.ColPrimaryAutoNumber.HeaderText = "PrimaryAutoNumber"
        Me.ColPrimaryAutoNumber.Name = "ColPrimaryAutoNumber"
        Me.ColPrimaryAutoNumber.ReadOnly = True
        Me.ColPrimaryAutoNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAutoNumber.Visible = False
        Me.ColPrimaryAutoNumber.Width = 106
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(660, 414)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(69, 58)
        Me.btnClose.TabIndex = 4
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.dgvItem)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 147)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(724, 204)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        '
        'TxtDisplay
        '
        Me.TxtDisplay.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!)
        Me.TxtDisplay.ForeColor = System.Drawing.Color.Blue
        Me.TxtDisplay.Location = New System.Drawing.Point(12, 357)
        Me.TxtDisplay.Multiline = True
        Me.TxtDisplay.Name = "TxtDisplay"
        Me.TxtDisplay.Size = New System.Drawing.Size(642, 115)
        Me.TxtDisplay.TabIndex = 2
        Me.TxtDisplay.TabStop = False
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(660, 357)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(69, 51)
        Me.BtnOK.TabIndex = 3
        Me.BtnOK.TabStop = False
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'btnVoid
        '
        Me.btnVoid.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnVoid.Location = New System.Drawing.Point(648, 38)
        Me.btnVoid.Name = "btnVoid"
        Me.btnVoid.Size = New System.Drawing.Size(69, 23)
        Me.btnVoid.TabIndex = 3
        Me.btnVoid.TabStop = False
        Me.btnVoid.Text = "Void"
        Me.btnVoid.UseVisualStyleBackColor = True
        '
        'btnVoidAll
        '
        Me.btnVoidAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnVoidAll.Location = New System.Drawing.Point(648, 67)
        Me.btnVoidAll.Name = "btnVoidAll"
        Me.btnVoidAll.Size = New System.Drawing.Size(69, 23)
        Me.btnVoidAll.TabIndex = 4
        Me.btnVoidAll.TabStop = False
        Me.btnVoidAll.Text = "Void All"
        Me.btnVoidAll.UseVisualStyleBackColor = True
        '
        'FrmTrans_Session_OpenClose
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(741, 506)
        Me.ControlBox = False
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.TxtDisplay)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(757, 522)
        Me.Name = "FrmTrans_Session_OpenClose"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transaction POS Paid"
        Me.GroupBox1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDisc_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSubTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAutoNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TxtPrimarySession As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents TxtIdNumber As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryMember As System.Windows.Forms.TextBox
    Friend WithEvents TxtMember As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtDisplay As System.Windows.Forms.TextBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Public WithEvents btnPayment As System.Windows.Forms.Button
    Public WithEvents btnVoid As System.Windows.Forms.Button
    Public WithEvents btnVoidAll As System.Windows.Forms.Button
End Class
