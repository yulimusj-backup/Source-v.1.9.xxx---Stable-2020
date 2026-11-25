<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Data_Transfer1
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.clbBranch = New System.Windows.Forms.CheckedListBox
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.dtpTo = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.prb = New System.Windows.Forms.ProgressBar
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.cboExpiry = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Label3 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo0 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustPONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStockOutNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLotNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocationCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyPerMaster = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarksItem = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryMasterUnits = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySecUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(463, 333)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.btnCancel)
        Me.TabPage1.Controls.Add(Me.btnOK)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.prb)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(455, 307)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "General"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(372, 35)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(372, 6)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Controls.Add(Me.clbBranch)
        Me.GroupBox1.Controls.Add(Me.LinkLabel2)
        Me.GroupBox1.Controls.Add(Me.LinkLabel1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 84)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(439, 186)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Branch"
        '
        'clbBranch
        '
        Me.clbBranch.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbBranch.FormattingEnabled = True
        Me.clbBranch.Location = New System.Drawing.Point(3, 16)
        Me.clbBranch.Name = "clbBranch"
        Me.clbBranch.Size = New System.Drawing.Size(433, 154)
        Me.clbBranch.TabIndex = 2
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.BackColor = System.Drawing.Color.White
        Me.LinkLabel2.Location = New System.Drawing.Point(305, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel2.TabIndex = 1
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Check All"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.BackColor = System.Drawing.Color.White
        Me.LinkLabel1.Location = New System.Drawing.Point(363, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel1.TabIndex = 0
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Uncheck All"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.dtpFrom)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(358, 75)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Session Period"
        '
        'dtpTo
        '
        Me.dtpTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpTo.CustomFormat = "dd MMM yyyy"
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(69, 45)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(283, 20)
        Me.dtpTo.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'dtpFrom
        '
        Me.dtpFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(69, 19)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(283, 20)
        Me.dtpFrom.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "From"
        '
        'prb
        '
        Me.prb.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prb.Location = New System.Drawing.Point(8, 276)
        Me.prb.Name = "prb"
        Me.prb.Size = New System.Drawing.Size(439, 23)
        Me.prb.TabIndex = 11
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox3)
        Me.TabPage2.Controls.Add(Me.RadioButton3)
        Me.TabPage2.Controls.Add(Me.RadioButton1)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(455, 307)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Settings"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.ComboBox2)
        Me.GroupBox3.Controls.Add(Me.cboExpiry)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 70)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(439, 118)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 94)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(233, 13)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "4. Create in adjustment to continue data transfer"
        '
        'ComboBox2
        '
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"ascending order", "descending order"})
        Me.ComboBox2.Location = New System.Drawing.Point(157, 65)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox2.TabIndex = 4
        '
        'cboExpiry
        '
        Me.cboExpiry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExpiry.FormattingEnabled = True
        Me.cboExpiry.Items.AddRange(New Object() {"ascending order", "descending order"})
        Me.cboExpiry.Location = New System.Drawing.Point(146, 38)
        Me.cboExpiry.Name = "cboExpiry"
        Me.cboExpiry.Size = New System.Drawing.Size(121, 21)
        Me.cboExpiry.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 41)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(134, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "2. Choose expiry date with "
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(145, 13)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "3. Choose batch number with"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(236, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "1. Choose location based on branch master data"
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(11, 47)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(169, 17)
        Me.RadioButton3.TabIndex = 3
        Me.RadioButton3.Text = "Automatically resolve problems"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(11, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(342, 17)
        Me.RadioButton1.TabIndex = 1
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Stop transfer, show me what's wrong, and let me do data correction"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(283, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "When there is no stock on inventory to apply data transfer:"
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
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo0, Me.ColCodeNo, Me.ColDescription, Me.ColCustCodeNo, Me.ColCustDesc, Me.ColSONO, Me.ColCustPONo, Me.ColDONO, Me.ColStockOutNo, Me.ColLotNumber, Me.ColExpiry, Me.ColLocationCode, Me.ColQtyMaster, Me.ColQtyPerMaster, Me.ColMasterUnits, Me.ColQty, Me.ColTotalQty, Me.ColUnit, Me.ColSecQty, Me.ColSecUnit, Me.ColCurrency, Me.ColUnitPrice, Me.ColDiscount, Me.ColNettPrice, Me.ColAmount, Me.ColRemarksItem, Me.ColPrimaryCodeNo, Me.ColPrimarySO, Me.ColPrimaryLocation, Me.ColPrimaryDO, Me.ColPrimaryMasterUnits, Me.ColPrimaryUnit, Me.ColPrimarySecUnit, Me.ColPrimaryCurrency})
        Me.dgvItem.Location = New System.Drawing.Point(-210, -55)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvItem.Size = New System.Drawing.Size(858, 296)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 4
        Me.dgvItem.TabStop = False
        Me.dgvItem.Visible = False
        '
        'ColNo0
        '
        Me.ColNo0.Frozen = True
        Me.ColNo0.HeaderText = "No"
        Me.ColNo0.Name = "ColNo0"
        Me.ColNo0.ReadOnly = True
        Me.ColNo0.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo0.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.Frozen = True
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.Width = 74
        '
        'ColDescription
        '
        Me.ColDescription.Frozen = True
        Me.ColDescription.HeaderText = "Description"
        Me.ColDescription.Name = "ColDescription"
        Me.ColDescription.ReadOnly = True
        Me.ColDescription.Width = 85
        '
        'ColCustCodeNo
        '
        Me.ColCustCodeNo.HeaderText = "Customer's Code No"
        Me.ColCustCodeNo.Name = "ColCustCodeNo"
        Me.ColCustCodeNo.ReadOnly = True
        Me.ColCustCodeNo.Visible = False
        Me.ColCustCodeNo.Width = 128
        '
        'ColCustDesc
        '
        Me.ColCustDesc.HeaderText = "Customer's Item Description"
        Me.ColCustDesc.Name = "ColCustDesc"
        Me.ColCustDesc.ReadOnly = True
        Me.ColCustDesc.Visible = False
        Me.ColCustDesc.Width = 162
        '
        'ColSONO
        '
        Me.ColSONO.HeaderText = "SO. No"
        Me.ColSONO.Name = "ColSONO"
        Me.ColSONO.ReadOnly = True
        Me.ColSONO.Visible = False
        Me.ColSONO.Width = 67
        '
        'ColCustPONo
        '
        Me.ColCustPONo.HeaderText = "Customer's PO No"
        Me.ColCustPONo.Name = "ColCustPONo"
        Me.ColCustPONo.ReadOnly = True
        Me.ColCustPONo.Visible = False
        Me.ColCustPONo.Width = 118
        '
        'ColDONO
        '
        Me.ColDONO.HeaderText = "DO No"
        Me.ColDONO.Name = "ColDONO"
        Me.ColDONO.ReadOnly = True
        Me.ColDONO.Visible = False
        Me.ColDONO.Width = 65
        '
        'ColStockOutNo
        '
        Me.ColStockOutNo.HeaderText = "Stock Out No"
        Me.ColStockOutNo.Name = "ColStockOutNo"
        Me.ColStockOutNo.ReadOnly = True
        Me.ColStockOutNo.Visible = False
        Me.ColStockOutNo.Width = 97
        '
        'ColLotNumber
        '
        Me.ColLotNumber.HeaderText = "Batch Number"
        Me.ColLotNumber.Name = "ColLotNumber"
        Me.ColLotNumber.ReadOnly = True
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry Date"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.ReadOnly = True
        Me.ColExpiry.Width = 86
        '
        'ColLocationCode
        '
        Me.ColLocationCode.HeaderText = "Location"
        Me.ColLocationCode.Name = "ColLocationCode"
        Me.ColLocationCode.ReadOnly = True
        Me.ColLocationCode.Width = 73
        '
        'ColQtyMaster
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMaster.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMaster.HeaderText = "Qty Master"
        Me.ColQtyMaster.Name = "ColQtyMaster"
        Me.ColQtyMaster.ReadOnly = True
        Me.ColQtyMaster.Visible = False
        Me.ColQtyMaster.Width = 83
        '
        'ColQtyPerMaster
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyPerMaster.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQtyPerMaster.HeaderText = "Qty Per Master"
        Me.ColQtyPerMaster.Name = "ColQtyPerMaster"
        Me.ColQtyPerMaster.ReadOnly = True
        Me.ColQtyPerMaster.Visible = False
        Me.ColQtyPerMaster.Width = 102
        '
        'ColMasterUnits
        '
        Me.ColMasterUnits.HeaderText = "Master Unit"
        Me.ColMasterUnits.Name = "ColMasterUnits"
        Me.ColMasterUnits.ReadOnly = True
        Me.ColMasterUnits.Visible = False
        Me.ColMasterUnits.Width = 86
        '
        'ColQty
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.Visible = False
        Me.ColQty.Width = 48
        '
        'ColTotalQty
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalQty.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColTotalQty.HeaderText = "Total Qty"
        Me.ColTotalQty.Name = "ColTotalQty"
        Me.ColTotalQty.ReadOnly = True
        Me.ColTotalQty.Width = 75
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.Width = 51
        '
        'ColSecQty
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSecQty.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColSecQty.HeaderText = "@"
        Me.ColSecQty.Name = "ColSecQty"
        Me.ColSecQty.ReadOnly = True
        Me.ColSecQty.Visible = False
        Me.ColSecQty.Width = 43
        '
        'ColSecUnit
        '
        Me.ColSecUnit.HeaderText = "@ Unit"
        Me.ColSecUnit.Name = "ColSecUnit"
        Me.ColSecUnit.ReadOnly = True
        Me.ColSecUnit.Visible = False
        Me.ColSecUnit.Width = 65
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.Name = "ColCurrency"
        Me.ColCurrency.ReadOnly = True
        Me.ColCurrency.Width = 74
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.Width = 78
        '
        'ColDiscount
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.ReadOnly = True
        Me.ColDiscount.Width = 91
        '
        'ColNettPrice
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.Width = 79
        '
        'ColAmount
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.Width = 68
        '
        'ColRemarksItem
        '
        Me.ColRemarksItem.HeaderText = "Remarks"
        Me.ColRemarksItem.Name = "ColRemarksItem"
        Me.ColRemarksItem.ReadOnly = True
        Me.ColRemarksItem.Width = 74
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "PrimaryCodeNo"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 105
        '
        'ColPrimarySO
        '
        Me.ColPrimarySO.HeaderText = "Primary SO"
        Me.ColPrimarySO.Name = "ColPrimarySO"
        Me.ColPrimarySO.ReadOnly = True
        Me.ColPrimarySO.Visible = False
        Me.ColPrimarySO.Width = 84
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "ColPrimaryLocation"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.ReadOnly = True
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 122
        '
        'ColPrimaryDO
        '
        Me.ColPrimaryDO.HeaderText = "PrimaryDO"
        Me.ColPrimaryDO.Name = "ColPrimaryDO"
        Me.ColPrimaryDO.ReadOnly = True
        Me.ColPrimaryDO.Visible = False
        Me.ColPrimaryDO.Width = 82
        '
        'ColPrimaryMasterUnits
        '
        Me.ColPrimaryMasterUnits.HeaderText = "Primary Master Unit"
        Me.ColPrimaryMasterUnits.Name = "ColPrimaryMasterUnits"
        Me.ColPrimaryMasterUnits.ReadOnly = True
        Me.ColPrimaryMasterUnits.Visible = False
        Me.ColPrimaryMasterUnits.Width = 123
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Units"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.ReadOnly = True
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 93
        '
        'ColPrimarySecUnit
        '
        Me.ColPrimarySecUnit.HeaderText = "Primary Sec Unit"
        Me.ColPrimarySecUnit.Name = "ColPrimarySecUnit"
        Me.ColPrimarySecUnit.ReadOnly = True
        Me.ColPrimarySecUnit.Visible = False
        Me.ColPrimarySecUnit.Width = 110
        '
        'ColPrimaryCurrency
        '
        Me.ColPrimaryCurrency.HeaderText = "PrimaryCurrency"
        Me.ColPrimaryCurrency.Name = "ColPrimaryCurrency"
        Me.ColPrimaryCurrency.ReadOnly = True
        Me.ColPrimaryCurrency.Visible = False
        Me.ColPrimaryCurrency.Width = 108
        '
        'FrmTrans_Data_Transfer1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(463, 333)
        Me.Controls.Add(Me.TabControl1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(383, 277)
        Me.Name = "FrmTrans_Data_Transfer1"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Start Session Data Transfer"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents clbBranch As System.Windows.Forms.CheckedListBox
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents prb As System.Windows.Forms.ProgressBar
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents cboExpiry As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Public WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo0 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustPONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStockOutNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLotNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocationCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyPerMaster As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarksItem As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryMasterUnits As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySecUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
