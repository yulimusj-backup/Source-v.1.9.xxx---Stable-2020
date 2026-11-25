<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanStockInOut
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLaporanStockInOut))
        Me.CboPilihGroupLevel = New System.Windows.Forms.ComboBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.DtpMinExpiryDate = New System.Windows.Forms.DateTimePicker
        Me.DtpMaxExpiryDate = New System.Windows.Forms.DateTimePicker
        Me.ChkFilterByDate = New System.Windows.Forms.CheckBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ChkShowAOPNo = New System.Windows.Forms.CheckBox
        Me.CboInPurchase = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.LblPONumber = New System.Windows.Forms.Label
        Me.btnFindPONumber = New System.Windows.Forms.Button
        Me.TxtPrimaryPO = New System.Windows.Forms.TextBox
        Me.TxtPONumber = New System.Windows.Forms.TextBox
        Me.CboScope = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.ChkShowSubTotal = New System.Windows.Forms.CheckBox
        Me.ChkShowAllItemsInTrans = New System.Windows.Forms.CheckBox
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvLocation = New System.Windows.Forms.DataGridView
        Me.ColLocation_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.DgvGroup = New System.Windows.Forms.DataGridView
        Me.ColGroup_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGroup_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvBatchNo = New System.Windows.Forms.DataGridView
        Me.ColBatch_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.GboAdjustmentType = New System.Windows.Forms.GroupBox
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel4 = New System.Windows.Forms.LinkLabel
        Me.clbAdjustmentType = New System.Windows.Forms.CheckedListBox
        Me.dgvCustomer = New System.Windows.Forms.DataGridView
        Me.ColCustomer_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomer_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomer_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
        Me.ChkFilterByWareHouse = New System.Windows.Forms.CheckBox
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DgvGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        CType(Me.DgvBatchNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage4.SuspendLayout()
        Me.GboAdjustmentType.SuspendLayout()
        CType(Me.dgvCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CboPilihGroupLevel
        '
        Me.CboPilihGroupLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPilihGroupLevel.FormattingEnabled = True
        Me.CboPilihGroupLevel.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.CboPilihGroupLevel.Location = New System.Drawing.Point(50, 11)
        Me.CboPilihGroupLevel.Margin = New System.Windows.Forms.Padding(2)
        Me.CboPilihGroupLevel.Name = "CboPilihGroupLevel"
        Me.CboPilihGroupLevel.Size = New System.Drawing.Size(178, 21)
        Me.CboPilihGroupLevel.TabIndex = 314
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(2)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkFilterByWareHouse)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpMinExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpMaxExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkFilterByDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowAOPNo)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboInPurchase)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LblPONumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFindPONumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPrimaryPO)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TxtPONumber)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboScope)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowSubTotal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowAllItemsInTrans)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboFormat)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crv1)
        Me.SplitContainer1.Size = New System.Drawing.Size(744, 514)
        Me.SplitContainer1.SplitterDistance = 253
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'DtpMinExpiryDate
        '
        Me.DtpMinExpiryDate.CustomFormat = "dd MMM yyyy"
        Me.DtpMinExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpMinExpiryDate.Location = New System.Drawing.Point(72, 123)
        Me.DtpMinExpiryDate.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpMinExpiryDate.Name = "DtpMinExpiryDate"
        Me.DtpMinExpiryDate.Size = New System.Drawing.Size(180, 20)
        Me.DtpMinExpiryDate.TabIndex = 4
        '
        'DtpMaxExpiryDate
        '
        Me.DtpMaxExpiryDate.CustomFormat = "dd MMM yyyy"
        Me.DtpMaxExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpMaxExpiryDate.Location = New System.Drawing.Point(72, 147)
        Me.DtpMaxExpiryDate.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpMaxExpiryDate.Name = "DtpMaxExpiryDate"
        Me.DtpMaxExpiryDate.Size = New System.Drawing.Size(180, 20)
        Me.DtpMaxExpiryDate.TabIndex = 5
        '
        'ChkFilterByDate
        '
        Me.ChkFilterByDate.AutoSize = True
        Me.ChkFilterByDate.Checked = True
        Me.ChkFilterByDate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkFilterByDate.Location = New System.Drawing.Point(72, 102)
        Me.ChkFilterByDate.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkFilterByDate.Name = "ChkFilterByDate"
        Me.ChkFilterByDate.Size = New System.Drawing.Size(89, 17)
        Me.ChkFilterByDate.TabIndex = 3
        Me.ChkFilterByDate.Text = "Filter By Date"
        Me.ChkFilterByDate.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 123)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "From"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 147)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "To"
        '
        'ChkShowAOPNo
        '
        Me.ChkShowAOPNo.AutoSize = True
        Me.ChkShowAOPNo.Checked = True
        Me.ChkShowAOPNo.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkShowAOPNo.Location = New System.Drawing.Point(72, 55)
        Me.ChkShowAOPNo.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowAOPNo.Name = "ChkShowAOPNo"
        Me.ChkShowAOPNo.Size = New System.Drawing.Size(118, 17)
        Me.ChkShowAOPNo.TabIndex = 2
        Me.ChkShowAOPNo.Text = "Show AOP Number"
        Me.ChkShowAOPNo.UseVisualStyleBackColor = True
        '
        'CboInPurchase
        '
        Me.CboInPurchase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboInPurchase.FormattingEnabled = True
        Me.CboInPurchase.Items.AddRange(New Object() {"With AOP", "W/O AOP", "Both"})
        Me.CboInPurchase.Location = New System.Drawing.Point(72, 30)
        Me.CboInPurchase.Margin = New System.Windows.Forms.Padding(2)
        Me.CboInPurchase.Name = "CboInPurchase"
        Me.CboInPurchase.Size = New System.Drawing.Size(180, 21)
        Me.CboInPurchase.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(4, 33)
        Me.Label7.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 13)
        Me.Label7.TabIndex = 320
        Me.Label7.Text = "In Purchase"
        '
        'LblPONumber
        '
        Me.LblPONumber.AutoSize = True
        Me.LblPONumber.Location = New System.Drawing.Point(3, 77)
        Me.LblPONumber.Name = "LblPONumber"
        Me.LblPONumber.Size = New System.Drawing.Size(62, 13)
        Me.LblPONumber.TabIndex = 318
        Me.LblPONumber.Text = "PO Number"
        '
        'btnFindPONumber
        '
        Me.btnFindPONumber.Location = New System.Drawing.Point(205, 77)
        Me.btnFindPONumber.Name = "btnFindPONumber"
        Me.btnFindPONumber.Size = New System.Drawing.Size(47, 21)
        Me.btnFindPONumber.TabIndex = 3
        Me.btnFindPONumber.TabStop = False
        Me.btnFindPONumber.Text = "Find"
        Me.btnFindPONumber.UseVisualStyleBackColor = True
        '
        'TxtPrimaryPO
        '
        Me.TxtPrimaryPO.Location = New System.Drawing.Point(72, 77)
        Me.TxtPrimaryPO.Name = "TxtPrimaryPO"
        Me.TxtPrimaryPO.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryPO.TabIndex = 319
        Me.TxtPrimaryPO.TabStop = False
        Me.TxtPrimaryPO.Text = "-1"
        Me.TxtPrimaryPO.Visible = False
        '
        'TxtPONumber
        '
        Me.TxtPONumber.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TxtPONumber.Location = New System.Drawing.Point(72, 77)
        Me.TxtPONumber.MaxLength = 150
        Me.TxtPONumber.Name = "TxtPONumber"
        Me.TxtPONumber.Size = New System.Drawing.Size(132, 20)
        Me.TxtPONumber.TabIndex = 2
        '
        'CboScope
        '
        Me.CboScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboScope.FormattingEnabled = True
        Me.CboScope.Items.AddRange(New Object() {"Goods", "Service", "Both"})
        Me.CboScope.Location = New System.Drawing.Point(72, 171)
        Me.CboScope.Margin = New System.Windows.Forms.Padding(2)
        Me.CboScope.Name = "CboScope"
        Me.CboScope.Size = New System.Drawing.Size(180, 21)
        Me.CboScope.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 181)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 13)
        Me.Label5.TabIndex = 315
        Me.Label5.Text = "Scope"
        '
        'ChkShowSubTotal
        '
        Me.ChkShowSubTotal.AutoSize = True
        Me.ChkShowSubTotal.Checked = True
        Me.ChkShowSubTotal.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkShowSubTotal.Location = New System.Drawing.Point(72, 233)
        Me.ChkShowSubTotal.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowSubTotal.Name = "ChkShowSubTotal"
        Me.ChkShowSubTotal.Size = New System.Drawing.Size(102, 17)
        Me.ChkShowSubTotal.TabIndex = 8
        Me.ChkShowSubTotal.Text = "Show Sub Total"
        Me.ChkShowSubTotal.UseVisualStyleBackColor = True
        '
        'ChkShowAllItemsInTrans
        '
        Me.ChkShowAllItemsInTrans.AutoSize = True
        Me.ChkShowAllItemsInTrans.Checked = True
        Me.ChkShowAllItemsInTrans.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkShowAllItemsInTrans.Location = New System.Drawing.Point(72, 214)
        Me.ChkShowAllItemsInTrans.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowAllItemsInTrans.Name = "ChkShowAllItemsInTrans"
        Me.ChkShowAllItemsInTrans.Size = New System.Drawing.Size(160, 17)
        Me.ChkShowAllItemsInTrans.TabIndex = 7
        Me.ChkShowAllItemsInTrans.Text = "Show All Item in Transaction"
        Me.ChkShowAllItemsInTrans.UseVisualStyleBackColor = True
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"Stock In Purchase", "Stock In Sales Return", "Stock In Adjust", "Stock Out Sales", "Stock Out Purchase Return", "Stock Out Adjust"})
        Me.CboFormat.Location = New System.Drawing.Point(72, 4)
        Me.CboFormat.Margin = New System.Windows.Forms.Padding(2)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(180, 21)
        Me.CboFormat.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 7)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Format"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(257, 29)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(46, 20)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(257, 4)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(46, 20)
        Me.btnFilter.TabIndex = 9
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(308, 4)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(436, 246)
        Me.TabControl1.TabIndex = 11
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvLocation)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(428, 220)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Location"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'DgvLocation
        '
        Me.DgvLocation.AllowUserToAddRows = False
        Me.DgvLocation.AllowUserToResizeColumns = False
        Me.DgvLocation.AllowUserToResizeRows = False
        Me.DgvLocation.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvLocation.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColLocation_Code, Me.ColLocation_Desc, Me.ColLocation_Primary})
        Me.DgvLocation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvLocation.Location = New System.Drawing.Point(2, 2)
        Me.DgvLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvLocation.MultiSelect = False
        Me.DgvLocation.Name = "DgvLocation"
        Me.DgvLocation.RowTemplate.Height = 24
        Me.DgvLocation.Size = New System.Drawing.Size(424, 216)
        Me.DgvLocation.TabIndex = 0
        '
        'ColLocation_Code
        '
        Me.ColLocation_Code.HeaderText = "Code"
        Me.ColLocation_Code.MaxInputLength = 250
        Me.ColLocation_Code.MinimumWidth = 100
        Me.ColLocation_Code.Name = "ColLocation_Code"
        '
        'ColLocation_Desc
        '
        Me.ColLocation_Desc.HeaderText = "Description"
        Me.ColLocation_Desc.MaxInputLength = 250
        Me.ColLocation_Desc.MinimumWidth = 150
        Me.ColLocation_Desc.Name = "ColLocation_Desc"
        Me.ColLocation_Desc.ReadOnly = True
        Me.ColLocation_Desc.Width = 150
        '
        'ColLocation_Primary
        '
        Me.ColLocation_Primary.HeaderText = "Primary"
        Me.ColLocation_Primary.Name = "ColLocation_Primary"
        Me.ColLocation_Primary.Visible = False
        Me.ColLocation_Primary.Width = 66
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(428, 220)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Product Group"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CboPilihGroupLevel)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.DgvGroup)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(2, 2)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox1.Size = New System.Drawing.Size(424, 216)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 13)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(33, 13)
        Me.Label3.TabIndex = 313
        Me.Label3.Text = "Level"
        '
        'DgvGroup
        '
        Me.DgvGroup.AllowUserToAddRows = False
        Me.DgvGroup.AllowUserToResizeColumns = False
        Me.DgvGroup.AllowUserToResizeRows = False
        Me.DgvGroup.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvGroup.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvGroup.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvGroup.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColGroup_Code, Me.ColGroup_Desc, Me.ColGroup_Primary})
        Me.DgvGroup.Location = New System.Drawing.Point(2, 36)
        Me.DgvGroup.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvGroup.MultiSelect = False
        Me.DgvGroup.Name = "DgvGroup"
        Me.DgvGroup.RowTemplate.Height = 24
        Me.DgvGroup.Size = New System.Drawing.Size(420, 175)
        Me.DgvGroup.TabIndex = 0
        '
        'ColGroup_Code
        '
        Me.ColGroup_Code.HeaderText = "Code"
        Me.ColGroup_Code.MaxInputLength = 250
        Me.ColGroup_Code.MinimumWidth = 100
        Me.ColGroup_Code.Name = "ColGroup_Code"
        '
        'ColGroup_Desc
        '
        Me.ColGroup_Desc.HeaderText = "Description"
        Me.ColGroup_Desc.MaxInputLength = 250
        Me.ColGroup_Desc.MinimumWidth = 150
        Me.ColGroup_Desc.Name = "ColGroup_Desc"
        Me.ColGroup_Desc.ReadOnly = True
        Me.ColGroup_Desc.Width = 150
        '
        'ColGroup_Primary
        '
        Me.ColGroup_Primary.HeaderText = "Primary"
        Me.ColGroup_Primary.Name = "ColGroup_Primary"
        Me.ColGroup_Primary.Visible = False
        Me.ColGroup_Primary.Width = 66
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvBatchNo)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(428, 220)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Batch No"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'DgvBatchNo
        '
        Me.DgvBatchNo.AllowUserToAddRows = False
        Me.DgvBatchNo.AllowUserToResizeColumns = False
        Me.DgvBatchNo.AllowUserToResizeRows = False
        Me.DgvBatchNo.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvBatchNo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvBatchNo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColBatch_BatchNo})
        Me.DgvBatchNo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvBatchNo.Location = New System.Drawing.Point(2, 2)
        Me.DgvBatchNo.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvBatchNo.MultiSelect = False
        Me.DgvBatchNo.Name = "DgvBatchNo"
        Me.DgvBatchNo.RowTemplate.Height = 24
        Me.DgvBatchNo.Size = New System.Drawing.Size(424, 216)
        Me.DgvBatchNo.TabIndex = 3
        '
        'ColBatch_BatchNo
        '
        Me.ColBatch_BatchNo.HeaderText = "Batch No"
        Me.ColBatch_BatchNo.MaxInputLength = 250
        Me.ColBatch_BatchNo.MinimumWidth = 100
        Me.ColBatch_BatchNo.Name = "ColBatch_BatchNo"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.GboAdjustmentType)
        Me.TabPage4.Controls.Add(Me.dgvCustomer)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage4.Size = New System.Drawing.Size(428, 220)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Adjustment Type"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'GboAdjustmentType
        '
        Me.GboAdjustmentType.Controls.Add(Me.LinkLabel3)
        Me.GboAdjustmentType.Controls.Add(Me.LinkLabel4)
        Me.GboAdjustmentType.Controls.Add(Me.clbAdjustmentType)
        Me.GboAdjustmentType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GboAdjustmentType.Location = New System.Drawing.Point(2, 2)
        Me.GboAdjustmentType.Name = "GboAdjustmentType"
        Me.GboAdjustmentType.Size = New System.Drawing.Size(424, 216)
        Me.GboAdjustmentType.TabIndex = 5
        Me.GboAdjustmentType.TabStop = False
        '
        'LinkLabel3
        '
        Me.LinkLabel3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(336, 0)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel3.TabIndex = 2
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Uncheck All"
        '
        'LinkLabel4
        '
        Me.LinkLabel4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LinkLabel4.AutoSize = True
        Me.LinkLabel4.Location = New System.Drawing.Point(279, 0)
        Me.LinkLabel4.Name = "LinkLabel4"
        Me.LinkLabel4.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel4.TabIndex = 1
        Me.LinkLabel4.TabStop = True
        Me.LinkLabel4.Text = "Check All"
        '
        'clbAdjustmentType
        '
        Me.clbAdjustmentType.CheckOnClick = True
        Me.clbAdjustmentType.ColumnWidth = 300
        Me.clbAdjustmentType.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbAdjustmentType.FormattingEnabled = True
        Me.clbAdjustmentType.Location = New System.Drawing.Point(3, 16)
        Me.clbAdjustmentType.MultiColumn = True
        Me.clbAdjustmentType.Name = "clbAdjustmentType"
        Me.clbAdjustmentType.Size = New System.Drawing.Size(418, 184)
        Me.clbAdjustmentType.TabIndex = 0
        Me.clbAdjustmentType.ThreeDCheckBoxes = True
        '
        'dgvCustomer
        '
        Me.dgvCustomer.AllowUserToAddRows = False
        Me.dgvCustomer.AllowUserToResizeColumns = False
        Me.dgvCustomer.AllowUserToResizeRows = False
        Me.dgvCustomer.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCustomer.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCustomer.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColCustomer_Code, Me.ColCustomer_Desc, Me.ColCustomer_Primary})
        Me.dgvCustomer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCustomer.Location = New System.Drawing.Point(2, 2)
        Me.dgvCustomer.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvCustomer.MultiSelect = False
        Me.dgvCustomer.Name = "dgvCustomer"
        Me.dgvCustomer.RowTemplate.Height = 24
        Me.dgvCustomer.Size = New System.Drawing.Size(424, 216)
        Me.dgvCustomer.TabIndex = 4
        '
        'ColCustomer_Code
        '
        Me.ColCustomer_Code.HeaderText = "Code"
        Me.ColCustomer_Code.MaxInputLength = 250
        Me.ColCustomer_Code.MinimumWidth = 100
        Me.ColCustomer_Code.Name = "ColCustomer_Code"
        '
        'ColCustomer_Desc
        '
        Me.ColCustomer_Desc.HeaderText = "Name"
        Me.ColCustomer_Desc.MaxInputLength = 250
        Me.ColCustomer_Desc.MinimumWidth = 150
        Me.ColCustomer_Desc.Name = "ColCustomer_Desc"
        Me.ColCustomer_Desc.ReadOnly = True
        Me.ColCustomer_Desc.Width = 150
        '
        'ColCustomer_Primary
        '
        Me.ColCustomer_Primary.HeaderText = "Primary"
        Me.ColCustomer_Primary.Name = "ColCustomer_Primary"
        Me.ColCustomer_Primary.Visible = False
        Me.ColCustomer_Primary.Width = 66
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(0, 0)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(744, 258)
        Me.crv1.TabIndex = 0
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker3
        '
        '
        'BackgroundWorker4
        '
        '
        'ChkFilterByWareHouse
        '
        Me.ChkFilterByWareHouse.AutoSize = True
        Me.ChkFilterByWareHouse.Checked = True
        Me.ChkFilterByWareHouse.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkFilterByWareHouse.Location = New System.Drawing.Point(72, 196)
        Me.ChkFilterByWareHouse.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkFilterByWareHouse.Name = "ChkFilterByWareHouse"
        Me.ChkFilterByWareHouse.Size = New System.Drawing.Size(121, 17)
        Me.ChkFilterByWareHouse.TabIndex = 321
        Me.ChkFilterByWareHouse.Text = "Filter By Warehouse"
        Me.ChkFilterByWareHouse.UseVisualStyleBackColor = True
        '
        'FrmLaporanStockInOut
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(744, 514)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmLaporanStockInOut"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock In/Out Report"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DgvLocation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DgvGroup, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.DgvBatchNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage4.ResumeLayout(False)
        Me.GboAdjustmentType.ResumeLayout(False)
        Me.GboAdjustmentType.PerformLayout()
        CType(Me.dgvCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents CboPilihGroupLevel As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DtpMaxExpiryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpMinExpiryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents ColLocation_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DgvGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColGroup_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents DgvBatchNo As System.Windows.Forms.DataGridView
    Friend WithEvents ColBatch_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkShowAllItemsInTrans As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents GboAdjustmentType As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel4 As System.Windows.Forms.LinkLabel
    Friend WithEvents clbAdjustmentType As System.Windows.Forms.CheckedListBox
    Friend WithEvents ChkShowSubTotal As System.Windows.Forms.CheckBox
    Friend WithEvents CboScope As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents LblPONumber As System.Windows.Forms.Label
    Friend WithEvents btnFindPONumber As System.Windows.Forms.Button
    Friend WithEvents TxtPrimaryPO As System.Windows.Forms.TextBox
    Friend WithEvents TxtPONumber As System.Windows.Forms.TextBox
    Friend WithEvents CboInPurchase As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents dgvCustomer As System.Windows.Forms.DataGridView
    Friend WithEvents ColCustomer_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomer_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomer_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkShowAOPNo As System.Windows.Forms.CheckBox
    Friend WithEvents ChkFilterByDate As System.Windows.Forms.CheckBox
    Friend WithEvents ChkFilterByWareHouse As System.Windows.Forms.CheckBox
End Class
