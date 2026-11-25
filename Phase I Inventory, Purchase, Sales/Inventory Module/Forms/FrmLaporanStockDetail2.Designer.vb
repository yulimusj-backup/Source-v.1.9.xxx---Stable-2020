<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanStockDetail2
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
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Label4 = New System.Windows.Forms.Label
        Me.Prg1 = New System.Windows.Forms.ProgressBar
        Me.chkUseWarehouse = New System.Windows.Forms.CheckBox
        Me.DtpStockasOff = New System.Windows.Forms.DateTimePicker
        Me.ChkStockAsOFF = New System.Windows.Forms.CheckBox
        Me.DtpMaxExpiryDate = New System.Windows.Forms.DateTimePicker
        Me.ChkMaxExpiryDate = New System.Windows.Forms.CheckBox
        Me.ChkShowSubTotal = New System.Windows.Forms.CheckBox
        Me.ChkMinExpiryDate = New System.Windows.Forms.CheckBox
        Me.DtpMinExpiryDate = New System.Windows.Forms.DateTimePicker
        Me.CboScope = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CboFormat = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ChkShowGrandTotal = New System.Windows.Forms.CheckBox
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
        Me.CboPilihGroupLevel = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.DgvGroup = New System.Windows.Forms.DataGridView
        Me.ColGroup_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColGroup_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.DgvBatchNo = New System.Windows.Forms.DataGridView
        Me.ColBatch_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker4 = New System.ComponentModel.BackgroundWorker
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
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Prg1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.chkUseWarehouse)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpStockasOff)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkStockAsOFF)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpMaxExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkMaxExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowSubTotal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkMinExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpMinExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboScope)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboFormat)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowGrandTotal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crv1)
        Me.SplitContainer1.Size = New System.Drawing.Size(1009, 635)
        Me.SplitContainer1.SplitterDistance = 204
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 226)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(0, 17)
        Me.Label4.TabIndex = 315
        '
        'Prg1
        '
        Me.Prg1.Location = New System.Drawing.Point(5, 194)
        Me.Prg1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Prg1.Name = "Prg1"
        Me.Prg1.Size = New System.Drawing.Size(403, 27)
        Me.Prg1.TabIndex = 48
        Me.Prg1.Visible = False
        '
        'chkUseWarehouse
        '
        Me.chkUseWarehouse.AutoSize = True
        Me.chkUseWarehouse.Location = New System.Drawing.Point(5, 174)
        Me.chkUseWarehouse.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.chkUseWarehouse.Name = "chkUseWarehouse"
        Me.chkUseWarehouse.Size = New System.Drawing.Size(132, 21)
        Me.chkUseWarehouse.TabIndex = 314
        Me.chkUseWarehouse.Text = "Use Warehouse"
        Me.chkUseWarehouse.UseVisualStyleBackColor = True
        '
        'DtpStockasOff
        '
        Me.DtpStockasOff.CustomFormat = "dd MMM yyyy"
        Me.DtpStockasOff.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpStockasOff.Location = New System.Drawing.Point(169, 149)
        Me.DtpStockasOff.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DtpStockasOff.Name = "DtpStockasOff"
        Me.DtpStockasOff.Size = New System.Drawing.Size(233, 22)
        Me.DtpStockasOff.TabIndex = 313
        '
        'ChkStockAsOFF
        '
        Me.ChkStockAsOFF.AutoSize = True
        Me.ChkStockAsOFF.Location = New System.Drawing.Point(5, 149)
        Me.ChkStockAsOFF.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkStockAsOFF.Name = "ChkStockAsOFF"
        Me.ChkStockAsOFF.Size = New System.Drawing.Size(104, 21)
        Me.ChkStockAsOFF.TabIndex = 312
        Me.ChkStockAsOFF.Text = "Stock As Of"
        Me.ChkStockAsOFF.UseVisualStyleBackColor = True
        '
        'DtpMaxExpiryDate
        '
        Me.DtpMaxExpiryDate.CustomFormat = "dd MMM yyyy"
        Me.DtpMaxExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpMaxExpiryDate.Location = New System.Drawing.Point(169, 123)
        Me.DtpMaxExpiryDate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DtpMaxExpiryDate.Name = "DtpMaxExpiryDate"
        Me.DtpMaxExpiryDate.Size = New System.Drawing.Size(233, 22)
        Me.DtpMaxExpiryDate.TabIndex = 7
        '
        'ChkMaxExpiryDate
        '
        Me.ChkMaxExpiryDate.AutoSize = True
        Me.ChkMaxExpiryDate.Location = New System.Drawing.Point(5, 123)
        Me.ChkMaxExpiryDate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkMaxExpiryDate.Name = "ChkMaxExpiryDate"
        Me.ChkMaxExpiryDate.Size = New System.Drawing.Size(135, 21)
        Me.ChkMaxExpiryDate.TabIndex = 4
        Me.ChkMaxExpiryDate.Text = "Max. Expiry Date"
        Me.ChkMaxExpiryDate.UseVisualStyleBackColor = True
        '
        'ChkShowSubTotal
        '
        Me.ChkShowSubTotal.AutoSize = True
        Me.ChkShowSubTotal.Location = New System.Drawing.Point(5, 74)
        Me.ChkShowSubTotal.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkShowSubTotal.Name = "ChkShowSubTotal"
        Me.ChkShowSubTotal.Size = New System.Drawing.Size(129, 21)
        Me.ChkShowSubTotal.TabIndex = 2
        Me.ChkShowSubTotal.Text = "Show Sub Total"
        Me.ChkShowSubTotal.UseVisualStyleBackColor = True
        '
        'ChkMinExpiryDate
        '
        Me.ChkMinExpiryDate.AutoSize = True
        Me.ChkMinExpiryDate.Location = New System.Drawing.Point(5, 97)
        Me.ChkMinExpiryDate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkMinExpiryDate.Name = "ChkMinExpiryDate"
        Me.ChkMinExpiryDate.Size = New System.Drawing.Size(132, 21)
        Me.ChkMinExpiryDate.TabIndex = 3
        Me.ChkMinExpiryDate.Text = "Min. Expiry Date"
        Me.ChkMinExpiryDate.UseVisualStyleBackColor = True
        '
        'DtpMinExpiryDate
        '
        Me.DtpMinExpiryDate.CustomFormat = "dd MMM yyyy"
        Me.DtpMinExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpMinExpiryDate.Location = New System.Drawing.Point(169, 97)
        Me.DtpMinExpiryDate.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DtpMinExpiryDate.Name = "DtpMinExpiryDate"
        Me.DtpMinExpiryDate.Size = New System.Drawing.Size(233, 22)
        Me.DtpMinExpiryDate.TabIndex = 6
        '
        'CboScope
        '
        Me.CboScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboScope.FormattingEnabled = True
        Me.CboScope.Items.AddRange(New Object() {"Goods Only", "Service Only", "Both"})
        Me.CboScope.Location = New System.Drawing.Point(64, 36)
        Me.CboScope.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CboScope.Name = "CboScope"
        Me.CboScope.Size = New System.Drawing.Size(241, 24)
        Me.CboScope.TabIndex = 1
        Me.CboScope.TabStop = False
        Me.CboScope.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 17)
        Me.Label2.TabIndex = 311
        Me.Label2.Text = "Scope"
        Me.Label2.Visible = False
        '
        'CboFormat
        '
        Me.CboFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboFormat.FormattingEnabled = True
        Me.CboFormat.Items.AddRange(New Object() {"No Grouping", "Group By Location", "Group By Product Group", "Group By Brand", "Group By Sub Brand", "Group By Category"})
        Me.CboFormat.Location = New System.Drawing.Point(64, 4)
        Me.CboFormat.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CboFormat.Name = "CboFormat"
        Me.CboFormat.Size = New System.Drawing.Size(241, 24)
        Me.CboFormat.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Format"
        '
        'ChkShowGrandTotal
        '
        Me.ChkShowGrandTotal.AutoSize = True
        Me.ChkShowGrandTotal.Location = New System.Drawing.Point(169, 74)
        Me.ChkShowGrandTotal.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ChkShowGrandTotal.Name = "ChkShowGrandTotal"
        Me.ChkShowGrandTotal.Size = New System.Drawing.Size(144, 21)
        Me.ChkShowGrandTotal.TabIndex = 5
        Me.ChkShowGrandTotal.Text = "Show Grand Total"
        Me.ChkShowGrandTotal.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(313, 34)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(91, 27)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(313, 2)
        Me.btnFilter.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(91, 27)
        Me.btnFilter.TabIndex = 8
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
        Me.TabControl1.Location = New System.Drawing.Point(413, 5)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(596, 194)
        Me.TabControl1.TabIndex = 10
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvLocation)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage1.Size = New System.Drawing.Size(588, 165)
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
        Me.DgvLocation.Location = New System.Drawing.Point(3, 2)
        Me.DgvLocation.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvLocation.MultiSelect = False
        Me.DgvLocation.Name = "DgvLocation"
        Me.DgvLocation.RowTemplate.Height = 24
        Me.DgvLocation.Size = New System.Drawing.Size(582, 161)
        Me.DgvLocation.TabIndex = 2
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
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage2.Size = New System.Drawing.Size(588, 165)
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
        Me.GroupBox1.Location = New System.Drawing.Point(3, 2)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(582, 161)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'CboPilihGroupLevel
        '
        Me.CboPilihGroupLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboPilihGroupLevel.FormattingEnabled = True
        Me.CboPilihGroupLevel.Items.AddRange(New Object() {"Product Group", "Brand", "Sub Brand", "Category", "Product"})
        Me.CboPilihGroupLevel.Location = New System.Drawing.Point(67, 14)
        Me.CboPilihGroupLevel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CboPilihGroupLevel.Name = "CboPilihGroupLevel"
        Me.CboPilihGroupLevel.Size = New System.Drawing.Size(236, 24)
        Me.CboPilihGroupLevel.TabIndex = 314
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 17)
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
        Me.DgvGroup.Location = New System.Drawing.Point(3, 44)
        Me.DgvGroup.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvGroup.MultiSelect = False
        Me.DgvGroup.Name = "DgvGroup"
        Me.DgvGroup.RowTemplate.Height = 24
        Me.DgvGroup.Size = New System.Drawing.Size(578, 109)
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
        Me.TabPage3.Location = New System.Drawing.Point(4, 25)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TabPage3.Size = New System.Drawing.Size(588, 212)
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
        Me.DgvBatchNo.Location = New System.Drawing.Point(3, 2)
        Me.DgvBatchNo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.DgvBatchNo.MultiSelect = False
        Me.DgvBatchNo.Name = "DgvBatchNo"
        Me.DgvBatchNo.RowTemplate.Height = 24
        Me.DgvBatchNo.Size = New System.Drawing.Size(582, 208)
        Me.DgvBatchNo.TabIndex = 3
        '
        'ColBatch_BatchNo
        '
        Me.ColBatch_BatchNo.HeaderText = "Batch No"
        Me.ColBatch_BatchNo.MaxInputLength = 250
        Me.ColBatch_BatchNo.MinimumWidth = 100
        Me.ColBatch_BatchNo.Name = "ColBatch_BatchNo"
        '
        'crv1
        '
        Me.crv1.ActiveViewIndex = -1
        Me.crv1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crv1.DisplayGroupTree = False
        Me.crv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crv1.EnableDrillDown = False
        Me.crv1.Location = New System.Drawing.Point(0, 0)
        Me.crv1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.crv1.Name = "crv1"
        Me.crv1.SelectionFormula = ""
        Me.crv1.ShowCloseButton = False
        Me.crv1.ShowGroupTreeButton = False
        Me.crv1.ShowRefreshButton = False
        Me.crv1.Size = New System.Drawing.Size(1009, 427)
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
        'FrmLaporanStockDetail2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1009, 635)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MinimumSize = New System.Drawing.Size(1005, 672)
        Me.Name = "FrmLaporanStockDetail2"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Detail Report"
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
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents DtpMaxExpiryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ChkMaxExpiryDate As System.Windows.Forms.CheckBox
    Friend WithEvents DtpMinExpiryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ChkMinExpiryDate As System.Windows.Forms.CheckBox
    Friend WithEvents CboScope As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ChkShowGrandTotal As System.Windows.Forms.CheckBox
    Friend WithEvents ChkShowSubTotal As System.Windows.Forms.CheckBox
    Friend WithEvents CboFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CboPilihGroupLevel As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DgvGroup As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents DgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents DgvBatchNo As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColLocation_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColGroup_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatch_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkStockAsOFF As System.Windows.Forms.CheckBox
    Friend WithEvents DtpStockasOff As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkUseWarehouse As System.Windows.Forms.CheckBox
    Friend WithEvents BackgroundWorker4 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Prg1 As System.Windows.Forms.ProgressBar
End Class
