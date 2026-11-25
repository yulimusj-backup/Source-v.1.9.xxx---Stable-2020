<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanStockTransfer
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLaporanStockTransfer))
        Me.ColBatch_BatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DgvBatchNo = New System.Windows.Forms.DataGridView
        Me.crv1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.ColGroup_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ColGroup_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker
        Me.ChkShowAllItemsInTrans = New System.Windows.Forms.CheckBox
        Me.ColGroup_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.CboPilihGroupLevel = New System.Windows.Forms.ComboBox
        Me.DtpMaxExpiryDate = New System.Windows.Forms.DateTimePicker
        Me.DtpMinExpiryDate = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFilter = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.CboScope = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ChkShowSubTotal = New System.Windows.Forms.CheckBox
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.DgvLocation = New System.Windows.Forms.DataGridView
        Me.ColLocation_Code = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Desc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation_Source = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_Destination = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColLocation_Primary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DgvGroup = New System.Windows.Forms.DataGridView
        Me.ChkFilterByWareHouse = New System.Windows.Forms.CheckBox
        CType(Me.DgvBatchNo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage3.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DgvLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DgvGroup, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ColBatch_BatchNo
        '
        Me.ColBatch_BatchNo.HeaderText = "Batch No"
        Me.ColBatch_BatchNo.MaxInputLength = 250
        Me.ColBatch_BatchNo.MinimumWidth = 100
        Me.ColBatch_BatchNo.Name = "ColBatch_BatchNo"
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
        Me.DgvBatchNo.Size = New System.Drawing.Size(422, 172)
        Me.DgvBatchNo.TabIndex = 3
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
        Me.crv1.Size = New System.Drawing.Size(744, 302)
        Me.crv1.TabIndex = 0
        Me.crv1.TabStop = False
        Me.crv1.ViewTimeSelectionFormula = ""
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
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.DgvBatchNo)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage3.Size = New System.Drawing.Size(426, 176)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Batch No"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'ColGroup_Primary
        '
        Me.ColGroup_Primary.HeaderText = "Primary"
        Me.ColGroup_Primary.Name = "ColGroup_Primary"
        Me.ColGroup_Primary.Visible = False
        Me.ColGroup_Primary.Width = 66
        '
        'BackgroundWorker2
        '
        '
        'BackgroundWorker3
        '
        '
        'ChkShowAllItemsInTrans
        '
        Me.ChkShowAllItemsInTrans.AutoSize = True
        Me.ChkShowAllItemsInTrans.Checked = True
        Me.ChkShowAllItemsInTrans.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkShowAllItemsInTrans.Location = New System.Drawing.Point(48, 78)
        Me.ChkShowAllItemsInTrans.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowAllItemsInTrans.Name = "ChkShowAllItemsInTrans"
        Me.ChkShowAllItemsInTrans.Size = New System.Drawing.Size(160, 17)
        Me.ChkShowAllItemsInTrans.TabIndex = 3
        Me.ChkShowAllItemsInTrans.Text = "Show All Item in Transaction"
        Me.ChkShowAllItemsInTrans.UseVisualStyleBackColor = True
        '
        'ColGroup_Code
        '
        Me.ColGroup_Code.HeaderText = "Code"
        Me.ColGroup_Code.MaxInputLength = 250
        Me.ColGroup_Code.MinimumWidth = 100
        Me.ColGroup_Code.Name = "ColGroup_Code"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(4, 32)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(20, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "To"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(4, 9)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "From"
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
        'DtpMaxExpiryDate
        '
        Me.DtpMaxExpiryDate.CustomFormat = "dd MMM yyyy"
        Me.DtpMaxExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpMaxExpiryDate.Location = New System.Drawing.Point(48, 28)
        Me.DtpMaxExpiryDate.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpMaxExpiryDate.Name = "DtpMaxExpiryDate"
        Me.DtpMaxExpiryDate.Size = New System.Drawing.Size(182, 20)
        Me.DtpMaxExpiryDate.TabIndex = 1
        '
        'DtpMinExpiryDate
        '
        Me.DtpMinExpiryDate.CustomFormat = "dd MMM yyyy"
        Me.DtpMinExpiryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtpMinExpiryDate.Location = New System.Drawing.Point(48, 6)
        Me.DtpMinExpiryDate.Margin = New System.Windows.Forms.Padding(2)
        Me.DtpMinExpiryDate.Name = "DtpMinExpiryDate"
        Me.DtpMinExpiryDate.Size = New System.Drawing.Size(182, 20)
        Me.DtpMinExpiryDate.TabIndex = 0
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
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(235, 29)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(68, 20)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFilter
        '
        Me.btnFilter.Location = New System.Drawing.Point(235, 4)
        Me.btnFilter.Name = "btnFilter"
        Me.btnFilter.Size = New System.Drawing.Size(68, 20)
        Me.btnFilter.TabIndex = 5
        Me.btnFilter.Text = "Filter"
        Me.btnFilter.UseVisualStyleBackColor = True
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
        Me.SplitContainer1.Panel1.Controls.Add(Me.CboScope)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowSubTotal)
        Me.SplitContainer1.Panel1.Controls.Add(Me.ChkShowAllItemsInTrans)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpMaxExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.DtpMinExpiryDate)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnFilter)
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.crv1)
        Me.SplitContainer1.Size = New System.Drawing.Size(744, 514)
        Me.SplitContainer1.SplitterDistance = 209
        Me.SplitContainer1.SplitterWidth = 3
        Me.SplitContainer1.TabIndex = 0
        Me.SplitContainer1.TabStop = False
        '
        'CboScope
        '
        Me.CboScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboScope.FormattingEnabled = True
        Me.CboScope.Items.AddRange(New Object() {"Goods", "Service", "Both"})
        Me.CboScope.Location = New System.Drawing.Point(48, 120)
        Me.CboScope.Margin = New System.Windows.Forms.Padding(2)
        Me.CboScope.Name = "CboScope"
        Me.CboScope.Size = New System.Drawing.Size(182, 21)
        Me.CboScope.TabIndex = 2
        Me.CboScope.TabStop = False
        Me.CboScope.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 120)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 313
        Me.Label1.Text = "Scope"
        Me.Label1.Visible = False
        '
        'ChkShowSubTotal
        '
        Me.ChkShowSubTotal.AutoSize = True
        Me.ChkShowSubTotal.Checked = True
        Me.ChkShowSubTotal.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkShowSubTotal.Location = New System.Drawing.Point(48, 99)
        Me.ChkShowSubTotal.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkShowSubTotal.Name = "ChkShowSubTotal"
        Me.ChkShowSubTotal.Size = New System.Drawing.Size(102, 17)
        Me.ChkShowSubTotal.TabIndex = 4
        Me.ChkShowSubTotal.Text = "Show Sub Total"
        Me.ChkShowSubTotal.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(310, 4)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(434, 202)
        Me.TabControl1.TabIndex = 7
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.DgvLocation)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage1.Size = New System.Drawing.Size(426, 176)
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
        Me.DgvLocation.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColLocation_Code, Me.ColLocation_Desc, Me.ColLocation_Source, Me.ColLocation_Destination, Me.ColLocation_Primary})
        Me.DgvLocation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvLocation.Location = New System.Drawing.Point(2, 2)
        Me.DgvLocation.Margin = New System.Windows.Forms.Padding(2)
        Me.DgvLocation.MultiSelect = False
        Me.DgvLocation.Name = "DgvLocation"
        Me.DgvLocation.RowTemplate.Height = 24
        Me.DgvLocation.Size = New System.Drawing.Size(422, 172)
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
        'ColLocation_Source
        '
        Me.ColLocation_Source.HeaderText = "Source"
        Me.ColLocation_Source.Name = "ColLocation_Source"
        Me.ColLocation_Source.Width = 47
        '
        'ColLocation_Destination
        '
        Me.ColLocation_Destination.HeaderText = "Destination"
        Me.ColLocation_Destination.Name = "ColLocation_Destination"
        Me.ColLocation_Destination.Width = 66
        '
        'ColLocation_Primary
        '
        Me.ColLocation_Primary.HeaderText = "Primary"
        Me.ColLocation_Primary.Name = "ColLocation_Primary"
        Me.ColLocation_Primary.Visible = False
        Me.ColLocation_Primary.Width = 81
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.GroupBox1)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPage2.Size = New System.Drawing.Size(426, 176)
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
        Me.GroupBox1.Size = New System.Drawing.Size(422, 172)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
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
        Me.DgvGroup.Size = New System.Drawing.Size(418, 131)
        Me.DgvGroup.TabIndex = 0
        '
        'ChkFilterByWareHouse
        '
        Me.ChkFilterByWareHouse.AutoSize = True
        Me.ChkFilterByWareHouse.Checked = True
        Me.ChkFilterByWareHouse.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ChkFilterByWareHouse.Location = New System.Drawing.Point(48, 57)
        Me.ChkFilterByWareHouse.Margin = New System.Windows.Forms.Padding(2)
        Me.ChkFilterByWareHouse.Name = "ChkFilterByWareHouse"
        Me.ChkFilterByWareHouse.Size = New System.Drawing.Size(121, 17)
        Me.ChkFilterByWareHouse.TabIndex = 322
        Me.ChkFilterByWareHouse.Text = "Filter By Warehouse"
        Me.ChkFilterByWareHouse.UseVisualStyleBackColor = True
        '
        'FrmLaporanStockTransfer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(744, 514)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(754, 552)
        Me.Name = "FrmLaporanStockTransfer"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Transfer Report"
        CType(Me.DgvBatchNo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage3.ResumeLayout(False)
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
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ColBatch_BatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DgvBatchNo As System.Windows.Forms.DataGridView
    Friend WithEvents crv1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents ColGroup_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColGroup_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ChkShowAllItemsInTrans As System.Windows.Forms.CheckBox
    Friend WithEvents ColGroup_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CboPilihGroupLevel As System.Windows.Forms.ComboBox
    Friend WithEvents DtpMaxExpiryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtpMinExpiryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFilter As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DgvLocation As System.Windows.Forms.DataGridView
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents DgvGroup As System.Windows.Forms.DataGridView
    Friend WithEvents ColLocation_Code As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Desc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation_Source As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_Destination As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColLocation_Primary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ChkShowSubTotal As System.Windows.Forms.CheckBox
    Friend WithEvents CboScope As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ChkFilterByWareHouse As System.Windows.Forms.CheckBox
End Class
