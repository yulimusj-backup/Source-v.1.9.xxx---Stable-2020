<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanStockDetail
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmLaporanStockDetail))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCodeNo = New System.Windows.Forms.TextBox
        Me.BtnFindCodeNo = New System.Windows.Forms.Button
        Me.Stock_detail_viewBindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.Stock_detail_viewBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DseLaporanStockDetail = New Inventory_Module.dseLaporanStockDetail
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.Stock_detail_viewDataGridView = New System.Windows.Forms.DataGridView
        Me.txtPrimaryCategory = New System.Windows.Forms.TextBox
        Me.TxtPrimaryCodeNo = New System.Windows.Forms.TextBox
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.Stock_detail_viewTableAdapter = New Inventory_Module.dseLaporanStockDetailTableAdapters.stock_detail_viewTableAdapter
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn11 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn12 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.DataGridViewTextBoxColumn14 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Quantity = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.Stock_detail_viewBindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Stock_detail_viewBindingNavigator.SuspendLayout()
        CType(Me.Stock_detail_viewBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DseLaporanStockDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Stock_detail_viewDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Code No"
        '
        'txtCodeNo
        '
        Me.txtCodeNo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodeNo.Location = New System.Drawing.Point(81, 12)
        Me.txtCodeNo.Name = "txtCodeNo"
        Me.txtCodeNo.Size = New System.Drawing.Size(209, 20)
        Me.txtCodeNo.TabIndex = 0
        '
        'BtnFindCodeNo
        '
        Me.BtnFindCodeNo.Location = New System.Drawing.Point(296, 12)
        Me.BtnFindCodeNo.Name = "BtnFindCodeNo"
        Me.BtnFindCodeNo.Size = New System.Drawing.Size(31, 20)
        Me.BtnFindCodeNo.TabIndex = 1
        Me.BtnFindCodeNo.TabStop = False
        Me.BtnFindCodeNo.Text = "..."
        Me.BtnFindCodeNo.UseVisualStyleBackColor = True
        '
        'Stock_detail_viewBindingNavigator
        '
        Me.Stock_detail_viewBindingNavigator.AddNewItem = Nothing
        Me.Stock_detail_viewBindingNavigator.BindingSource = Me.Stock_detail_viewBindingSource
        Me.Stock_detail_viewBindingNavigator.CountItem = Me.BindingNavigatorCountItem
        Me.Stock_detail_viewBindingNavigator.DeleteItem = Nothing
        Me.Stock_detail_viewBindingNavigator.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Stock_detail_viewBindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2})
        Me.Stock_detail_viewBindingNavigator.Location = New System.Drawing.Point(0, 468)
        Me.Stock_detail_viewBindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.Stock_detail_viewBindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.Stock_detail_viewBindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.Stock_detail_viewBindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.Stock_detail_viewBindingNavigator.Name = "Stock_detail_viewBindingNavigator"
        Me.Stock_detail_viewBindingNavigator.PositionItem = Me.BindingNavigatorPositionItem
        Me.Stock_detail_viewBindingNavigator.Size = New System.Drawing.Size(891, 25)
        Me.Stock_detail_viewBindingNavigator.TabIndex = 6
        Me.Stock_detail_viewBindingNavigator.Text = "BindingNavigator1"
        '
        'Stock_detail_viewBindingSource
        '
        Me.Stock_detail_viewBindingSource.DataMember = "stock_detail_view"
        Me.Stock_detail_viewBindingSource.DataSource = Me.DseLaporanStockDetail
        '
        'DseLaporanStockDetail
        '
        Me.DseLaporanStockDetail.DataSetName = "dseLaporanStockDetail"
        Me.DseLaporanStockDetail.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 22)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'Stock_detail_viewDataGridView
        '
        Me.Stock_detail_viewDataGridView.AllowUserToAddRows = False
        Me.Stock_detail_viewDataGridView.AllowUserToDeleteRows = False
        Me.Stock_detail_viewDataGridView.AllowUserToResizeColumns = False
        Me.Stock_detail_viewDataGridView.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Stock_detail_viewDataGridView.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Stock_detail_viewDataGridView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Stock_detail_viewDataGridView.AutoGenerateColumns = False
        Me.Stock_detail_viewDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.Stock_detail_viewDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.ColQty, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8, Me.DataGridViewTextBoxColumn9, Me.DataGridViewTextBoxColumn10, Me.DataGridViewTextBoxColumn11, Me.DataGridViewTextBoxColumn12, Me.DataGridViewTextBoxColumn13, Me.DataGridViewTextBoxColumn14, Me.Quantity})
        Me.Stock_detail_viewDataGridView.DataSource = Me.Stock_detail_viewBindingSource
        Me.Stock_detail_viewDataGridView.Location = New System.Drawing.Point(12, 64)
        Me.Stock_detail_viewDataGridView.MultiSelect = False
        Me.Stock_detail_viewDataGridView.Name = "Stock_detail_viewDataGridView"
        Me.Stock_detail_viewDataGridView.ReadOnly = True
        Me.Stock_detail_viewDataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.Stock_detail_viewDataGridView.Size = New System.Drawing.Size(867, 401)
        Me.Stock_detail_viewDataGridView.TabIndex = 5
        Me.Stock_detail_viewDataGridView.TabStop = False
        '
        'txtPrimaryCategory
        '
        Me.txtPrimaryCategory.Location = New System.Drawing.Point(538, 13)
        Me.txtPrimaryCategory.Name = "txtPrimaryCategory"
        Me.txtPrimaryCategory.Size = New System.Drawing.Size(25, 20)
        Me.txtPrimaryCategory.TabIndex = 97
        Me.txtPrimaryCategory.TabStop = False
        Me.txtPrimaryCategory.Visible = False
        '
        'TxtPrimaryCodeNo
        '
        Me.TxtPrimaryCodeNo.Location = New System.Drawing.Point(363, 12)
        Me.TxtPrimaryCodeNo.Name = "TxtPrimaryCodeNo"
        Me.TxtPrimaryCodeNo.Size = New System.Drawing.Size(25, 20)
        Me.TxtPrimaryCodeNo.TabIndex = 96
        Me.TxtPrimaryCodeNo.TabStop = False
        Me.TxtPrimaryCodeNo.Visible = False
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(333, 11)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(55, 22)
        Me.btnRefresh.TabIndex = 3
        Me.btnRefresh.Text = "&Refresh"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(81, 38)
        Me.txtDesc.MaxLength = 6
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReadOnly = True
        Me.txtDesc.Size = New System.Drawing.Size(246, 20)
        Me.txtDesc.TabIndex = 2
        Me.txtDesc.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(12, 41)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 13)
        Me.Label11.TabIndex = 100
        Me.Label11.Text = "Description"
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(333, 37)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(55, 21)
        Me.btnClose.TabIndex = 4
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Stock_detail_viewTableAdapter
        '
        Me.Stock_detail_viewTableAdapter.ClearBeforeFill = True
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "Category Code"
        Me.DataGridViewTextBoxColumn1.HeaderText = "Category Code"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.ReadOnly = True
        Me.DataGridViewTextBoxColumn1.Width = 102
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "Category Description"
        Me.DataGridViewTextBoxColumn2.HeaderText = "Category Description"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 130
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "Product Code"
        Me.DataGridViewTextBoxColumn3.HeaderText = "Product Code"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        Me.DataGridViewTextBoxColumn3.Width = 97
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "Specification"
        Me.DataGridViewTextBoxColumn4.HeaderText = "Specification"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        Me.DataGridViewTextBoxColumn4.Width = 93
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "Product Description"
        Me.DataGridViewTextBoxColumn5.HeaderText = "Product Description"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        Me.DataGridViewTextBoxColumn5.Width = 125
        '
        'ColQty
        '
        Me.ColQty.DataPropertyName = "Qty"
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.Width = 48
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.DataPropertyName = "Qty Secondary"
        Me.DataGridViewTextBoxColumn7.HeaderText = "Qty Secondary"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        Me.DataGridViewTextBoxColumn7.Width = 102
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.DataPropertyName = "Lot No"
        Me.DataGridViewTextBoxColumn8.HeaderText = "Lot No"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.ReadOnly = True
        Me.DataGridViewTextBoxColumn8.Width = 64
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.DataPropertyName = "Expiry Date"
        DataGridViewCellStyle2.Format = "dd MMM yyyy"
        Me.DataGridViewTextBoxColumn9.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridViewTextBoxColumn9.HeaderText = "Expiry Date"
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.ReadOnly = True
        Me.DataGridViewTextBoxColumn9.Width = 86
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.DataPropertyName = "Primary Storage Area"
        Me.DataGridViewTextBoxColumn10.HeaderText = "Primary Storage Area"
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.ReadOnly = True
        Me.DataGridViewTextBoxColumn10.Visible = False
        Me.DataGridViewTextBoxColumn10.Width = 131
        '
        'DataGridViewTextBoxColumn11
        '
        Me.DataGridViewTextBoxColumn11.DataPropertyName = "Location"
        Me.DataGridViewTextBoxColumn11.HeaderText = "Location"
        Me.DataGridViewTextBoxColumn11.Name = "DataGridViewTextBoxColumn11"
        Me.DataGridViewTextBoxColumn11.ReadOnly = True
        Me.DataGridViewTextBoxColumn11.Width = 73
        '
        'DataGridViewTextBoxColumn12
        '
        Me.DataGridViewTextBoxColumn12.DataPropertyName = "Primary Unit"
        Me.DataGridViewTextBoxColumn12.HeaderText = "Primary Unit"
        Me.DataGridViewTextBoxColumn12.Name = "DataGridViewTextBoxColumn12"
        Me.DataGridViewTextBoxColumn12.ReadOnly = True
        Me.DataGridViewTextBoxColumn12.Visible = False
        Me.DataGridViewTextBoxColumn12.Width = 88
        '
        'DataGridViewTextBoxColumn13
        '
        Me.DataGridViewTextBoxColumn13.DataPropertyName = "Primary Master Unit"
        Me.DataGridViewTextBoxColumn13.HeaderText = "Primary Master Unit"
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.ReadOnly = True
        Me.DataGridViewTextBoxColumn13.Visible = False
        Me.DataGridViewTextBoxColumn13.Width = 123
        '
        'DataGridViewTextBoxColumn14
        '
        Me.DataGridViewTextBoxColumn14.DataPropertyName = "Primary Secondary Unit"
        Me.DataGridViewTextBoxColumn14.HeaderText = "Primary Secondary Unit"
        Me.DataGridViewTextBoxColumn14.Name = "DataGridViewTextBoxColumn14"
        Me.DataGridViewTextBoxColumn14.ReadOnly = True
        Me.DataGridViewTextBoxColumn14.Visible = False
        Me.DataGridViewTextBoxColumn14.Width = 142
        '
        'Quantity
        '
        Me.Quantity.DataPropertyName = "Quantity"
        Me.Quantity.HeaderText = "Quantity"
        Me.Quantity.Name = "Quantity"
        Me.Quantity.ReadOnly = True
        Me.Quantity.Width = 71
        '
        'FrmLaporanStockDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(891, 493)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtDesc)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.txtPrimaryCategory)
        Me.Controls.Add(Me.TxtPrimaryCodeNo)
        Me.Controls.Add(Me.Stock_detail_viewDataGridView)
        Me.Controls.Add(Me.Stock_detail_viewBindingNavigator)
        Me.Controls.Add(Me.BtnFindCodeNo)
        Me.Controls.Add(Me.txtCodeNo)
        Me.Controls.Add(Me.Label1)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(907, 529)
        Me.Name = "FrmLaporanStockDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Stock Detail Reports"
        CType(Me.Stock_detail_viewBindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Stock_detail_viewBindingNavigator.ResumeLayout(False)
        Me.Stock_detail_viewBindingNavigator.PerformLayout()
        CType(Me.Stock_detail_viewBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DseLaporanStockDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Stock_detail_viewDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents BtnFindCodeNo As System.Windows.Forms.Button
    Friend WithEvents DseLaporanStockDetail As Inventory_Module.dseLaporanStockDetail
    Friend WithEvents Stock_detail_viewBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Stock_detail_viewTableAdapter As Inventory_Module.dseLaporanStockDetailTableAdapters.stock_detail_viewTableAdapter
    Friend WithEvents Stock_detail_viewBindingNavigator As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Stock_detail_viewDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents txtPrimaryCategory As System.Windows.Forms.TextBox
    Friend WithEvents TxtPrimaryCodeNo As System.Windows.Forms.TextBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn11 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn12 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn14 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Quantity As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
