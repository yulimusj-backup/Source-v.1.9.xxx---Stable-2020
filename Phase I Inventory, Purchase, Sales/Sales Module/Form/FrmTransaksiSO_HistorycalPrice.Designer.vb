<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSO_HistorycalPrice
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
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFindCustomer = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPrimaryCustomer = New System.Windows.Forms.TextBox
        Me.txtCustomerCode = New System.Windows.Forms.TextBox
        Me.TxtCustomerName = New System.Windows.Forms.TextBox
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOrderQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatio = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOrderQtyDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCrrr = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriceByMultiUnit_Str = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDisc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriceExclude = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColSONo, Me.ColPONO, Me.ColDate, Me.ColOrderQtyMulti, Me.ColUnitMulti, Me.ColRatio, Me.ColOrderQtyDefault, Me.ColUnitDefault, Me.ColCrrr, Me.ColUnitPrice, Me.ColPriceByMultiUnit_Str, Me.ColDisc, Me.ColDiscAmount, Me.ColNettPrice, Me.ColAmount, Me.ColPriceExclude, Me.ColRemarks})
        Me.dgvItem.Location = New System.Drawing.Point(16, 47)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(950, 491)
        Me.dgvItem.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(866, 545)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 28)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnFindCustomer
        '
        Me.btnFindCustomer.Location = New System.Drawing.Point(251, 15)
        Me.btnFindCustomer.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFindCustomer.Name = "btnFindCustomer"
        Me.btnFindCustomer.Size = New System.Drawing.Size(43, 25)
        Me.btnFindCustomer.TabIndex = 2
        Me.btnFindCustomer.TabStop = False
        Me.btnFindCustomer.Text = "...."
        Me.btnFindCustomer.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 21)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 17)
        Me.Label1.TabIndex = 137
        Me.Label1.Text = "Customer"
        '
        'txtPrimaryCustomer
        '
        Me.txtPrimaryCustomer.Location = New System.Drawing.Point(115, 15)
        Me.txtPrimaryCustomer.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPrimaryCustomer.Name = "txtPrimaryCustomer"
        Me.txtPrimaryCustomer.Size = New System.Drawing.Size(32, 22)
        Me.txtPrimaryCustomer.TabIndex = 134
        Me.txtPrimaryCustomer.TabStop = False
        Me.txtPrimaryCustomer.Text = "-1"
        Me.txtPrimaryCustomer.Visible = False
        '
        'txtCustomerCode
        '
        Me.txtCustomerCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCustomerCode.Location = New System.Drawing.Point(97, 15)
        Me.txtCustomerCode.Margin = New System.Windows.Forms.Padding(4)
        Me.txtCustomerCode.MaxLength = 50
        Me.txtCustomerCode.Name = "txtCustomerCode"
        Me.txtCustomerCode.Size = New System.Drawing.Size(144, 22)
        Me.txtCustomerCode.TabIndex = 1
        '
        'TxtCustomerName
        '
        Me.TxtCustomerName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtCustomerName.Location = New System.Drawing.Point(301, 15)
        Me.TxtCustomerName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCustomerName.MaxLength = 150
        Me.TxtCustomerName.Name = "TxtCustomerName"
        Me.TxtCustomerName.ReadOnly = True
        Me.TxtCustomerName.Size = New System.Drawing.Size(664, 22)
        Me.TxtCustomerName.TabIndex = 3
        Me.TxtCustomerName.TabStop = False
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColSONo
        '
        Me.ColSONo.HeaderText = "SO No"
        Me.ColSONo.Name = "ColSONo"
        Me.ColSONo.ReadOnly = True
        Me.ColSONo.Width = 75
        '
        'ColPONO
        '
        Me.ColPONO.HeaderText = "Cust. PO No"
        Me.ColPONO.Name = "ColPONO"
        Me.ColPONO.ReadOnly = True
        Me.ColPONO.Width = 111
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.Width = 63
        '
        'ColOrderQtyMulti
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOrderQtyMulti.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColOrderQtyMulti.HeaderText = "Order Qty"
        Me.ColOrderQtyMulti.Name = "ColOrderQtyMulti"
        Me.ColOrderQtyMulti.ReadOnly = True
        Me.ColOrderQtyMulti.Width = 96
        '
        'ColUnitMulti
        '
        Me.ColUnitMulti.HeaderText = ""
        Me.ColUnitMulti.Name = "ColUnitMulti"
        Me.ColUnitMulti.ReadOnly = True
        Me.ColUnitMulti.Width = 19
        '
        'ColRatio
        '
        Me.ColRatio.HeaderText = "Ratio"
        Me.ColRatio.Name = "ColRatio"
        Me.ColRatio.ReadOnly = True
        Me.ColRatio.Width = 66
        '
        'ColOrderQtyDefault
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOrderQtyDefault.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColOrderQtyDefault.HeaderText = "Order Qty (Default)"
        Me.ColOrderQtyDefault.Name = "ColOrderQtyDefault"
        Me.ColOrderQtyDefault.ReadOnly = True
        Me.ColOrderQtyDefault.Width = 155
        '
        'ColUnitDefault
        '
        Me.ColUnitDefault.HeaderText = ""
        Me.ColUnitDefault.Name = "ColUnitDefault"
        Me.ColUnitDefault.ReadOnly = True
        Me.ColUnitDefault.Width = 19
        '
        'ColCrrr
        '
        Me.ColCrrr.HeaderText = "Curr."
        Me.ColCrrr.Name = "ColCrrr"
        Me.ColCrrr.ReadOnly = True
        Me.ColCrrr.Width = 64
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.Width = 94
        '
        'ColPriceByMultiUnit_Str
        '
        Me.ColPriceByMultiUnit_Str.HeaderText = ""
        Me.ColPriceByMultiUnit_Str.Name = "ColPriceByMultiUnit_Str"
        Me.ColPriceByMultiUnit_Str.ReadOnly = True
        Me.ColPriceByMultiUnit_Str.Width = 19
        '
        'ColDisc
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDisc.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColDisc.HeaderText = "Disc %"
        Me.ColDisc.Name = "ColDisc"
        Me.ColDisc.ReadOnly = True
        Me.ColDisc.Width = 76
        '
        'ColDiscAmount
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDiscAmount.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColDiscAmount.HeaderText = "Disc. Amount"
        Me.ColDiscAmount.Name = "ColDiscAmount"
        Me.ColDiscAmount.ReadOnly = True
        Me.ColDiscAmount.Width = 116
        '
        'ColNettPrice
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.Width = 95
        '
        'ColAmount
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.Width = 81
        '
        'ColPriceExclude
        '
        Me.ColPriceExclude.HeaderText = "Exclude VAT"
        Me.ColPriceExclude.Name = "ColPriceExclude"
        Me.ColPriceExclude.ReadOnly = True
        Me.ColPriceExclude.Width = 94
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 70
        '
        'FrmTransaksiSO_HistorycalPrice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(982, 588)
        Me.Controls.Add(Me.btnFindCustomer)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtPrimaryCustomer)
        Me.Controls.Add(Me.txtCustomerCode)
        Me.Controls.Add(Me.TxtCustomerName)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.dgvItem)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1000, 600)
        Me.Name = "FrmTransaksiSO_HistorycalPrice"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Historical Sales Price"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnFindCustomer As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPrimaryCustomer As System.Windows.Forms.TextBox
    Friend WithEvents txtCustomerCode As System.Windows.Forms.TextBox
    Friend WithEvents TxtCustomerName As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOrderQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatio As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOrderQtyDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCrrr As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriceByMultiUnit_Str As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDisc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriceExclude As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
