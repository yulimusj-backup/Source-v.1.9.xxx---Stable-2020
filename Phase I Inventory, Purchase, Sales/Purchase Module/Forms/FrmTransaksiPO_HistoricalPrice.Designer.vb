<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPO_HistoricalPrice
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
        Me.dgvPrice = New System.Windows.Forms.DataGridView
        Me.Button1 = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTanggal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPONumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupplierCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSupplierName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCodeMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrencyCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPriceByMultiUnit_str = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRatio = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvPrice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvPrice
        '
        Me.dgvPrice.AllowUserToAddRows = False
        Me.dgvPrice.AllowUserToDeleteRows = False
        Me.dgvPrice.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dgvPrice.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvPrice.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPrice.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvPrice.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvPrice.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColTanggal, Me.ColPONumber, Me.ColSupplierCode, Me.ColSupplierName, Me.ColQtyMulti, Me.ColUnitCodeMulti, Me.ColCurrencyCode, Me.ColUnitPrice, Me.ColPriceByMultiUnit_str, Me.ColDiscount, Me.ColNettPrice, Me.ColRatio, Me.ColQty, Me.ColUnitCode})
        Me.dgvPrice.Location = New System.Drawing.Point(16, 15)
        Me.dgvPrice.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dgvPrice.Name = "dgvPrice"
        Me.dgvPrice.ReadOnly = True
        Me.dgvPrice.RowHeadersVisible = False
        Me.dgvPrice.RowTemplate.Height = 24
        Me.dgvPrice.Size = New System.Drawing.Size(1199, 338)
        Me.dgvPrice.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(1131, 361)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 30)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColTanggal
        '
        Me.ColTanggal.HeaderText = "Date"
        Me.ColTanggal.Name = "ColTanggal"
        Me.ColTanggal.ReadOnly = True
        Me.ColTanggal.Width = 63
        '
        'ColPONumber
        '
        Me.ColPONumber.HeaderText = "PO Number"
        Me.ColPONumber.Name = "ColPONumber"
        Me.ColPONumber.ReadOnly = True
        Me.ColPONumber.Width = 107
        '
        'ColSupplierCode
        '
        Me.ColSupplierCode.HeaderText = "Supplier Code"
        Me.ColSupplierCode.Name = "ColSupplierCode"
        Me.ColSupplierCode.ReadOnly = True
        Me.ColSupplierCode.Width = 122
        '
        'ColSupplierName
        '
        Me.ColSupplierName.HeaderText = "Supplier Name"
        Me.ColSupplierName.Name = "ColSupplierName"
        Me.ColSupplierName.ReadOnly = True
        Me.ColSupplierName.Width = 126
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.ReadOnly = True
        Me.ColQtyMulti.Width = 55
        '
        'ColUnitCodeMulti
        '
        Me.ColUnitCodeMulti.HeaderText = "Unit Code"
        Me.ColUnitCodeMulti.Name = "ColUnitCodeMulti"
        Me.ColUnitCodeMulti.ReadOnly = True
        Me.ColUnitCodeMulti.Width = 95
        '
        'ColCurrencyCode
        '
        Me.ColCurrencyCode.HeaderText = "Currency"
        Me.ColCurrencyCode.Name = "ColCurrencyCode"
        Me.ColCurrencyCode.ReadOnly = True
        Me.ColCurrencyCode.Width = 90
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.Width = 94
        '
        'ColPriceByMultiUnit_str
        '
        Me.ColPriceByMultiUnit_str.HeaderText = " "
        Me.ColPriceByMultiUnit_str.Name = "ColPriceByMultiUnit_str"
        Me.ColPriceByMultiUnit_str.ReadOnly = True
        Me.ColPriceByMultiUnit_str.Width = 37
        '
        'ColDiscount
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColDiscount.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColDiscount.HeaderText = "Discount (%)"
        Me.ColDiscount.Name = "ColDiscount"
        Me.ColDiscount.ReadOnly = True
        Me.ColDiscount.Width = 114
        '
        'ColNettPrice
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.Width = 95
        '
        'ColRatio
        '
        Me.ColRatio.HeaderText = "Ratio"
        Me.ColRatio.Name = "ColRatio"
        Me.ColRatio.ReadOnly = True
        Me.ColRatio.Width = 66
        '
        'ColQty
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColQty.HeaderText = "@ Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.Visible = False
        Me.ColQty.Width = 73
        '
        'ColUnitCode
        '
        Me.ColUnitCode.HeaderText = "@ Unit"
        Me.ColUnitCode.Name = "ColUnitCode"
        Me.ColUnitCode.ReadOnly = True
        Me.ColUnitCode.Visible = False
        Me.ColUnitCode.Width = 76
        '
        'FrmTransaksiPO_HistoricalPrice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(1241, 400)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.dgvPrice)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(1249, 436)
        Me.Name = "FrmTransaksiPO_HistoricalPrice"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Historical Purchase Price"
        CType(Me.dgvPrice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvPrice As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTanggal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPONumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupplierCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSupplierName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCodeMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrencyCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPriceByMultiUnit_str As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRatio As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitCode As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
