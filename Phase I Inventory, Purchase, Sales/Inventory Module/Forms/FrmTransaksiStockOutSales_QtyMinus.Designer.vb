<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockOutSales_QtyMinus
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvMinus = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWarehouse = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyStockOut = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyOnHand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColShortage = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.dgvMinus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvMinus
        '
        Me.dgvMinus.AllowUserToAddRows = False
        Me.dgvMinus.AllowUserToDeleteRows = False
        Me.dgvMinus.AllowUserToResizeColumns = False
        Me.dgvMinus.AllowUserToResizeRows = False
        Me.dgvMinus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMinus.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvMinus.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvMinus.ColumnHeadersHeight = 25
        Me.dgvMinus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvMinus.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDesc, Me.ColBatch, Me.ColExpiry, Me.ColLocation, Me.ColWarehouse, Me.ColQtyStockOut, Me.ColQtyOnHand, Me.ColShortage, Me.ColUnit})
        Me.dgvMinus.Location = New System.Drawing.Point(12, 29)
        Me.dgvMinus.Name = "dgvMinus"
        Me.dgvMinus.RowHeadersVisible = False
        Me.dgvMinus.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvMinus.RowTemplate.Height = 24
        Me.dgvMinus.Size = New System.Drawing.Size(743, 440)
        Me.dgvMinus.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 51
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.Width = 88
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.Width = 104
        '
        'ColBatch
        '
        Me.ColBatch.HeaderText = "Batch No"
        Me.ColBatch.Name = "ColBatch"
        Me.ColBatch.Width = 91
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.Width = 71
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.Width = 87
        '
        'ColWarehouse
        '
        Me.ColWarehouse.HeaderText = "Wahouse"
        Me.ColWarehouse.Name = "ColWarehouse"
        Me.ColWarehouse.Width = 93
        '
        'ColQtyStockOut
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyStockOut.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQtyStockOut.HeaderText = "Out Qty"
        Me.ColQtyStockOut.Name = "ColQtyStockOut"
        Me.ColQtyStockOut.Width = 82
        '
        'ColQtyOnHand
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyOnHand.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColQtyOnHand.HeaderText = "Qty On Hand"
        Me.ColQtyOnHand.Name = "ColQtyOnHand"
        Me.ColQtyOnHand.Width = 116
        '
        'ColShortage
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColShortage.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColShortage.HeaderText = "Shortage"
        Me.ColShortage.Name = "ColShortage"
        Me.ColShortage.Width = 91
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.Width = 58
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(680, 475)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(466, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "The following product doesn't have adequate on hand qty, please check:"
        '
        'FrmTransaksiStockOutSales_QtyMinus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(767, 510)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.dgvMinus)
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiStockOutSales_QtyMinus"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Minus data detail"
        CType(Me.dgvMinus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvMinus As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWarehouse As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyStockOut As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyOnHand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColShortage As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
