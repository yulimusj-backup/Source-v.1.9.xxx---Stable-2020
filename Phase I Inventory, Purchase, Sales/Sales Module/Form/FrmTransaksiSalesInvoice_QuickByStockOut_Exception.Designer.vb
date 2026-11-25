<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSalesInvoice_QuickByStockOut_Exception
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNoOutSales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colPrimaryStockOutSales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnClose = New System.Windows.Forms.Button
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
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.ColNoOutSales, Me.colReason, Me.colPrimaryStockOutSales})
        Me.dgvItem.Location = New System.Drawing.Point(9, 10)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(524, 383)
        Me.dgvItem.TabIndex = 0
        '
        'colNo
        '
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        '
        'ColNoOutSales
        '
        Me.ColNoOutSales.HeaderText = "Stock Out Sales No"
        Me.ColNoOutSales.Name = "ColNoOutSales"
        Me.ColNoOutSales.ReadOnly = True
        '
        'colReason
        '
        Me.colReason.HeaderText = "Reason"
        Me.colReason.Name = "colReason"
        Me.colReason.ReadOnly = True
        '
        'colPrimaryStockOutSales
        '
        Me.colPrimaryStockOutSales.HeaderText = "Primary Stock Out Sales"
        Me.colPrimaryStockOutSales.Name = "colPrimaryStockOutSales"
        Me.colPrimaryStockOutSales.ReadOnly = True
        Me.colPrimaryStockOutSales.Visible = False
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(477, 400)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(56, 19)
        Me.BtnClose.TabIndex = 1
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSalesInvoice_QuickByStockOut_Exception
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(544, 427)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.dgvItem)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(560, 465)
        Me.Name = "FrmTransaksiSalesInvoice_QuickByStockOut_Exception"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exception"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNoOutSales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPrimaryStockOutSales As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
