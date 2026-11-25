<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_CutOffDB_ExceptionQty
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
        Me.DgvQty = New System.Windows.Forms.DataGridView
        Me.ColPrimaryProd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProdCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProdQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProdQty_StockCard = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvQty, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvQty
        '
        Me.DgvQty.AllowUserToAddRows = False
        Me.DgvQty.AllowUserToDeleteRows = False
        Me.DgvQty.AllowUserToResizeColumns = False
        Me.DgvQty.AllowUserToResizeRows = False
        Me.DgvQty.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvQty.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvQty.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvQty.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvQty.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPrimaryProd, Me.ColProdCode, Me.ColProdQty, Me.ColProdQty_StockCard})
        Me.DgvQty.Location = New System.Drawing.Point(12, 12)
        Me.DgvQty.Name = "DgvQty"
        Me.DgvQty.ReadOnly = True
        Me.DgvQty.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvQty.Size = New System.Drawing.Size(518, 249)
        Me.DgvQty.TabIndex = 0
        '
        'ColPrimaryProd
        '
        Me.ColPrimaryProd.HeaderText = "Primary Product"
        Me.ColPrimaryProd.Name = "ColPrimaryProd"
        Me.ColPrimaryProd.ReadOnly = True
        Me.ColPrimaryProd.Visible = False
        Me.ColPrimaryProd.Width = 106
        '
        'ColProdCode
        '
        Me.ColProdCode.HeaderText = "Product Code"
        Me.ColProdCode.Name = "ColProdCode"
        Me.ColProdCode.ReadOnly = True
        Me.ColProdCode.Width = 97
        '
        'ColProdQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColProdQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColProdQty.HeaderText = "Product Qty"
        Me.ColProdQty.Name = "ColProdQty"
        Me.ColProdQty.ReadOnly = True
        Me.ColProdQty.Width = 88
        '
        'ColProdQty_StockCard
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColProdQty_StockCard.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColProdQty_StockCard.HeaderText = "Stock Card Qty"
        Me.ColProdQty_StockCard.Name = "ColProdQty_StockCard"
        Me.ColProdQty_StockCard.ReadOnly = True
        Me.ColProdQty_StockCard.Width = 104
        '
        'FrmTrans_CutOffDB_ExceptionQty
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(542, 273)
        Me.Controls.Add(Me.DgvQty)
        Me.MinimumSize = New System.Drawing.Size(550, 300)
        Me.Name = "FrmTrans_CutOffDB_ExceptionQty"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product Qty & Stock Card"
        CType(Me.DgvQty, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvQty As System.Windows.Forms.DataGridView
    Friend WithEvents ColPrimaryProd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProdCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProdQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProdQty_StockCard As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
