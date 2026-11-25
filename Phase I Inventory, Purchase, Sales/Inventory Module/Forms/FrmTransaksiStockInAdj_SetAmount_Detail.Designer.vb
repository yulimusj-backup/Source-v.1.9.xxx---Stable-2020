<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiStockInAdj_SetAmount_Detail
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
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.ColDetail_No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_DocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_Date = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_Qty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_Unit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_UnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_PrimaryProd = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDetail_PrimaryTrans = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDetail
        '
        Me.dgvDetail.AllowUserToAddRows = False
        Me.dgvDetail.AllowUserToDeleteRows = False
        Me.dgvDetail.AllowUserToResizeColumns = False
        Me.dgvDetail.AllowUserToResizeRows = False
        Me.dgvDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetail.ColumnHeadersHeight = 25
        Me.dgvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDetail_No, Me.ColDetail_DocNo, Me.ColDetail_Date, Me.ColDetail_Qty, Me.ColDetail_Unit, Me.ColDetail_UnitPrice, Me.ColDetail_Amount, Me.ColDetail_PrimaryProd, Me.ColDetail_PrimaryTrans})
        Me.dgvDetail.Location = New System.Drawing.Point(9, 10)
        Me.dgvDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.ReadOnly = True
        Me.dgvDetail.RowHeadersWidth = 24
        Me.dgvDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDetail.RowTemplate.Height = 24
        Me.dgvDetail.Size = New System.Drawing.Size(576, 275)
        Me.dgvDetail.TabIndex = 2
        '
        'ColDetail_No
        '
        Me.ColDetail_No.HeaderText = "No"
        Me.ColDetail_No.Name = "ColDetail_No"
        Me.ColDetail_No.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_No.Width = 27
        '
        'ColDetail_DocNo
        '
        Me.ColDetail_DocNo.HeaderText = "Document No"
        Me.ColDetail_DocNo.Name = "ColDetail_DocNo"
        Me.ColDetail_DocNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_DocNo.Width = 79
        '
        'ColDetail_Date
        '
        Me.ColDetail_Date.HeaderText = "Date"
        Me.ColDetail_Date.Name = "ColDetail_Date"
        Me.ColDetail_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_Date.Width = 36
        '
        'ColDetail_Qty
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDetail_Qty.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColDetail_Qty.HeaderText = "Qty"
        Me.ColDetail_Qty.Name = "ColDetail_Qty"
        Me.ColDetail_Qty.ReadOnly = True
        Me.ColDetail_Qty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_Qty.Width = 29
        '
        'ColDetail_Unit
        '
        Me.ColDetail_Unit.HeaderText = "Unit"
        Me.ColDetail_Unit.Name = "ColDetail_Unit"
        Me.ColDetail_Unit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_Unit.Width = 32
        '
        'ColDetail_UnitPrice
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDetail_UnitPrice.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColDetail_UnitPrice.HeaderText = "Unit Price"
        Me.ColDetail_UnitPrice.Name = "ColDetail_UnitPrice"
        Me.ColDetail_UnitPrice.ReadOnly = True
        Me.ColDetail_UnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_UnitPrice.Width = 59
        '
        'ColDetail_Amount
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDetail_Amount.DefaultCellStyle = DataGridViewCellStyle15
        Me.ColDetail_Amount.HeaderText = "Amount"
        Me.ColDetail_Amount.Name = "ColDetail_Amount"
        Me.ColDetail_Amount.ReadOnly = True
        Me.ColDetail_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_Amount.Width = 49
        '
        'ColDetail_PrimaryProd
        '
        Me.ColDetail_PrimaryProd.HeaderText = "Primary Prod"
        Me.ColDetail_PrimaryProd.Name = "ColDetail_PrimaryProd"
        Me.ColDetail_PrimaryProd.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_PrimaryProd.Visible = False
        Me.ColDetail_PrimaryProd.Width = 72
        '
        'ColDetail_PrimaryTrans
        '
        Me.ColDetail_PrimaryTrans.HeaderText = "Primary Trans"
        Me.ColDetail_PrimaryTrans.Name = "ColDetail_PrimaryTrans"
        Me.ColDetail_PrimaryTrans.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDetail_PrimaryTrans.Visible = False
        Me.ColDetail_PrimaryTrans.Width = 77
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(362, 89)
        Me.btnClose.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(56, 19)
        Me.btnClose.TabIndex = 3
        Me.btnClose.TabStop = False
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmTransaksiStockInAdj_SetAmount_Detail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(594, 295)
        Me.Controls.Add(Me.dgvDetail)
        Me.Controls.Add(Me.btnClose)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(610, 334)
        Me.Name = "FrmTransaksiStockInAdj_SetAmount_Detail"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Detail Transaction"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents ColDetail_No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_DocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_Date As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_Qty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_Unit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_UnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_PrimaryProd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDetail_PrimaryTrans As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
