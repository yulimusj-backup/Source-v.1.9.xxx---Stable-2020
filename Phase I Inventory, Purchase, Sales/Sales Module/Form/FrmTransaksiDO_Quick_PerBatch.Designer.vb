<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiDO_Quick_PerBatch
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
        Me.btnOK = New System.Windows.Forms.Button
        Me.Dgv = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyAvail = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyIn = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDecimal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.txtProductCode = New System.Windows.Forms.TextBox
        Me.txtDesciption = New System.Windows.Forms.TextBox
        Me.TxtQtyTotal = New System.Windows.Forms.TextBox
        Me.txtQtyTotalIn = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblTransactionNo = New System.Windows.Forms.Label
        Me.txtTransactionNo = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(508, 461)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(87, 23)
        Me.btnOK.TabIndex = 84
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Dgv
        '
        Me.Dgv.AllowUserToAddRows = False
        Me.Dgv.AllowUserToDeleteRows = False
        Me.Dgv.AllowUserToResizeColumns = False
        Me.Dgv.AllowUserToResizeRows = False
        Me.Dgv.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.Dgv.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColBatchNo, Me.ColExpiryDate, Me.ColQtyAvail, Me.ColQtyIn, Me.ColUnit, Me.ColDecimal})
        Me.Dgv.Location = New System.Drawing.Point(9, 99)
        Me.Dgv.Margin = New System.Windows.Forms.Padding(2)
        Me.Dgv.Name = "Dgv"
        Me.Dgv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.Dgv.RowTemplate.Height = 24
        Me.Dgv.Size = New System.Drawing.Size(587, 356)
        Me.Dgv.TabIndex = 85
        '
        'ColNo
        '
        Me.ColNo.FillWeight = 50.0!
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 50
        '
        'ColBatchNo
        '
        Me.ColBatchNo.HeaderText = "Batch No"
        Me.ColBatchNo.Name = "ColBatchNo"
        Me.ColBatchNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColExpiryDate
        '
        Me.ColExpiryDate.HeaderText = "Expiry Date"
        Me.ColExpiryDate.Name = "ColExpiryDate"
        Me.ColExpiryDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColQtyAvail
        '
        Me.ColQtyAvail.HeaderText = "Qty Available"
        Me.ColQtyAvail.Name = "ColQtyAvail"
        '
        'ColQtyIn
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyIn.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColQtyIn.HeaderText = "Qty"
        Me.ColQtyIn.Name = "ColQtyIn"
        Me.ColQtyIn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Units"
        Me.ColUnit.Name = "ColUnit"
        '
        'ColDecimal
        '
        Me.ColDecimal.HeaderText = "Decimal"
        Me.ColDecimal.Name = "ColDecimal"
        Me.ColDecimal.ReadOnly = True
        Me.ColDecimal.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'txtProductCode
        '
        Me.txtProductCode.Location = New System.Drawing.Point(90, 12)
        Me.txtProductCode.Name = "txtProductCode"
        Me.txtProductCode.ReadOnly = True
        Me.txtProductCode.Size = New System.Drawing.Size(191, 20)
        Me.txtProductCode.TabIndex = 87
        Me.txtProductCode.TabStop = False
        '
        'txtDesciption
        '
        Me.txtDesciption.Location = New System.Drawing.Point(90, 42)
        Me.txtDesciption.Name = "txtDesciption"
        Me.txtDesciption.ReadOnly = True
        Me.txtDesciption.Size = New System.Drawing.Size(191, 20)
        Me.txtDesciption.TabIndex = 88
        Me.txtDesciption.TabStop = False
        '
        'TxtQtyTotal
        '
        Me.TxtQtyTotal.Location = New System.Drawing.Point(376, 12)
        Me.TxtQtyTotal.MaxLength = 50
        Me.TxtQtyTotal.Name = "TxtQtyTotal"
        Me.TxtQtyTotal.ReadOnly = True
        Me.TxtQtyTotal.Size = New System.Drawing.Size(192, 20)
        Me.TxtQtyTotal.TabIndex = 89
        Me.TxtQtyTotal.TabStop = False
        Me.TxtQtyTotal.Text = "0"
        Me.TxtQtyTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtQtyTotalIn
        '
        Me.txtQtyTotalIn.Location = New System.Drawing.Point(376, 38)
        Me.txtQtyTotalIn.MaxLength = 50
        Me.txtQtyTotalIn.Name = "txtQtyTotalIn"
        Me.txtQtyTotalIn.ReadOnly = True
        Me.txtQtyTotalIn.Size = New System.Drawing.Size(192, 20)
        Me.txtQtyTotalIn.TabIndex = 90
        Me.txtQtyTotalIn.TabStop = False
        Me.txtQtyTotalIn.Text = "0"
        Me.txtQtyTotalIn.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 13)
        Me.Label1.TabIndex = 91
        Me.Label1.Text = "Product Code"
        '
        'lblTransactionNo
        '
        Me.lblTransactionNo.AutoSize = True
        Me.lblTransactionNo.Location = New System.Drawing.Point(12, 71)
        Me.lblTransactionNo.Name = "lblTransactionNo"
        Me.lblTransactionNo.Size = New System.Drawing.Size(62, 13)
        Me.lblTransactionNo.TabIndex = 92
        Me.lblTransactionNo.Text = "SO Number"
        '
        'txtTransactionNo
        '
        Me.txtTransactionNo.Location = New System.Drawing.Point(90, 68)
        Me.txtTransactionNo.Name = "txtTransactionNo"
        Me.txtTransactionNo.ReadOnly = True
        Me.txtTransactionNo.Size = New System.Drawing.Size(191, 20)
        Me.txtTransactionNo.TabIndex = 93
        Me.txtTransactionNo.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 13)
        Me.Label2.TabIndex = 94
        Me.Label2.Text = "Description"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(287, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 95
        Me.Label3.Text = "Total Qty"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(287, 45)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 13)
        Me.Label4.TabIndex = 96
        Me.Label4.Text = "Total Qty Input"
        '
        'FrmTransaksiDO_Quick_PerBatch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(605, 501)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtTransactionNo)
        Me.Controls.Add(Me.lblTransactionNo)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtQtyTotalIn)
        Me.Controls.Add(Me.TxtQtyTotal)
        Me.Controls.Add(Me.txtDesciption)
        Me.Controls.Add(Me.txtProductCode)
        Me.Controls.Add(Me.Dgv)
        Me.Controls.Add(Me.btnOK)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(613, 528)
        Me.Name = "FrmTransaksiDO_Quick_PerBatch"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Option Batch No"
        CType(Me.Dgv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Dgv As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyAvail As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyIn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDecimal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents txtProductCode As System.Windows.Forms.TextBox
    Friend WithEvents txtDesciption As System.Windows.Forms.TextBox
    Friend WithEvents TxtQtyTotal As System.Windows.Forms.TextBox
    Friend WithEvents txtQtyTotalIn As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblTransactionNo As System.Windows.Forms.Label
    Friend WithEvents txtTransactionNo As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
