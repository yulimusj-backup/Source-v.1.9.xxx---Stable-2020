<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Session_View
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
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBatchNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDisc_Amount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNettPrice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSubTotal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUnit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDisc3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscAmount3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLocation = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDiscString3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAutoNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnExit = New System.Windows.Forms.Button
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColCodeNo, Me.ColDesc, Me.ColLocation, Me.ColBatchNumber, Me.ColExpiry, Me.ColQty, Me.ColUnit, Me.ColUnitPrice, Me.ColDiscString, Me.ColDisc_Amount, Me.ColNettPrice, Me.ColSubTotal, Me.ColPrimaryCodeNo, Me.ColPrimaryUnit, Me.ColPrimaryDisc1, Me.ColDiscAmount1, Me.ColPrimaryDisc2, Me.ColDiscAmount2, Me.ColPrimaryDisc3, Me.ColDiscAmount3, Me.ColPrimaryLocation, Me.ColDiscString1, Me.ColDiscString2, Me.ColDiscString3, Me.ColPrimaryAutoNumber})
        Me.dgvItem.Location = New System.Drawing.Point(12, 5)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvItem.Size = New System.Drawing.Size(657, 347)
        Me.dgvItem.StandardTab = True
        Me.dgvItem.TabIndex = 3
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColCodeNo
        '
        Me.ColCodeNo.HeaderText = "Code No"
        Me.ColCodeNo.Name = "ColCodeNo"
        Me.ColCodeNo.ReadOnly = True
        Me.ColCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCodeNo.Width = 55
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'ColLocation
        '
        Me.ColLocation.HeaderText = "Location"
        Me.ColLocation.Name = "ColLocation"
        Me.ColLocation.ReadOnly = True
        Me.ColLocation.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColLocation.Visible = False
        Me.ColLocation.Width = 54
        '
        'ColBatchNumber
        '
        Me.ColBatchNumber.HeaderText = "Lot No"
        Me.ColBatchNumber.Name = "ColBatchNumber"
        Me.ColBatchNumber.ReadOnly = True
        Me.ColBatchNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColBatchNumber.Visible = False
        Me.ColBatchNumber.Width = 45
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.ReadOnly = True
        Me.ColExpiry.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExpiry.Visible = False
        Me.ColExpiry.Width = 41
        '
        'ColQty
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColQty.HeaderText = "Qty"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Width = 29
        '
        'ColUnit
        '
        Me.ColUnit.HeaderText = "Unit"
        Me.ColUnit.Name = "ColUnit"
        Me.ColUnit.ReadOnly = True
        Me.ColUnit.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnit.Width = 32
        '
        'ColUnitPrice
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColUnitPrice.DefaultCellStyle = DataGridViewCellStyle12
        Me.ColUnitPrice.HeaderText = "Unit Price"
        Me.ColUnitPrice.Name = "ColUnitPrice"
        Me.ColUnitPrice.ReadOnly = True
        Me.ColUnitPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColUnitPrice.Width = 59
        '
        'ColDiscString
        '
        Me.ColDiscString.HeaderText = "Discount String"
        Me.ColDiscString.Name = "ColDiscString"
        Me.ColDiscString.ReadOnly = True
        Me.ColDiscString.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDiscString.Visible = False
        Me.ColDiscString.Width = 85
        '
        'ColDisc_Amount
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDisc_Amount.DefaultCellStyle = DataGridViewCellStyle13
        Me.ColDisc_Amount.HeaderText = "Discount"
        Me.ColDisc_Amount.Name = "ColDisc_Amount"
        Me.ColDisc_Amount.ReadOnly = True
        Me.ColDisc_Amount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDisc_Amount.Width = 55
        '
        'ColNettPrice
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColNettPrice.DefaultCellStyle = DataGridViewCellStyle14
        Me.ColNettPrice.HeaderText = "Nett Price"
        Me.ColNettPrice.Name = "ColNettPrice"
        Me.ColNettPrice.ReadOnly = True
        Me.ColNettPrice.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNettPrice.Width = 60
        '
        'ColSubTotal
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSubTotal.DefaultCellStyle = DataGridViewCellStyle15
        Me.ColSubTotal.HeaderText = "Sub Total"
        Me.ColSubTotal.Name = "ColSubTotal"
        Me.ColSubTotal.ReadOnly = True
        Me.ColSubTotal.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSubTotal.Width = 59
        '
        'ColPrimaryCodeNo
        '
        Me.ColPrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColPrimaryCodeNo.Name = "ColPrimaryCodeNo"
        Me.ColPrimaryCodeNo.ReadOnly = True
        Me.ColPrimaryCodeNo.Visible = False
        Me.ColPrimaryCodeNo.Width = 111
        '
        'ColPrimaryUnit
        '
        Me.ColPrimaryUnit.HeaderText = "Primary Unit"
        Me.ColPrimaryUnit.Name = "ColPrimaryUnit"
        Me.ColPrimaryUnit.ReadOnly = True
        Me.ColPrimaryUnit.Visible = False
        Me.ColPrimaryUnit.Width = 88
        '
        'ColPrimaryDisc1
        '
        Me.ColPrimaryDisc1.HeaderText = "PrimaryDisc1"
        Me.ColPrimaryDisc1.Name = "ColPrimaryDisc1"
        Me.ColPrimaryDisc1.ReadOnly = True
        Me.ColPrimaryDisc1.Visible = False
        Me.ColPrimaryDisc1.Width = 93
        '
        'ColDiscAmount1
        '
        Me.ColDiscAmount1.HeaderText = "DiscAmount1"
        Me.ColDiscAmount1.Name = "ColDiscAmount1"
        Me.ColDiscAmount1.ReadOnly = True
        Me.ColDiscAmount1.Visible = False
        Me.ColDiscAmount1.Width = 95
        '
        'ColPrimaryDisc2
        '
        Me.ColPrimaryDisc2.HeaderText = "Primary Disc 2"
        Me.ColPrimaryDisc2.Name = "ColPrimaryDisc2"
        Me.ColPrimaryDisc2.ReadOnly = True
        Me.ColPrimaryDisc2.Visible = False
        Me.ColPrimaryDisc2.Width = 99
        '
        'ColDiscAmount2
        '
        Me.ColDiscAmount2.HeaderText = "Disc Amount2"
        Me.ColDiscAmount2.Name = "ColDiscAmount2"
        Me.ColDiscAmount2.ReadOnly = True
        Me.ColDiscAmount2.Visible = False
        Me.ColDiscAmount2.Width = 98
        '
        'ColPrimaryDisc3
        '
        Me.ColPrimaryDisc3.HeaderText = "Primary Disc 3"
        Me.ColPrimaryDisc3.Name = "ColPrimaryDisc3"
        Me.ColPrimaryDisc3.ReadOnly = True
        Me.ColPrimaryDisc3.Visible = False
        Me.ColPrimaryDisc3.Width = 99
        '
        'ColDiscAmount3
        '
        Me.ColDiscAmount3.HeaderText = "Disc Amount 3"
        Me.ColDiscAmount3.Name = "ColDiscAmount3"
        Me.ColDiscAmount3.ReadOnly = True
        Me.ColDiscAmount3.Visible = False
        Me.ColDiscAmount3.Width = 101
        '
        'ColPrimaryLocation
        '
        Me.ColPrimaryLocation.HeaderText = "Primary Location"
        Me.ColPrimaryLocation.Name = "ColPrimaryLocation"
        Me.ColPrimaryLocation.ReadOnly = True
        Me.ColPrimaryLocation.Visible = False
        Me.ColPrimaryLocation.Width = 110
        '
        'ColDiscString1
        '
        Me.ColDiscString1.HeaderText = "DiscString1"
        Me.ColDiscString1.Name = "ColDiscString1"
        Me.ColDiscString1.ReadOnly = True
        Me.ColDiscString1.Visible = False
        Me.ColDiscString1.Width = 86
        '
        'ColDiscString2
        '
        Me.ColDiscString2.HeaderText = "DiscString2"
        Me.ColDiscString2.Name = "ColDiscString2"
        Me.ColDiscString2.ReadOnly = True
        Me.ColDiscString2.Visible = False
        Me.ColDiscString2.Width = 86
        '
        'ColDiscString3
        '
        Me.ColDiscString3.HeaderText = "DiscString3"
        Me.ColDiscString3.Name = "ColDiscString3"
        Me.ColDiscString3.ReadOnly = True
        Me.ColDiscString3.Visible = False
        Me.ColDiscString3.Width = 86
        '
        'ColPrimaryAutoNumber
        '
        Me.ColPrimaryAutoNumber.HeaderText = "PrimaryAutoNumber"
        Me.ColPrimaryAutoNumber.Name = "ColPrimaryAutoNumber"
        Me.ColPrimaryAutoNumber.ReadOnly = True
        Me.ColPrimaryAutoNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryAutoNumber.Visible = False
        Me.ColPrimaryAutoNumber.Width = 106
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(541, 175)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(110, 26)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'FrmTrans_Session_View
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(681, 364)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.btnExit)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_Session_View"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "View Data"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBatchNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDisc_Amount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNettPrice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSubTotal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUnit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDisc3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscAmount3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLocation As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDiscString3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAutoNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnExit As System.Windows.Forms.Button
End Class
