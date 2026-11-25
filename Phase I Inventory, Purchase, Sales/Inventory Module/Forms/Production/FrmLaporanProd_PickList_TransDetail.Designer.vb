<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProd_PickList_TransDetail
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
        Me.DgvTransferDetail = New System.Windows.Forms.DataGridView
        Me.BtnOK = New System.Windows.Forms.Button
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.ColTranDet_CodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_LocationFrom = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransDet_LocationTo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_Batch = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_ExpDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_QtyOnHand = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_QtyTrans = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_PrimaryCodeNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_PrimaryLocFrom = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTranDet_PrimaryLocTo = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvTransferDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvTransferDetail
        '
        Me.DgvTransferDetail.AllowUserToAddRows = False
        Me.DgvTransferDetail.AllowUserToDeleteRows = False
        Me.DgvTransferDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvTransferDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvTransferDetail.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvTransferDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColTranDet_CodeNo, Me.ColTranDet_LocationFrom, Me.ColTransDet_LocationTo, Me.ColTranDet_Batch, Me.ColTranDet_ExpDate, Me.ColTranDet_QtyOnHand, Me.ColTranDet_QtyTrans, Me.ColTranDet_PrimaryCodeNo, Me.ColTranDet_PrimaryLocFrom, Me.ColTranDet_PrimaryLocTo})
        Me.DgvTransferDetail.Location = New System.Drawing.Point(4, 5)
        Me.DgvTransferDetail.MultiSelect = False
        Me.DgvTransferDetail.Name = "DgvTransferDetail"
        Me.DgvTransferDetail.Size = New System.Drawing.Size(588, 187)
        Me.DgvTransferDetail.TabIndex = 0
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(479, 198)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(53, 23)
        Me.BtnOK.TabIndex = 1
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(538, 198)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(53, 23)
        Me.BtnCancel.TabIndex = 2
        Me.BtnCancel.Text = "Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'ColTranDet_CodeNo
        '
        Me.ColTranDet_CodeNo.HeaderText = "Code No"
        Me.ColTranDet_CodeNo.Name = "ColTranDet_CodeNo"
        Me.ColTranDet_CodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_CodeNo.Width = 55
        '
        'ColTranDet_LocationFrom
        '
        Me.ColTranDet_LocationFrom.HeaderText = "Location From"
        Me.ColTranDet_LocationFrom.Name = "ColTranDet_LocationFrom"
        Me.ColTranDet_LocationFrom.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_LocationFrom.Width = 80
        '
        'ColTransDet_LocationTo
        '
        Me.ColTransDet_LocationTo.HeaderText = "Location To"
        Me.ColTransDet_LocationTo.Name = "ColTransDet_LocationTo"
        Me.ColTransDet_LocationTo.ReadOnly = True
        Me.ColTransDet_LocationTo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTransDet_LocationTo.Width = 70
        '
        'ColTranDet_Batch
        '
        Me.ColTranDet_Batch.HeaderText = "Batch"
        Me.ColTranDet_Batch.Name = "ColTranDet_Batch"
        Me.ColTranDet_Batch.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_Batch.Width = 41
        '
        'ColTranDet_ExpDate
        '
        Me.ColTranDet_ExpDate.HeaderText = "Expired Date"
        Me.ColTranDet_ExpDate.Name = "ColTranDet_ExpDate"
        Me.ColTranDet_ExpDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_ExpDate.Width = 74
        '
        'ColTranDet_QtyOnHand
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTranDet_QtyOnHand.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColTranDet_QtyOnHand.HeaderText = "Qty On Hand"
        Me.ColTranDet_QtyOnHand.Name = "ColTranDet_QtyOnHand"
        Me.ColTranDet_QtyOnHand.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_QtyOnHand.Width = 75
        '
        'ColTranDet_QtyTrans
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTranDet_QtyTrans.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColTranDet_QtyTrans.HeaderText = "Qty Transfer"
        Me.ColTranDet_QtyTrans.Name = "ColTranDet_QtyTrans"
        Me.ColTranDet_QtyTrans.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_QtyTrans.Width = 71
        '
        'ColTranDet_PrimaryCodeNo
        '
        Me.ColTranDet_PrimaryCodeNo.HeaderText = "Primary Code No"
        Me.ColTranDet_PrimaryCodeNo.Name = "ColTranDet_PrimaryCodeNo"
        Me.ColTranDet_PrimaryCodeNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_PrimaryCodeNo.Visible = False
        Me.ColTranDet_PrimaryCodeNo.Width = 92
        '
        'ColTranDet_PrimaryLocFrom
        '
        Me.ColTranDet_PrimaryLocFrom.HeaderText = "Primary Location From"
        Me.ColTranDet_PrimaryLocFrom.Name = "ColTranDet_PrimaryLocFrom"
        Me.ColTranDet_PrimaryLocFrom.ReadOnly = True
        Me.ColTranDet_PrimaryLocFrom.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_PrimaryLocFrom.Visible = False
        Me.ColTranDet_PrimaryLocFrom.Width = 117
        '
        'ColTranDet_PrimaryLocTo
        '
        Me.ColTranDet_PrimaryLocTo.HeaderText = "Primary Location To"
        Me.ColTranDet_PrimaryLocTo.Name = "ColTranDet_PrimaryLocTo"
        Me.ColTranDet_PrimaryLocTo.ReadOnly = True
        Me.ColTranDet_PrimaryLocTo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTranDet_PrimaryLocTo.Visible = False
        Me.ColTranDet_PrimaryLocTo.Width = 107
        '
        'FrmLaporanProd_PickList_TransDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(596, 227)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.DgvTransferDetail)
        Me.MinimumSize = New System.Drawing.Size(439, 252)
        Me.Name = "FrmLaporanProd_PickList_TransDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transfer Detail"
        CType(Me.DgvTransferDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvTransferDetail As System.Windows.Forms.DataGridView
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents ColTranDet_CodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_LocationFrom As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransDet_LocationTo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_Batch As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_ExpDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_QtyOnHand As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_QtyTrans As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_PrimaryCodeNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_PrimaryLocFrom As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTranDet_PrimaryLocTo As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
