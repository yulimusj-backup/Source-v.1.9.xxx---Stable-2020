<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiQuotation_SetDelivery
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
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPlusPlus = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColPriorityLevel = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColPrimaryPriority = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvDetail
        '
        Me.dgvDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDeliveryDate, Me.ColQtyMulti, Me.ColPlusPlus, Me.ColPriorityLevel, Me.ColPrimaryPriority, Me.ColRemarks, Me.ColQty})
        Me.dgvDetail.Location = New System.Drawing.Point(10, 9)
        Me.dgvDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDetail.RowTemplate.Height = 24
        Me.dgvDetail.Size = New System.Drawing.Size(403, 168)
        Me.dgvDetail.TabIndex = 3
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(296, 184)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 21)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(357, 184)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 21)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'BackgroundWorker2
        '
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColDeliveryDate
        '
        Me.ColDeliveryDate.HeaderText = "Delivery Date"
        Me.ColDeliveryDate.Name = "ColDeliveryDate"
        Me.ColDeliveryDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDeliveryDate.Width = 77
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColPlusPlus
        '
        Me.ColPlusPlus.HeaderText = "Set"
        Me.ColPlusPlus.Name = "ColPlusPlus"
        Me.ColPlusPlus.Text = ""
        Me.ColPlusPlus.Width = 29
        '
        'ColPriorityLevel
        '
        Me.ColPriorityLevel.HeaderText = "Priority"
        Me.ColPriorityLevel.Name = "ColPriorityLevel"
        Me.ColPriorityLevel.Width = 44
        '
        'ColPrimaryPriority
        '
        Me.ColPrimaryPriority.HeaderText = "Primary Priority"
        Me.ColPrimaryPriority.Name = "ColPrimaryPriority"
        Me.ColPrimaryPriority.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPrimaryPriority.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryPriority.Visible = False
        Me.ColPrimaryPriority.Width = 81
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty Default"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.ReadOnly = True
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Visible = False
        Me.ColQty.Width = 66
        '
        'FrmTransaksiQuotation_SetDelivery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(423, 214)
        Me.Controls.Add(Me.dgvDetail)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(433, 251)
        Me.Name = "FrmTransaksiQuotation_SetDelivery"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Delivery Detail"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPlusPlus As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColPriorityLevel As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColPrimaryPriority As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
