<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPO_SetDeliveryAll
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
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDeliveryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDelay = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColQtyMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPlusPlus = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.chkUseDeliveryDate = New System.Windows.Forms.CheckBox
        Me.GboDeliveryDate = New System.Windows.Forms.GroupBox
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GboDeliveryDate.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvDetail
        '
        Me.dgvDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColDeliveryDate, Me.ColDelay, Me.ColQty, Me.ColQtyMulti, Me.ColPlusPlus, Me.ColRemarks})
        Me.dgvDetail.Location = New System.Drawing.Point(6, 18)
        Me.dgvDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.RowTemplate.Height = 24
        Me.dgvDetail.Size = New System.Drawing.Size(612, 239)
        Me.dgvDetail.TabIndex = 0
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
        'ColDelay
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDelay.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColDelay.HeaderText = "Preparation Delay (days)"
        Me.ColDelay.Name = "ColDelay"
        Me.ColDelay.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColDelay.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDelay.Width = 128
        '
        'ColQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColQty.HeaderText = "Qty Default"
        Me.ColQty.Name = "ColQty"
        Me.ColQty.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColQty.Visible = False
        Me.ColQty.Width = 66
        '
        'ColQtyMulti
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopRight
        Me.ColQtyMulti.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColQtyMulti.HeaderText = "Qty"
        Me.ColQtyMulti.Name = "ColQtyMulti"
        Me.ColQtyMulti.Width = 48
        '
        'ColPlusPlus
        '
        Me.ColPlusPlus.HeaderText = "Set"
        Me.ColPlusPlus.Name = "ColPlusPlus"
        Me.ColPlusPlus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPlusPlus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.ColPlusPlus.Width = 48
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 55
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(576, 274)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 22)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(515, 274)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(2)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 22)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'chkUseDeliveryDate
        '
        Me.chkUseDeliveryDate.AutoSize = True
        Me.chkUseDeliveryDate.Location = New System.Drawing.Point(6, 0)
        Me.chkUseDeliveryDate.Name = "chkUseDeliveryDate"
        Me.chkUseDeliveryDate.Size = New System.Drawing.Size(112, 17)
        Me.chkUseDeliveryDate.TabIndex = 167
        Me.chkUseDeliveryDate.Text = "Use Delivery Date"
        Me.chkUseDeliveryDate.UseVisualStyleBackColor = True
        '
        'GboDeliveryDate
        '
        Me.GboDeliveryDate.Controls.Add(Me.chkUseDeliveryDate)
        Me.GboDeliveryDate.Controls.Add(Me.dgvDetail)
        Me.GboDeliveryDate.Location = New System.Drawing.Point(9, 7)
        Me.GboDeliveryDate.Name = "GboDeliveryDate"
        Me.GboDeliveryDate.Size = New System.Drawing.Size(623, 262)
        Me.GboDeliveryDate.TabIndex = 168
        Me.GboDeliveryDate.TabStop = False
        '
        'FrmTransaksiPO_SetDeliveryAll
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(641, 303)
        Me.Controls.Add(Me.GboDeliveryDate)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(431, 251)
        Me.Name = "FrmTransaksiPO_SetDeliveryAll"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set delivery date detail for all item"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GboDeliveryDate.ResumeLayout(False)
        Me.GboDeliveryDate.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDeliveryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDelay As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColQtyMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPlusPlus As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkUseDeliveryDate As System.Windows.Forms.CheckBox
    Friend WithEvents GboDeliveryDate As System.Windows.Forms.GroupBox
End Class
