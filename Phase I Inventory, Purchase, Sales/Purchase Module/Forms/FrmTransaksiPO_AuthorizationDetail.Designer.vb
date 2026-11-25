<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiPO_AuthorizationDetail
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
        Me.ColAuthorizeBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAuthorizationDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAuthorization = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button1 = New System.Windows.Forms.Button
        Me.dgvDetail = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.chkDP = New System.Windows.Forms.CheckBox
        Me.chkPOC = New System.Windows.Forms.CheckBox
        Me.chkPackingList = New System.Windows.Forms.CheckBox
        Me.chkStockIn = New System.Windows.Forms.CheckBox
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ColAuthorizeBy
        '
        Me.ColAuthorizeBy.HeaderText = "By"
        Me.ColAuthorizeBy.Name = "ColAuthorizeBy"
        Me.ColAuthorizeBy.ReadOnly = True
        Me.ColAuthorizeBy.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAuthorizeBy.Width = 30
        '
        'ColAuthorizationDate
        '
        Me.ColAuthorizationDate.HeaderText = "Authorization Date"
        Me.ColAuthorizationDate.Name = "ColAuthorizationDate"
        Me.ColAuthorizationDate.ReadOnly = True
        Me.ColAuthorizationDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAuthorizationDate.Width = 131
        '
        'ColAuthorization
        '
        Me.ColAuthorization.HeaderText = "Authorization"
        Me.ColAuthorization.Name = "ColAuthorization"
        Me.ColAuthorization.ReadOnly = True
        Me.ColAuthorization.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAuthorization.Width = 97
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button1.Location = New System.Drawing.Point(304, 106)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(153, 36)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dgvDetail
        '
        Me.dgvDetail.AllowUserToAddRows = False
        Me.dgvDetail.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.SkyBlue
        Me.dgvDetail.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvDetail.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColAuthorization, Me.ColAuthorizeBy, Me.ColAuthorizationDate})
        Me.dgvDetail.Location = New System.Drawing.Point(13, 13)
        Me.dgvDetail.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvDetail.MultiSelect = False
        Me.dgvDetail.Name = "dgvDetail"
        Me.dgvDetail.ReadOnly = True
        Me.dgvDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvDetail.RowTemplate.Height = 24
        Me.dgvDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvDetail.Size = New System.Drawing.Size(587, 168)
        Me.dgvDetail.TabIndex = 2
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'chkDP
        '
        Me.chkDP.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.chkDP.AutoSize = True
        Me.chkDP.Enabled = False
        Me.chkDP.Location = New System.Drawing.Point(492, 188)
        Me.chkDP.Name = "chkDP"
        Me.chkDP.Size = New System.Drawing.Size(113, 21)
        Me.chkDP.TabIndex = 9
        Me.chkDP.Text = "Purchase DP"
        Me.chkDP.UseVisualStyleBackColor = True
        '
        'chkPOC
        '
        Me.chkPOC.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.chkPOC.AutoSize = True
        Me.chkPOC.Enabled = False
        Me.chkPOC.Location = New System.Drawing.Point(274, 188)
        Me.chkPOC.Name = "chkPOC"
        Me.chkPOC.Size = New System.Drawing.Size(212, 21)
        Me.chkPOC.TabIndex = 8
        Me.chkPOC.Text = "Purchase Order Cancellation"
        Me.chkPOC.UseVisualStyleBackColor = True
        '
        'chkPackingList
        '
        Me.chkPackingList.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.chkPackingList.AutoSize = True
        Me.chkPackingList.Enabled = False
        Me.chkPackingList.Location = New System.Drawing.Point(162, 188)
        Me.chkPackingList.Name = "chkPackingList"
        Me.chkPackingList.Size = New System.Drawing.Size(106, 21)
        Me.chkPackingList.TabIndex = 7
        Me.chkPackingList.Text = "Packing List"
        Me.chkPackingList.UseVisualStyleBackColor = True
        '
        'chkStockIn
        '
        Me.chkStockIn.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.chkStockIn.AutoSize = True
        Me.chkStockIn.Enabled = False
        Me.chkStockIn.Location = New System.Drawing.Point(12, 188)
        Me.chkStockIn.Name = "chkStockIn"
        Me.chkStockIn.Size = New System.Drawing.Size(144, 21)
        Me.chkStockIn.TabIndex = 6
        Me.chkStockIn.Text = "Stock In Purchase"
        Me.chkStockIn.UseVisualStyleBackColor = True
        '
        'FrmTransaksiPO_AuthorizationDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(613, 217)
        Me.Controls.Add(Me.chkDP)
        Me.Controls.Add(Me.chkPOC)
        Me.Controls.Add(Me.chkPackingList)
        Me.Controls.Add(Me.chkStockIn)
        Me.Controls.Add(Me.dgvDetail)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiPO_AuthorizationDetail"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Purchase Order Authorization Detail"
        CType(Me.dgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ColAuthorizeBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAuthorizationDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAuthorization As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents dgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents chkDP As System.Windows.Forms.CheckBox
    Friend WithEvents chkPOC As System.Windows.Forms.CheckBox
    Friend WithEvents chkPackingList As System.Windows.Forms.CheckBox
    Friend WithEvents chkStockIn As System.Windows.Forms.CheckBox
End Class
