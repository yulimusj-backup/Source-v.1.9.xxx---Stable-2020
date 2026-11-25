<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKondisiStock_PickUpOrderList
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
        Me.DgvDetail = New System.Windows.Forms.DataGridView
        Me.Col_PrimaryPickUpOrder = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Col_PickUpOrderNo = New System.Windows.Forms.DataGridViewLinkColumn
        Me.BtnClose = New System.Windows.Forms.Button
        CType(Me.DgvDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvDetail
        '
        Me.DgvDetail.AllowUserToAddRows = False
        Me.DgvDetail.AllowUserToDeleteRows = False
        Me.DgvDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col_PrimaryPickUpOrder, Me.Col_PickUpOrderNo})
        Me.DgvDetail.Location = New System.Drawing.Point(12, 12)
        Me.DgvDetail.MultiSelect = False
        Me.DgvDetail.Name = "DgvDetail"
        Me.DgvDetail.ReadOnly = True
        Me.DgvDetail.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvDetail.Size = New System.Drawing.Size(255, 214)
        Me.DgvDetail.TabIndex = 0
        '
        'Col_PrimaryPickUpOrder
        '
        Me.Col_PrimaryPickUpOrder.HeaderText = "Primary Pick Up Order"
        Me.Col_PrimaryPickUpOrder.Name = "Col_PrimaryPickUpOrder"
        Me.Col_PrimaryPickUpOrder.ReadOnly = True
        Me.Col_PrimaryPickUpOrder.Visible = False
        '
        'Col_PickUpOrderNo
        '
        Me.Col_PickUpOrderNo.HeaderText = "Pick Up Order No"
        Me.Col_PickUpOrderNo.Name = "Col_PickUpOrderNo"
        Me.Col_PickUpOrderNo.ReadOnly = True
        Me.Col_PickUpOrderNo.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Col_PickUpOrderNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Col_PickUpOrderNo.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(116, 76)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 1
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmMasterKondisiStock_PickUpOrderList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(278, 238)
        Me.Controls.Add(Me.DgvDetail)
        Me.Controls.Add(Me.BtnClose)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(294, 276)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(294, 276)
        Me.Name = "FrmMasterKondisiStock_PickUpOrderList"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pick Up Order List"
        CType(Me.DgvDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvDetail As System.Windows.Forms.DataGridView
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents Col_PrimaryPickUpOrder As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Col_PickUpOrderNo As System.Windows.Forms.DataGridViewLinkColumn
End Class
