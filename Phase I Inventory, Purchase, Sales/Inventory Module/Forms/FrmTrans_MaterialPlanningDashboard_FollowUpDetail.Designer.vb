<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_MaterialPlanningDashboard_FollowUpDetail
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.BtnClose = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColDocumentNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColJONo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDefaultQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitDefault = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColMultiQty = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUnitMulti = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTimestamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryPO_MO = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryProduct = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColType, Me.ColDocumentNo, Me.ColDate, Me.ColJONo, Me.ColDefaultQty, Me.ColUnitDefault, Me.ColMultiQty, Me.ColUnitMulti, Me.ColUserCode, Me.ColTimestamp, Me.ColRemarks, Me.ColPrimaryPO_MO, Me.ColPrimaryProduct})
        Me.DgvItem.Location = New System.Drawing.Point(13, 5)
        Me.DgvItem.Margin = New System.Windows.Forms.Padding(4)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.ReadOnly = True
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvItem.Size = New System.Drawing.Size(926, 377)
        Me.DgvItem.TabIndex = 0
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(848, 388)
        Me.BtnClose.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(100, 26)
        Me.BtnClose.TabIndex = 1
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.Frozen = True
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColType
        '
        Me.ColType.Frozen = True
        Me.ColType.HeaderText = "Type"
        Me.ColType.Items.AddRange(New Object() {"Create new PO", "Add to PO", "Create new MO", "Add to MO"})
        Me.ColType.Name = "ColType"
        Me.ColType.ReadOnly = True
        Me.ColType.Width = 46
        '
        'ColDocumentNo
        '
        Me.ColDocumentNo.Frozen = True
        Me.ColDocumentNo.HeaderText = "Document No"
        Me.ColDocumentNo.Name = "ColDocumentNo"
        Me.ColDocumentNo.ReadOnly = True
        Me.ColDocumentNo.Width = 119
        '
        'ColDate
        '
        Me.ColDate.Frozen = True
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate.Width = 44
        '
        'ColJONo
        '
        Me.ColJONo.HeaderText = "JO No"
        Me.ColJONo.Name = "ColJONo"
        Me.ColJONo.ReadOnly = True
        Me.ColJONo.Width = 73
        '
        'ColDefaultQty
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDefaultQty.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColDefaultQty.HeaderText = "Qty"
        Me.ColDefaultQty.Name = "ColDefaultQty"
        Me.ColDefaultQty.ReadOnly = True
        Me.ColDefaultQty.Width = 55
        '
        'ColUnitDefault
        '
        Me.ColUnitDefault.HeaderText = "Unit"
        Me.ColUnitDefault.Name = "ColUnitDefault"
        Me.ColUnitDefault.ReadOnly = True
        Me.ColUnitDefault.Width = 58
        '
        'ColMultiQty
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColMultiQty.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColMultiQty.HeaderText = "Qty Multi"
        Me.ColMultiQty.Name = "ColMultiQty"
        Me.ColMultiQty.ReadOnly = True
        Me.ColMultiQty.Width = 88
        '
        'ColUnitMulti
        '
        Me.ColUnitMulti.HeaderText = "Multi Unit"
        Me.ColUnitMulti.Name = "ColUnitMulti"
        Me.ColUnitMulti.ReadOnly = True
        Me.ColUnitMulti.Width = 91
        '
        'ColUserCode
        '
        Me.ColUserCode.HeaderText = "User Follow Up"
        Me.ColUserCode.Name = "ColUserCode"
        Me.ColUserCode.ReadOnly = True
        Me.ColUserCode.Width = 128
        '
        'ColTimestamp
        '
        Me.ColTimestamp.HeaderText = "Follow Up Time"
        Me.ColTimestamp.Name = "ColTimestamp"
        Me.ColTimestamp.ReadOnly = True
        Me.ColTimestamp.Width = 129
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.Width = 89
        '
        'ColPrimaryPO_MO
        '
        Me.ColPrimaryPO_MO.HeaderText = "Primary PO MO"
        Me.ColPrimaryPO_MO.Name = "ColPrimaryPO_MO"
        Me.ColPrimaryPO_MO.ReadOnly = True
        Me.ColPrimaryPO_MO.Visible = False
        Me.ColPrimaryPO_MO.Width = 131
        '
        'ColPrimaryProduct
        '
        Me.ColPrimaryProduct.HeaderText = "Primary Product"
        Me.ColPrimaryProduct.Name = "ColPrimaryProduct"
        Me.ColPrimaryProduct.ReadOnly = True
        Me.ColPrimaryProduct.Visible = False
        Me.ColPrimaryProduct.Width = 134
        '
        'FrmTrans_MaterialPlanningDashboard_FollowUpDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(952, 418)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.DgvItem)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(958, 446)
        Me.Name = "FrmTrans_MaterialPlanningDashboard_FollowUpDetail"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Follow Up Detail"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColDocumentNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColJONo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDefaultQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitDefault As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColMultiQty As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUnitMulti As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTimestamp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryPO_MO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryProduct As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
