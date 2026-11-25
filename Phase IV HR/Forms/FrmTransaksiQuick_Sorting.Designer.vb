<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiQuick_Sorting
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.BtnOK = New System.Windows.Forms.Button
        Me.ColDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSortType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColPriority = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryItem = New System.Windows.Forms.DataGridViewTextBoxColumn
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
        Me.DgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDesc, Me.ColSortType, Me.ColPriority, Me.ColPrimaryItem})
        Me.DgvItem.Location = New System.Drawing.Point(12, 12)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.Size = New System.Drawing.Size(475, 241)
        Me.DgvItem.TabIndex = 0
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnOK.Location = New System.Drawing.Point(412, 259)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(75, 23)
        Me.BtnOK.TabIndex = 1
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'ColDesc
        '
        Me.ColDesc.HeaderText = "Description"
        Me.ColDesc.Name = "ColDesc"
        Me.ColDesc.ReadOnly = True
        Me.ColDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDesc.Width = 66
        '
        'ColSortType
        '
        Me.ColSortType.HeaderText = "Sorting Type"
        Me.ColSortType.Items.AddRange(New Object() {"Ascending", "Descending", "Unsorted"})
        Me.ColSortType.Name = "ColSortType"
        Me.ColSortType.Width = 73
        '
        'ColPriority
        '
        Me.ColPriority.HeaderText = "Priority"
        Me.ColPriority.Name = "ColPriority"
        Me.ColPriority.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColPriority.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPriority.Width = 44
        '
        'ColPrimaryItem
        '
        Me.ColPrimaryItem.HeaderText = "Primary Item"
        Me.ColPrimaryItem.Name = "ColPrimaryItem"
        Me.ColPrimaryItem.ReadOnly = True
        Me.ColPrimaryItem.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPrimaryItem.Visible = False
        Me.ColPrimaryItem.Width = 70
        '
        'FrmTransaksiQuick_Sorting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnOK
        Me.ClientSize = New System.Drawing.Size(499, 292)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.DgvItem)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(515, 330)
        Me.Name = "FrmTransaksiQuick_Sorting"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quick Salary Sorting"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents ColDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSortType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColPriority As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryItem As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
