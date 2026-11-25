<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPLApprovalBudget
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
        Me.ColDept = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPeriod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewLinkColumn
        Me.ColDeptPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BtnClose = New System.Windows.Forms.Button
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColDept, Me.ColPeriod, Me.ColStatus, Me.ColDeptPrimary})
        Me.DgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DgvItem.Location = New System.Drawing.Point(0, 0)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.ReadOnly = True
        Me.DgvItem.RowHeadersVisible = False
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.Size = New System.Drawing.Size(449, 293)
        Me.DgvItem.TabIndex = 0
        '
        'ColDept
        '
        Me.ColDept.HeaderText = "Department"
        Me.ColDept.Name = "ColDept"
        Me.ColDept.ReadOnly = True
        Me.ColDept.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColPeriod
        '
        Me.ColPeriod.HeaderText = "Period"
        Me.ColPeriod.Name = "ColPeriod"
        Me.ColPeriod.ReadOnly = True
        Me.ColPeriod.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.ReadOnly = True
        Me.ColStatus.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColStatus.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'ColDeptPrimary
        '
        Me.ColDeptPrimary.HeaderText = "Primary Departmen"
        Me.ColDeptPrimary.Name = "ColDeptPrimary"
        Me.ColDeptPrimary.ReadOnly = True
        Me.ColDeptPrimary.Visible = False
        '
        'BtnClose
        '
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(374, 0)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 1
        Me.BtnClose.TabStop = False
        Me.BtnClose.Text = "Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'FrmTrans_SPLApprovalBudget
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(449, 293)
        Me.Controls.Add(Me.DgvItem)
        Me.Controls.Add(Me.BtnClose)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(465, 331)
        Me.Name = "FrmTrans_SPLApprovalBudget"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Order Approval Status"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColDept As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPeriod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewLinkColumn
    Friend WithEvents ColDeptPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BtnClose As System.Windows.Forms.Button
End Class
