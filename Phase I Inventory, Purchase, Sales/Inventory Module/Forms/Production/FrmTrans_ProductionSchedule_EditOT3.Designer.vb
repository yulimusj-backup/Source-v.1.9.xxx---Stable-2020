<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_ProductionSchedule_EditOT3
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
        Me.dgvCek = New System.Windows.Forms.DataGridView
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWUCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColWUDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvCek, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvCek
        '
        Me.dgvCek.AllowUserToAddRows = False
        Me.dgvCek.AllowUserToDeleteRows = False
        Me.dgvCek.AllowUserToResizeColumns = False
        Me.dgvCek.AllowUserToResizeRows = False
        Me.dgvCek.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCek.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCek.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvCek.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColWUCode, Me.ColWUDesc, Me.ColDate, Me.ColDate1})
        Me.dgvCek.Location = New System.Drawing.Point(12, 29)
        Me.dgvCek.Name = "dgvCek"
        Me.dgvCek.ReadOnly = True
        Me.dgvCek.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvCek.RowTemplate.Height = 24
        Me.dgvCek.Size = New System.Drawing.Size(568, 193)
        Me.dgvCek.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(505, 228)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(326, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Please allocate workshift for the following work unit"
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 32
        '
        'ColWUCode
        '
        Me.ColWUCode.HeaderText = "Code"
        Me.ColWUCode.Name = "ColWUCode"
        Me.ColWUCode.ReadOnly = True
        Me.ColWUCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWUCode.Width = 47
        '
        'ColWUDesc
        '
        Me.ColWUDesc.HeaderText = "Description"
        Me.ColWUDesc.Name = "ColWUDesc"
        Me.ColWUDesc.ReadOnly = True
        Me.ColWUDesc.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColWUDesc.Width = 85
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Start Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.ReadOnly = True
        Me.ColDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate.Width = 78
        '
        'ColDate1
        '
        Me.ColDate1.HeaderText = "End Date"
        Me.ColDate1.Name = "ColDate1"
        Me.ColDate1.ReadOnly = True
        Me.ColDate1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDate1.Width = 73
        '
        'FrmTrans_ProductionSchedule_EditOT3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(592, 263)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.dgvCek)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(600, 300)
        Me.Name = "FrmTrans_ProductionSchedule_EditOT3"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Workshift Allocation Needed"
        CType(Me.dgvCek, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvCek As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWUCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColWUDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate1 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
