<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmVersionControl
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColComputerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLastUserLogin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLastLoginTime = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductKey = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVersion = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiry = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AllowUserToResizeColumns = False
        Me.dgvItem.AllowUserToResizeRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColComputerName, Me.ColLastUserLogin, Me.ColLastLoginTime, Me.ColProductKey, Me.ColVersion, Me.ColExpiry})
        Me.dgvItem.Location = New System.Drawing.Point(12, 12)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersVisible = False
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(808, 509)
        Me.dgvItem.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColComputerName
        '
        Me.ColComputerName.HeaderText = "Computer Name"
        Me.ColComputerName.Name = "ColComputerName"
        Me.ColComputerName.ReadOnly = True
        '
        'ColLastUserLogin
        '
        Me.ColLastUserLogin.HeaderText = "Last User Login"
        Me.ColLastUserLogin.Name = "ColLastUserLogin"
        Me.ColLastUserLogin.ReadOnly = True
        '
        'ColLastLoginTime
        '
        Me.ColLastLoginTime.HeaderText = "Last Login Time"
        Me.ColLastLoginTime.Name = "ColLastLoginTime"
        Me.ColLastLoginTime.ReadOnly = True
        '
        'ColProductKey
        '
        Me.ColProductKey.HeaderText = "Product Key"
        Me.ColProductKey.Name = "ColProductKey"
        Me.ColProductKey.ReadOnly = True
        '
        'ColVersion
        '
        Me.ColVersion.HeaderText = "Version"
        Me.ColVersion.Name = "ColVersion"
        Me.ColVersion.ReadOnly = True
        '
        'ColExpiry
        '
        Me.ColExpiry.HeaderText = "Expiry"
        Me.ColExpiry.Name = "ColExpiry"
        Me.ColExpiry.ReadOnly = True
        '
        'FrmVersionControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(832, 533)
        Me.Controls.Add(Me.dgvItem)
        Me.MinimumSize = New System.Drawing.Size(840, 570)
        Me.Name = "FrmVersionControl"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Version Control"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColComputerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLastUserLogin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLastLoginTime As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductKey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVersion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiry As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
