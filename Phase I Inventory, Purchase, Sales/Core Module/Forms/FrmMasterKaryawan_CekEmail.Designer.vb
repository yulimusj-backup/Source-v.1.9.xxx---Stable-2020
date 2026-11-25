<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterKaryawan_CekEmail
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
        Me.dgvEmailFix = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPosition = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDepartment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmailAddress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryKaryawan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.dgvEmailFix, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvEmailFix
        '
        Me.dgvEmailFix.AllowUserToAddRows = False
        Me.dgvEmailFix.AllowUserToDeleteRows = False
        Me.dgvEmailFix.AllowUserToOrderColumns = True
        Me.dgvEmailFix.AllowUserToResizeColumns = False
        Me.dgvEmailFix.AllowUserToResizeRows = False
        Me.dgvEmailFix.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvEmailFix.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmailFix.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvEmailFix.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColPosition, Me.ColDepartment, Me.ColEmailAddress, Me.ColPrimaryKaryawan})
        Me.dgvEmailFix.Location = New System.Drawing.Point(12, 29)
        Me.dgvEmailFix.MultiSelect = False
        Me.dgvEmailFix.Name = "dgvEmailFix"
        Me.dgvEmailFix.RowHeadersVisible = False
        Me.dgvEmailFix.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvEmailFix.RowTemplate.Height = 24
        Me.dgvEmailFix.Size = New System.Drawing.Size(646, 450)
        Me.dgvEmailFix.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 51
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "EmployeeCode"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.Width = 128
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.Width = 70
        '
        'ColPosition
        '
        Me.ColPosition.HeaderText = "Position"
        Me.ColPosition.Name = "ColPosition"
        Me.ColPosition.Width = 83
        '
        'ColDepartment
        '
        Me.ColDepartment.HeaderText = "Department"
        Me.ColDepartment.Name = "ColDepartment"
        Me.ColDepartment.Width = 107
        '
        'ColEmailAddress
        '
        Me.ColEmailAddress.HeaderText = "Email Address"
        Me.ColEmailAddress.Name = "ColEmailAddress"
        Me.ColEmailAddress.Width = 123
        '
        'ColPrimaryKaryawan
        '
        Me.ColPrimaryKaryawan.HeaderText = "Primary"
        Me.ColPrimaryKaryawan.Name = "ColPrimaryKaryawan"
        Me.ColPrimaryKaryawan.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(328, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Please fix email address for the following employee"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(583, 485)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmMasterKaryawan_CekEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(670, 520)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvEmailFix)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(678, 557)
        Me.Name = "FrmMasterKaryawan_CekEmail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Check employee email"
        CType(Me.dgvEmailFix, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgvEmailFix As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPosition As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDepartment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmailAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryKaryawan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
