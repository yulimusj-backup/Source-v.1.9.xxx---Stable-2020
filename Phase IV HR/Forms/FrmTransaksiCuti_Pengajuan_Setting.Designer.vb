<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiCuti_Pengajuan_Setting
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
        Me.ColUserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrivilage = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColEmployeeCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColEmployeeName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryEmployee = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColUserCode, Me.ColUserName, Me.ColPrivilage, Me.ColEmployeeCode, Me.ColEmployeeName, Me.ColPrimaryEmployee, Me.ColRemarks, Me.ColPrimaryUser})
        Me.DgvItem.Location = New System.Drawing.Point(12, 12)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.Size = New System.Drawing.Size(567, 425)
        Me.DgvItem.TabIndex = 0
        '
        'ColUserCode
        '
        Me.ColUserCode.HeaderText = "User Code"
        Me.ColUserCode.Name = "ColUserCode"
        Me.ColUserCode.Width = 82
        '
        'ColUserName
        '
        Me.ColUserName.HeaderText = "User Name"
        Me.ColUserName.Name = "ColUserName"
        Me.ColUserName.Width = 85
        '
        'ColPrivilage
        '
        Me.ColPrivilage.HeaderText = "Privilage"
        Me.ColPrivilage.Items.AddRange(New Object() {"Apply For All Employee", "Apply For Concerned Employee Only"})
        Me.ColPrivilage.Name = "ColPrivilage"
        Me.ColPrivilage.Width = 53
        '
        'ColEmployeeCode
        '
        Me.ColEmployeeCode.HeaderText = "Employee Code"
        Me.ColEmployeeCode.Name = "ColEmployeeCode"
        Me.ColEmployeeCode.Width = 106
        '
        'ColEmployeeName
        '
        Me.ColEmployeeName.HeaderText = "Employee Name"
        Me.ColEmployeeName.Name = "ColEmployeeName"
        Me.ColEmployeeName.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColEmployeeName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColEmployeeName.Width = 90
        '
        'ColPrimaryEmployee
        '
        Me.ColPrimaryEmployee.HeaderText = "Primary Employee"
        Me.ColPrimaryEmployee.Name = "ColPrimaryEmployee"
        Me.ColPrimaryEmployee.Visible = False
        Me.ColPrimaryEmployee.Width = 115
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.Width = 74
        '
        'ColPrimaryUser
        '
        Me.ColPrimaryUser.HeaderText = "Primary User"
        Me.ColPrimaryUser.Name = "ColPrimaryUser"
        Me.ColPrimaryUser.Visible = False
        Me.ColPrimaryUser.Width = 91
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(504, 443)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(423, 443)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        '
        'FrmTransaksiCuti_Pengajuan_Setting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(591, 478)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.DgvItem)
        Me.MinimizeBox = False
        Me.Name = "FrmTransaksiCuti_Pengajuan_Setting"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Leave Aplication Setting"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ColUserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrivilage As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColEmployeeCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColEmployeeName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryEmployee As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
