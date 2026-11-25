<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SettingSalaryNPWP
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.BtnClose = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNama = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNIK = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWPLama = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryKaryawan = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColKode, Me.ColNama, Me.ColNIK, Me.ColNPWP, Me.ColNPWPLama, Me.ColReason, Me.ColPrimaryKaryawan})
        Me.DgvItem.Location = New System.Drawing.Point(12, 32)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(812, 411)
        Me.DgvItem.TabIndex = 0
        '
        'BtnOK
        '
        Me.BtnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnOK.Location = New System.Drawing.Point(658, 449)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(80, 27)
        Me.BtnOK.TabIndex = 1
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseCompatibleTextRendering = True
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(573, 20)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Based On Government Regulation No.: PMK 112/PMK.03/2022 NPWP No. Contains of 16 D" & _
            "igit."
        Me.Label3.UseCompatibleTextRendering = True
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.Location = New System.Drawing.Point(744, 449)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(80, 27)
        Me.BtnClose.TabIndex = 6
        Me.BtnClose.Text = "&Close"
        Me.BtnClose.UseCompatibleTextRendering = True
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 51
        '
        'ColKode
        '
        Me.ColKode.HeaderText = "Employee Code"
        Me.ColKode.Name = "ColKode"
        Me.ColKode.Width = 132
        '
        'ColNama
        '
        Me.ColNama.HeaderText = "Employee Name"
        Me.ColNama.Name = "ColNama"
        Me.ColNama.Width = 136
        '
        'ColNIK
        '
        Me.ColNIK.HeaderText = "NIK (Nomor Induk Kependudukan)"
        Me.ColNIK.Name = "ColNIK"
        Me.ColNIK.Width = 249
        '
        'ColNPWP
        '
        Me.ColNPWP.HeaderText = "NPWP No"
        Me.ColNPWP.MaxInputLength = 22
        Me.ColNPWP.Name = "ColNPWP"
        Me.ColNPWP.Width = 96
        '
        'ColNPWPLama
        '
        Me.ColNPWPLama.HeaderText = "NPWPLama"
        Me.ColNPWPLama.MaxInputLength = 22
        Me.ColNPWPLama.Name = "ColNPWPLama"
        Me.ColNPWPLama.ReadOnly = True
        Me.ColNPWPLama.Visible = False
        Me.ColNPWPLama.Width = 109
        '
        'ColReason
        '
        Me.ColReason.HeaderText = "Reason"
        Me.ColReason.Name = "ColReason"
        Me.ColReason.ReadOnly = True
        Me.ColReason.Width = 82
        '
        'ColPrimaryKaryawan
        '
        Me.ColPrimaryKaryawan.HeaderText = "Primary"
        Me.ColPrimaryKaryawan.Name = "ColPrimaryKaryawan"
        Me.ColPrimaryKaryawan.ReadOnly = True
        Me.ColPrimaryKaryawan.Visible = False
        Me.ColPrimaryKaryawan.Width = 81
        '
        'FrmTrans_SalaryNPWPMassage
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(848, 502)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.DgvItem)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.BtnOK)
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_SalaryNPWPMassage"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Information NPWP No. Failed"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNama As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNIK As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWPLama As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryKaryawan As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
