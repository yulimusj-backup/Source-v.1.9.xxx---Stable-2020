<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSettingPO
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
        Me.TxtPrimary = New System.Windows.Forms.TextBox
        Me.TxtReadOnly = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.TxtAcknowledgeBy_Jabatan = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TxtAcknowledgeBy = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtApprovedBy_Jabatan = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtApprovedBy = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TxtPreparedBy_Jabatan = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TxtPreparedBy = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TxtPrimary
        '
        Me.TxtPrimary.Location = New System.Drawing.Point(73, 58)
        Me.TxtPrimary.Name = "TxtPrimary"
        Me.TxtPrimary.Size = New System.Drawing.Size(61, 20)
        Me.TxtPrimary.TabIndex = 24
        Me.TxtPrimary.TabStop = False
        Me.TxtPrimary.Visible = False
        '
        'TxtReadOnly
        '
        Me.TxtReadOnly.Location = New System.Drawing.Point(73, 84)
        Me.TxtReadOnly.Name = "TxtReadOnly"
        Me.TxtReadOnly.ReadOnly = True
        Me.TxtReadOnly.Size = New System.Drawing.Size(61, 20)
        Me.TxtReadOnly.TabIndex = 26
        Me.TxtReadOnly.TabStop = False
        Me.TxtReadOnly.Visible = False
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(276, 162)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 22
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(7, 139)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(127, 13)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Acknowledge By Position"
        '
        'TxtAcknowledgeBy_Jabatan
        '
        Me.TxtAcknowledgeBy_Jabatan.Location = New System.Drawing.Point(140, 136)
        Me.TxtAcknowledgeBy_Jabatan.Name = "TxtAcknowledgeBy_Jabatan"
        Me.TxtAcknowledgeBy_Jabatan.Size = New System.Drawing.Size(292, 20)
        Me.TxtAcknowledgeBy_Jabatan.TabIndex = 20
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(7, 113)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(87, 13)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "Acknowledge By"
        '
        'TxtAcknowledgeBy
        '
        Me.TxtAcknowledgeBy.Location = New System.Drawing.Point(140, 110)
        Me.TxtAcknowledgeBy.Name = "TxtAcknowledgeBy"
        Me.TxtAcknowledgeBy.Size = New System.Drawing.Size(292, 20)
        Me.TxtAcknowledgeBy.TabIndex = 18
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(7, 87)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(108, 13)
        Me.Label3.TabIndex = 25
        Me.Label3.Text = "Approved By Position"
        '
        'TxtApprovedBy_Jabatan
        '
        Me.TxtApprovedBy_Jabatan.Location = New System.Drawing.Point(140, 84)
        Me.TxtApprovedBy_Jabatan.Name = "TxtApprovedBy_Jabatan"
        Me.TxtApprovedBy_Jabatan.Size = New System.Drawing.Size(292, 20)
        Me.TxtApprovedBy_Jabatan.TabIndex = 17
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(7, 61)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Approved By"
        '
        'TxtApprovedBy
        '
        Me.TxtApprovedBy.Location = New System.Drawing.Point(140, 58)
        Me.TxtApprovedBy.Name = "TxtApprovedBy"
        Me.TxtApprovedBy.Size = New System.Drawing.Size(292, 20)
        Me.TxtApprovedBy.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(105, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Prepared By Position"
        '
        'TxtPreparedBy_Jabatan
        '
        Me.TxtPreparedBy_Jabatan.Location = New System.Drawing.Point(140, 32)
        Me.TxtPreparedBy_Jabatan.Name = "TxtPreparedBy_Jabatan"
        Me.TxtPreparedBy_Jabatan.Size = New System.Drawing.Size(292, 20)
        Me.TxtPreparedBy_Jabatan.TabIndex = 14
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Prepared By"
        '
        'TxtPreparedBy
        '
        Me.TxtPreparedBy.Location = New System.Drawing.Point(140, 6)
        Me.TxtPreparedBy.Name = "TxtPreparedBy"
        Me.TxtPreparedBy.Size = New System.Drawing.Size(292, 20)
        Me.TxtPreparedBy.TabIndex = 13
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(357, 162)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 23)
        Me.btnClose.TabIndex = 23
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmTransaksiSettingPO
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(438, 191)
        Me.Controls.Add(Me.TxtPrimary)
        Me.Controls.Add(Me.TxtReadOnly)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TxtAcknowledgeBy_Jabatan)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TxtAcknowledgeBy)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TxtApprovedBy_Jabatan)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TxtApprovedBy)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtPreparedBy_Jabatan)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtPreparedBy)
        Me.Controls.Add(Me.btnClose)
        Me.MinimumSize = New System.Drawing.Size(454, 227)
        Me.Name = "FrmTransaksiSettingPO"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "PO Settings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtPrimary As System.Windows.Forms.TextBox
    Friend WithEvents TxtReadOnly As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtAcknowledgeBy_Jabatan As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TxtAcknowledgeBy As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtApprovedBy_Jabatan As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TxtApprovedBy As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtPreparedBy_Jabatan As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPreparedBy As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
