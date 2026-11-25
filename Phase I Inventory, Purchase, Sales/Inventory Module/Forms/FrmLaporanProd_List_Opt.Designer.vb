<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanProd_List_Opt
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
        Me.TxtDesc = New System.Windows.Forms.TextBox
        Me.BtnOK = New System.Windows.Forms.Button
        Me.lblDesc = New System.Windows.Forms.Label
        Me.TextKosong = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TxtDesc
        '
        Me.TxtDesc.Location = New System.Drawing.Point(75, 12)
        Me.TxtDesc.MaxLength = 250
        Me.TxtDesc.Name = "TxtDesc"
        Me.TxtDesc.Size = New System.Drawing.Size(540, 20)
        Me.TxtDesc.TabIndex = 0
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(75, 38)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(53, 23)
        Me.BtnOK.TabIndex = 1
        Me.BtnOK.Text = "OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'lblDesc
        '
        Me.lblDesc.AutoSize = True
        Me.lblDesc.Location = New System.Drawing.Point(9, 15)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(60, 13)
        Me.lblDesc.TabIndex = 2
        Me.lblDesc.Text = "Description"
        '
        'TextKosong
        '
        Me.TextKosong.Location = New System.Drawing.Point(31, 38)
        Me.TextKosong.Name = "TextKosong"
        Me.TextKosong.Size = New System.Drawing.Size(38, 20)
        Me.TextKosong.TabIndex = 3
        Me.TextKosong.Visible = False
        '
        'btnClose
        '
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(75, 39)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(53, 20)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'FrmLaporanProd_List_Opt
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(627, 62)
        Me.Controls.Add(Me.TextKosong)
        Me.Controls.Add(Me.lblDesc)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.TxtDesc)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmLaporanProd_List_Opt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Option"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtDesc As System.Windows.Forms.TextBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Public WithEvents lblDesc As System.Windows.Forms.Label
    Friend WithEvents TextKosong As System.Windows.Forms.TextBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
End Class
