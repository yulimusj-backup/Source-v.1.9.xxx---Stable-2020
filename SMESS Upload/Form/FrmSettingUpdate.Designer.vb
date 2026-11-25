<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmSettingUpdate
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
        Me.TxtDef = New System.Windows.Forms.TextBox
        Me.TxtPublic = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.BtnOk = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.TxtProductName = New System.Windows.Forms.TextBox
        Me.BtnFind = New System.Windows.Forms.Button
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.SuspendLayout()
        '
        'TxtDef
        '
        Me.TxtDef.Location = New System.Drawing.Point(93, 13)
        Me.TxtDef.Name = "TxtDef"
        Me.TxtDef.ReadOnly = True
        Me.TxtDef.Size = New System.Drawing.Size(386, 20)
        Me.TxtDef.TabIndex = 0
        '
        'TxtPublic
        '
        Me.TxtPublic.Location = New System.Drawing.Point(93, 41)
        Me.TxtPublic.Name = "TxtPublic"
        Me.TxtPublic.ReadOnly = True
        Me.TxtPublic.Size = New System.Drawing.Size(318, 20)
        Me.TxtPublic.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Default Path"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(61, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Public Path"
        '
        'BtnOk
        '
        Me.BtnOk.Location = New System.Drawing.Point(404, 66)
        Me.BtnOk.Name = "BtnOk"
        Me.BtnOk.Size = New System.Drawing.Size(75, 23)
        Me.BtnOk.TabIndex = 4
        Me.BtnOk.Text = "Ok"
        Me.BtnOk.UseVisualStyleBackColor = True
        Me.BtnOk.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Product Name"
        Me.Label3.Visible = False
        '
        'TxtProductName
        '
        Me.TxtProductName.Location = New System.Drawing.Point(93, 69)
        Me.TxtProductName.Name = "TxtProductName"
        Me.TxtProductName.ReadOnly = True
        Me.TxtProductName.Size = New System.Drawing.Size(386, 20)
        Me.TxtProductName.TabIndex = 5
        Me.TxtProductName.Visible = False
        '
        'BtnFind
        '
        Me.BtnFind.Location = New System.Drawing.Point(417, 39)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(62, 23)
        Me.BtnFind.TabIndex = 7
        Me.BtnFind.Text = "Find"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Server Path"
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'FrmSettingUpdate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(483, 65)
        Me.Controls.Add(Me.BtnFind)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.BtnOk)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TxtPublic)
        Me.Controls.Add(Me.TxtDef)
        Me.Controls.Add(Me.TxtProductName)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(499, 103)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(499, 103)
        Me.Name = "FrmSettingUpdate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TxtDef As System.Windows.Forms.TextBox
    Friend WithEvents TxtPublic As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnOk As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TxtProductName As System.Windows.Forms.TextBox
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
End Class
