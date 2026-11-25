<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SettingNPWP
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.CboTipeNPWP = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TxtNPWPNo = New System.Windows.Forms.TextBox
        Me.TxtNIK = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BtnOK = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Txt = New System.Windows.Forms.TextBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.TxtNPWPLama = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.CboTipeNPWP)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TxtNPWPNo)
        Me.GroupBox1.Controls.Add(Me.TxtNIK)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(-5, 38)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(426, 92)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 67)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(23, 17)
        Me.Label2.TabIndex = 73
        Me.Label2.Text = "NIK"
        Me.Label2.UseCompatibleTextRendering = True
        '
        'CboTipeNPWP
        '
        Me.CboTipeNPWP.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.CboTipeNPWP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboTipeNPWP.FormattingEnabled = True
        Me.CboTipeNPWP.Items.AddRange(New Object() {"Company", "Individual"})
        Me.CboTipeNPWP.Location = New System.Drawing.Point(106, 16)
        Me.CboTipeNPWP.Name = "CboTipeNPWP"
        Me.CboTipeNPWP.Size = New System.Drawing.Size(217, 21)
        Me.CboTipeNPWP.TabIndex = 72
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 19)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(66, 17)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "NPWP Type"
        Me.Label4.UseCompatibleTextRendering = True
        '
        'TxtNPWPNo
        '
        Me.TxtNPWPNo.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.TxtNPWPNo.Location = New System.Drawing.Point(106, 41)
        Me.TxtNPWPNo.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtNPWPNo.MaxLength = 21
        Me.TxtNPWPNo.Name = "TxtNPWPNo"
        Me.TxtNPWPNo.Size = New System.Drawing.Size(217, 20)
        Me.TxtNPWPNo.TabIndex = 4
        Me.TxtNPWPNo.Text = "000.000.000.0-000.000"
        '
        'TxtNIK
        '
        Me.TxtNIK.Location = New System.Drawing.Point(106, 64)
        Me.TxtNIK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtNIK.Name = "TxtNIK"
        Me.TxtNIK.Size = New System.Drawing.Size(217, 20)
        Me.TxtNIK.TabIndex = 6
        Me.TxtNIK.Text = "-"
        Me.TxtNIK.Visible = False
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 44)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "NPWP No."
        Me.Label1.UseCompatibleTextRendering = True
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(188, 135)
        Me.BtnOK.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(62, 25)
        Me.BtnOK.TabIndex = 2
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseCompatibleTextRendering = True
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 7)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(331, 30)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Based On Government Regulation No.: PMK 112/PMK.03/2022 " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "NPWP No. Contains of 16" & _
            " Digit."
        Me.Label3.UseCompatibleTextRendering = True
        '
        'Txt
        '
        Me.Txt.Location = New System.Drawing.Point(4, 140)
        Me.Txt.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Txt.Name = "Txt"
        Me.Txt.Size = New System.Drawing.Size(44, 20)
        Me.Txt.TabIndex = 6
        Me.Txt.Visible = False
        '
        'BtnCancel
        '
        Me.BtnCancel.Location = New System.Drawing.Point(254, 135)
        Me.BtnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(62, 25)
        Me.BtnCancel.TabIndex = 7
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseCompatibleTextRendering = True
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'TxtNPWPLama
        '
        Me.TxtNPWPLama.Location = New System.Drawing.Point(44, 142)
        Me.TxtNPWPLama.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.TxtNPWPLama.MaxLength = 21
        Me.TxtNPWPLama.Name = "TxtNPWPLama"
        Me.TxtNPWPLama.Size = New System.Drawing.Size(109, 20)
        Me.TxtNPWPLama.TabIndex = 8
        Me.TxtNPWPLama.Text = "000.000.000.0-000.000"
        Me.TxtNPWPLama.Visible = False
        '
        'FrmTrans_SettingNPWP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(344, 203)
        Me.ControlBox = False
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.TxtNPWPLama)
        Me.Controls.Add(Me.Txt)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmTrans_SettingNPWP"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Setting NPWP"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents TxtNPWPNo As System.Windows.Forms.TextBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtNIK As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents CboTipeNPWP As System.Windows.Forms.ComboBox
    Friend WithEvents Txt As System.Windows.Forms.TextBox
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Public WithEvents TxtNPWPLama As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
