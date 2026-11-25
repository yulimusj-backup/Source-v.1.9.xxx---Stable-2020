<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmBackUpData1
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txthost = New System.Windows.Forms.TextBox
        Me.txtdatabase = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtuserid = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtport = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtpassw = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.chkCompress = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.FrmDia = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.btnEncrypt = New System.Windows.Forms.Button
        Me.btnDecrypt = New System.Windows.Forms.Button
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.btnRestructure = New System.Windows.Forms.Button
        Me.chkExcludeMatdash = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(98, 152)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(101, 26)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Back Up"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "host"
        '
        'txthost
        '
        Me.txthost.Location = New System.Drawing.Point(98, 12)
        Me.txthost.Name = "txthost"
        Me.txthost.Size = New System.Drawing.Size(440, 22)
        Me.txthost.TabIndex = 2
        '
        'txtdatabase
        '
        Me.txtdatabase.Location = New System.Drawing.Point(98, 40)
        Me.txtdatabase.Name = "txtdatabase"
        Me.txtdatabase.Size = New System.Drawing.Size(440, 22)
        Me.txtdatabase.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "database"
        '
        'txtuserid
        '
        Me.txtuserid.Location = New System.Drawing.Point(98, 96)
        Me.txtuserid.Name = "txtuserid"
        Me.txtuserid.Size = New System.Drawing.Size(440, 22)
        Me.txtuserid.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 99)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "user id"
        '
        'txtport
        '
        Me.txtport.Location = New System.Drawing.Point(98, 68)
        Me.txtport.Name = "txtport"
        Me.txtport.Size = New System.Drawing.Size(440, 22)
        Me.txtport.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 71)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(33, 17)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "port"
        '
        'txtpassw
        '
        Me.txtpassw.Location = New System.Drawing.Point(98, 124)
        Me.txtpassw.Name = "txtpassw"
        Me.txtpassw.Size = New System.Drawing.Size(440, 22)
        Me.txtpassw.TabIndex = 10
        Me.txtpassw.UseSystemPasswordChar = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 127)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 17)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "password"
        '
        'chkCompress
        '
        Me.chkCompress.AutoSize = True
        Me.chkCompress.Location = New System.Drawing.Point(210, 154)
        Me.chkCompress.Name = "chkCompress"
        Me.chkCompress.Size = New System.Drawing.Size(93, 21)
        Me.chkCompress.TabIndex = 11
        Me.chkCompress.Text = "Compress"
        Me.chkCompress.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(309, 154)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(126, 21)
        Me.CheckBox1.TabIndex = 12
        Me.CheckBox1.Text = "Shutdown After"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(98, 184)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(101, 26)
        Me.Button2.TabIndex = 13
        Me.Button2.Text = "Restore"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'FrmDia
        '
        Me.FrmDia.DefaultExt = "sql"
        Me.FrmDia.Filter = "sql files (*.sql)|*.sql"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.AddExtension = False
        Me.OpenFileDialog1.DefaultExt = "sql"
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Filter = "sql files (*.sql)|*.sql"
        '
        'btnEncrypt
        '
        Me.btnEncrypt.Location = New System.Drawing.Point(322, 184)
        Me.btnEncrypt.Name = "btnEncrypt"
        Me.btnEncrypt.Size = New System.Drawing.Size(101, 26)
        Me.btnEncrypt.TabIndex = 14
        Me.btnEncrypt.Text = "Encrypt"
        Me.btnEncrypt.UseVisualStyleBackColor = True
        '
        'btnDecrypt
        '
        Me.btnDecrypt.Location = New System.Drawing.Point(434, 184)
        Me.btnDecrypt.Name = "btnDecrypt"
        Me.btnDecrypt.Size = New System.Drawing.Size(101, 26)
        Me.btnDecrypt.TabIndex = 15
        Me.btnDecrypt.Text = "Decrypt"
        Me.btnDecrypt.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 249)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(526, 23)
        Me.ProgressBar1.TabIndex = 16
        '
        'btnRestructure
        '
        Me.btnRestructure.Location = New System.Drawing.Point(210, 184)
        Me.btnRestructure.Name = "btnRestructure"
        Me.btnRestructure.Size = New System.Drawing.Size(101, 26)
        Me.btnRestructure.TabIndex = 17
        Me.btnRestructure.Text = "Restructure"
        Me.btnRestructure.UseVisualStyleBackColor = True
        '
        'chkExcludeMatdash
        '
        Me.chkExcludeMatdash.AutoSize = True
        Me.chkExcludeMatdash.Location = New System.Drawing.Point(98, 216)
        Me.chkExcludeMatdash.Name = "chkExcludeMatdash"
        Me.chkExcludeMatdash.Size = New System.Drawing.Size(137, 21)
        Me.chkExcludeMatdash.TabIndex = 18
        Me.chkExcludeMatdash.Text = "Exclude Matdash"
        Me.chkExcludeMatdash.UseVisualStyleBackColor = True
        '
        'FrmBackUpData1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(550, 284)
        Me.Controls.Add(Me.chkExcludeMatdash)
        Me.Controls.Add(Me.btnRestructure)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btnDecrypt)
        Me.Controls.Add(Me.btnEncrypt)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.chkCompress)
        Me.Controls.Add(Me.txtpassw)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtuserid)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtport)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtdatabase)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txthost)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "FrmBackUpData1"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SMESS Databack Up"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txthost As System.Windows.Forms.TextBox
    Friend WithEvents txtdatabase As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtuserid As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtport As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtpassw As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents chkCompress As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents FrmDia As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnEncrypt As System.Windows.Forms.Button
    Friend WithEvents btnDecrypt As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnRestructure As System.Windows.Forms.Button
    Friend WithEvents chkExcludeMatdash As System.Windows.Forms.CheckBox
End Class
