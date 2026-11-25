<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransksiCreateEmail
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
        Me.clbUser = New System.Windows.Forms.CheckedListBox
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.llbSend = New System.Windows.Forms.LinkLabel
        Me.llbExit = New System.Windows.Forms.LinkLabel
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LinkLabel4 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel3 = New System.Windows.Forms.LinkLabel
        Me.txtSubject = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'clbUser
        '
        Me.clbUser.CheckOnClick = True
        Me.clbUser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbUser.FormattingEnabled = True
        Me.clbUser.Location = New System.Drawing.Point(3, 16)
        Me.clbUser.MultiColumn = True
        Me.clbUser.Name = "clbUser"
        Me.clbUser.Size = New System.Drawing.Size(606, 64)
        Me.clbUser.TabIndex = 0
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(12, 163)
        Me.txtEmail.MaxLength = 250
        Me.txtEmail.Multiline = True
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(612, 250)
        Me.txtEmail.TabIndex = 2
        '
        'llbSend
        '
        Me.llbSend.AutoSize = True
        Me.llbSend.Location = New System.Drawing.Point(483, 427)
        Me.llbSend.Name = "llbSend"
        Me.llbSend.Size = New System.Drawing.Size(60, 13)
        Me.llbSend.TabIndex = 3
        Me.llbSend.TabStop = True
        Me.llbSend.Text = "Send Email"
        '
        'llbExit
        '
        Me.llbExit.AutoSize = True
        Me.llbExit.Location = New System.Drawing.Point(584, 427)
        Me.llbExit.Name = "llbExit"
        Me.llbExit.Size = New System.Drawing.Size(40, 13)
        Me.llbExit.TabIndex = 4
        Me.llbExit.TabStop = True
        Me.llbExit.Text = "Cancel"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.E_mail_Module.My.Resources.Resources.edit
        Me.PictureBox1.Location = New System.Drawing.Point(448, 419)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.E_mail_Module.My.Resources.Resources.x
        Me.PictureBox2.Location = New System.Drawing.Point(549, 419)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox2.TabIndex = 5
        Me.PictureBox2.TabStop = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LinkLabel4)
        Me.GroupBox1.Controls.Add(Me.LinkLabel3)
        Me.GroupBox1.Controls.Add(Me.clbUser)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 63)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(612, 94)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "To"
        '
        'LinkLabel4
        '
        Me.LinkLabel4.AutoSize = True
        Me.LinkLabel4.Location = New System.Drawing.Point(541, 0)
        Me.LinkLabel4.Name = "LinkLabel4"
        Me.LinkLabel4.Size = New System.Drawing.Size(65, 13)
        Me.LinkLabel4.TabIndex = 2
        Me.LinkLabel4.TabStop = True
        Me.LinkLabel4.Text = "Uncheck All"
        '
        'LinkLabel3
        '
        Me.LinkLabel3.AutoSize = True
        Me.LinkLabel3.Location = New System.Drawing.Point(483, 0)
        Me.LinkLabel3.Name = "LinkLabel3"
        Me.LinkLabel3.Size = New System.Drawing.Size(52, 13)
        Me.LinkLabel3.TabIndex = 1
        Me.LinkLabel3.TabStop = True
        Me.LinkLabel3.Text = "Check All"
        '
        'txtSubject
        '
        Me.txtSubject.Location = New System.Drawing.Point(64, 12)
        Me.txtSubject.MaxLength = 150
        Me.txtSubject.Multiline = True
        Me.txtSubject.Name = "txtSubject"
        Me.txtSubject.Size = New System.Drawing.Size(560, 45)
        Me.txtSubject.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(46, 13)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Subject:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(297, 215)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(98, 26)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmTransksiCreateEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(636, 449)
        Me.Controls.Add(Me.txtSubject)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.llbExit)
        Me.Controls.Add(Me.llbSend)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.Button1)
        Me.Name = "FrmTransksiCreateEmail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create New Email"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents clbUser As System.Windows.Forms.CheckedListBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents llbSend As System.Windows.Forms.LinkLabel
    Friend WithEvents llbExit As System.Windows.Forms.LinkLabel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel4 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel3 As System.Windows.Forms.LinkLabel
    Friend WithEvents txtSubject As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
