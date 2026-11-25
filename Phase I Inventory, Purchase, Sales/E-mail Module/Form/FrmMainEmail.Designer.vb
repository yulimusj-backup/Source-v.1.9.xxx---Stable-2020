<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMainEmail
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
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMainEmail))
        Me.gboFolders = New System.Windows.Forms.GroupBox
        Me.PictureBox8 = New System.Windows.Forms.PictureBox
        Me.llbClearAlreadyRead = New System.Windows.Forms.LinkLabel
        Me.llbClearSentItem = New System.Windows.Forms.LinkLabel
        Me.PictureBox7 = New System.Windows.Forms.PictureBox
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.llbclearInbox = New System.Windows.Forms.LinkLabel
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.llbCreate = New System.Windows.Forms.LinkLabel
        Me.llbSentItem = New System.Windows.Forms.LinkLabel
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.llbInbox = New System.Windows.Forms.LinkLabel
        Me.llbRespond = New System.Windows.Forms.LinkLabel
        Me.llbForward = New System.Windows.Forms.LinkLabel
        Me.llbDelete = New System.Windows.Forms.LinkLabel
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.gboItem = New System.Windows.Forms.GroupBox
        Me.dgvEmail = New System.Windows.Forms.DataGridView
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Column1 = New System.Windows.Forms.DataGridViewImageColumn
        Me.llbReply = New System.Windows.Forms.LinkLabel
        Me.llbRespond2 = New System.Windows.Forms.LinkLabel
        Me.DataGridViewImageColumn1 = New System.Windows.Forms.DataGridViewImageColumn
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.Button1 = New System.Windows.Forms.Button
        Me.gboFolders.SuspendLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gboItem.SuspendLayout()
        CType(Me.dgvEmail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gboFolders
        '
        Me.gboFolders.Controls.Add(Me.PictureBox8)
        Me.gboFolders.Controls.Add(Me.llbClearAlreadyRead)
        Me.gboFolders.Controls.Add(Me.llbClearSentItem)
        Me.gboFolders.Controls.Add(Me.PictureBox7)
        Me.gboFolders.Controls.Add(Me.PictureBox6)
        Me.gboFolders.Controls.Add(Me.llbclearInbox)
        Me.gboFolders.Controls.Add(Me.PictureBox3)
        Me.gboFolders.Controls.Add(Me.PictureBox2)
        Me.gboFolders.Controls.Add(Me.llbCreate)
        Me.gboFolders.Controls.Add(Me.llbSentItem)
        Me.gboFolders.Controls.Add(Me.PictureBox1)
        Me.gboFolders.Controls.Add(Me.llbInbox)
        Me.gboFolders.Location = New System.Drawing.Point(4, 12)
        Me.gboFolders.Name = "gboFolders"
        Me.gboFolders.Size = New System.Drawing.Size(153, 226)
        Me.gboFolders.TabIndex = 0
        Me.gboFolders.TabStop = False
        '
        'PictureBox8
        '
        Me.PictureBox8.Image = Global.E_mail_Module.My.Resources.Resources.delete
        Me.PictureBox8.Location = New System.Drawing.Point(6, 154)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox8.TabIndex = 28
        Me.PictureBox8.TabStop = False
        '
        'llbClearAlreadyRead
        '
        Me.llbClearAlreadyRead.AutoSize = True
        Me.llbClearAlreadyRead.Location = New System.Drawing.Point(41, 162)
        Me.llbClearAlreadyRead.Name = "llbClearAlreadyRead"
        Me.llbClearAlreadyRead.Size = New System.Drawing.Size(98, 13)
        Me.llbClearAlreadyRead.TabIndex = 27
        Me.llbClearAlreadyRead.TabStop = True
        Me.llbClearAlreadyRead.Text = "Clear Already Read"
        '
        'llbClearSentItem
        '
        Me.llbClearSentItem.AutoSize = True
        Me.llbClearSentItem.Location = New System.Drawing.Point(41, 198)
        Me.llbClearSentItem.Name = "llbClearSentItem"
        Me.llbClearSentItem.Size = New System.Drawing.Size(84, 13)
        Me.llbClearSentItem.TabIndex = 4
        Me.llbClearSentItem.TabStop = True
        Me.llbClearSentItem.Text = "Clear Sent Items"
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = Global.E_mail_Module.My.Resources.Resources.clear_sent_item
        Me.PictureBox7.Location = New System.Drawing.Point(6, 190)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox7.TabIndex = 26
        Me.PictureBox7.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = Global.E_mail_Module.My.Resources.Resources.clear_inbox
        Me.PictureBox6.Location = New System.Drawing.Point(6, 121)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox6.TabIndex = 25
        Me.PictureBox6.TabStop = False
        '
        'llbclearInbox
        '
        Me.llbclearInbox.AutoSize = True
        Me.llbclearInbox.Location = New System.Drawing.Point(41, 129)
        Me.llbclearInbox.Name = "llbclearInbox"
        Me.llbclearInbox.Size = New System.Drawing.Size(60, 13)
        Me.llbclearInbox.TabIndex = 3
        Me.llbclearInbox.TabStop = True
        Me.llbclearInbox.Text = "Clear Inbox"
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = Global.E_mail_Module.My.Resources.Resources.add_file
        Me.PictureBox3.Location = New System.Drawing.Point(6, 87)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox3.TabIndex = 4
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.E_mail_Module.My.Resources.Resources.archive_32
        Me.PictureBox2.Location = New System.Drawing.Point(6, 53)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox2.TabIndex = 3
        Me.PictureBox2.TabStop = False
        '
        'llbCreate
        '
        Me.llbCreate.AutoSize = True
        Me.llbCreate.Location = New System.Drawing.Point(41, 95)
        Me.llbCreate.Name = "llbCreate"
        Me.llbCreate.Size = New System.Drawing.Size(91, 13)
        Me.llbCreate.TabIndex = 2
        Me.llbCreate.TabStop = True
        Me.llbCreate.Text = "Create New Email"
        '
        'llbSentItem
        '
        Me.llbSentItem.AutoSize = True
        Me.llbSentItem.Location = New System.Drawing.Point(41, 61)
        Me.llbSentItem.Name = "llbSentItem"
        Me.llbSentItem.Size = New System.Drawing.Size(57, 13)
        Me.llbSentItem.TabIndex = 1
        Me.llbSentItem.TabStop = True
        Me.llbSentItem.Text = "Sent Items"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.E_mail_Module.My.Resources.Resources.open
        Me.PictureBox1.Location = New System.Drawing.Point(6, 19)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'llbInbox
        '
        Me.llbInbox.AutoSize = True
        Me.llbInbox.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.llbInbox.Location = New System.Drawing.Point(41, 27)
        Me.llbInbox.Name = "llbInbox"
        Me.llbInbox.Size = New System.Drawing.Size(33, 13)
        Me.llbInbox.TabIndex = 0
        Me.llbInbox.TabStop = True
        Me.llbInbox.Text = "Inbox"
        Me.llbInbox.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'llbRespond
        '
        Me.llbRespond.AutoSize = True
        Me.llbRespond.Enabled = False
        Me.llbRespond.Location = New System.Drawing.Point(366, 218)
        Me.llbRespond.Name = "llbRespond"
        Me.llbRespond.Size = New System.Drawing.Size(78, 13)
        Me.llbRespond.TabIndex = 5
        Me.llbRespond.TabStop = True
        Me.llbRespond.Text = "Respon Button"
        Me.llbRespond.Visible = False
        '
        'llbForward
        '
        Me.llbForward.AutoSize = True
        Me.llbForward.Enabled = False
        Me.llbForward.Location = New System.Drawing.Point(315, 218)
        Me.llbForward.Name = "llbForward"
        Me.llbForward.Size = New System.Drawing.Size(45, 13)
        Me.llbForward.TabIndex = 4
        Me.llbForward.TabStop = True
        Me.llbForward.Text = "Forward"
        '
        'llbDelete
        '
        Me.llbDelete.AutoSize = True
        Me.llbDelete.Enabled = False
        Me.llbDelete.Location = New System.Drawing.Point(196, 218)
        Me.llbDelete.Name = "llbDelete"
        Me.llbDelete.Size = New System.Drawing.Size(38, 13)
        Me.llbDelete.TabIndex = 2
        Me.llbDelete.TabStop = True
        Me.llbDelete.Text = "Delete"
        '
        'txtEmail
        '
        Me.txtEmail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEmail.BackColor = System.Drawing.Color.White
        Me.txtEmail.Location = New System.Drawing.Point(4, 244)
        Me.txtEmail.Multiline = True
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.ReadOnly = True
        Me.txtEmail.Size = New System.Drawing.Size(806, 237)
        Me.txtEmail.TabIndex = 7
        '
        'gboItem
        '
        Me.gboItem.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gboItem.Controls.Add(Me.dgvEmail)
        Me.gboItem.Controls.Add(Me.Button1)
        Me.gboItem.Location = New System.Drawing.Point(163, 12)
        Me.gboItem.Name = "gboItem"
        Me.gboItem.Size = New System.Drawing.Size(647, 192)
        Me.gboItem.TabIndex = 1
        Me.gboItem.TabStop = False
        '
        'dgvEmail
        '
        Me.dgvEmail.AllowUserToAddRows = False
        Me.dgvEmail.AllowUserToResizeRows = False
        Me.dgvEmail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvEmail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column2, Me.Column1})
        Me.dgvEmail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvEmail.Location = New System.Drawing.Point(3, 16)
        Me.dgvEmail.MultiSelect = False
        Me.dgvEmail.Name = "dgvEmail"
        Me.dgvEmail.ReadOnly = True
        Me.dgvEmail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvEmail.Size = New System.Drawing.Size(641, 173)
        Me.dgvEmail.TabIndex = 0
        '
        'Column2
        '
        Me.Column2.HeaderText = "Column2"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 73
        '
        'Column1
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.NullValue = CType(resources.GetObject("DataGridViewCellStyle2.NullValue"), Object)
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column1.DefaultCellStyle = DataGridViewCellStyle2
        Me.Column1.HeaderText = ""
        Me.Column1.Image = Global.E_mail_Module.My.Resources.Resources.open
        Me.Column1.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 5
        '
        'llbReply
        '
        Me.llbReply.AutoSize = True
        Me.llbReply.Enabled = False
        Me.llbReply.Location = New System.Drawing.Point(275, 218)
        Me.llbReply.Name = "llbReply"
        Me.llbReply.Size = New System.Drawing.Size(34, 13)
        Me.llbReply.TabIndex = 3
        Me.llbReply.TabStop = True
        Me.llbReply.Text = "Reply"
        '
        'llbRespond2
        '
        Me.llbRespond2.AutoSize = True
        Me.llbRespond2.Enabled = False
        Me.llbRespond2.Location = New System.Drawing.Point(450, 218)
        Me.llbRespond2.Name = "llbRespond2"
        Me.llbRespond2.Size = New System.Drawing.Size(84, 13)
        Me.llbRespond2.TabIndex = 6
        Me.llbRespond2.TabStop = True
        Me.llbRespond2.Text = "Respon Button2"
        Me.llbRespond2.Visible = False
        '
        'DataGridViewImageColumn1
        '
        Me.DataGridViewImageColumn1.HeaderText = "Image"
        Me.DataGridViewImageColumn1.Image = Global.E_mail_Module.My.Resources.Resources.open
        Me.DataGridViewImageColumn1.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom
        Me.DataGridViewImageColumn1.Name = "DataGridViewImageColumn1"
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = Global.E_mail_Module.My.Resources.Resources.delete
        Me.PictureBox4.Location = New System.Drawing.Point(166, 210)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox4.TabIndex = 5
        Me.PictureBox4.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = Global.E_mail_Module.My.Resources.Resources.Button_Reload
        Me.PictureBox5.Location = New System.Drawing.Point(240, 210)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(29, 28)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox5.TabIndex = 23
        Me.PictureBox5.TabStop = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 471)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(814, 22)
        Me.StatusStrip1.TabIndex = 24
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(383, 82)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'FrmMainEmail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button1
        Me.ClientSize = New System.Drawing.Size(814, 493)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.PictureBox5)
        Me.Controls.Add(Me.PictureBox4)
        Me.Controls.Add(Me.llbRespond2)
        Me.Controls.Add(Me.llbRespond)
        Me.Controls.Add(Me.llbReply)
        Me.Controls.Add(Me.gboItem)
        Me.Controls.Add(Me.llbForward)
        Me.Controls.Add(Me.llbDelete)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.gboFolders)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimumSize = New System.Drawing.Size(820, 519)
        Me.Name = "FrmMainEmail"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Internal Mail"
        Me.gboFolders.ResumeLayout(False)
        Me.gboFolders.PerformLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gboItem.ResumeLayout(False)
        CType(Me.dgvEmail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gboFolders As System.Windows.Forms.GroupBox
    Friend WithEvents llbRespond As System.Windows.Forms.LinkLabel
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents llbCreate As System.Windows.Forms.LinkLabel
    Friend WithEvents llbSentItem As System.Windows.Forms.LinkLabel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents llbInbox As System.Windows.Forms.LinkLabel
    Friend WithEvents llbForward As System.Windows.Forms.LinkLabel
    Friend WithEvents llbDelete As System.Windows.Forms.LinkLabel
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents gboItem As System.Windows.Forms.GroupBox
    Friend WithEvents dgvEmail As System.Windows.Forms.DataGridView
    Friend WithEvents llbReply As System.Windows.Forms.LinkLabel
    Friend WithEvents llbRespond2 As System.Windows.Forms.LinkLabel
    Friend WithEvents DataGridViewImageColumn1 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewImageColumn
    Friend WithEvents llbClearSentItem As System.Windows.Forms.LinkLabel
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents llbclearInbox As System.Windows.Forms.LinkLabel
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
    Friend WithEvents llbClearAlreadyRead As System.Windows.Forms.LinkLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
