<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiMonitor_Production
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTransaksiMonitor_Production))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCoputerID = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUserType = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIPAddress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTake = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColSend = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Button2 = New System.Windows.Forms.Button
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Button1 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 4)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(720, 289)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Online Users"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.ColumnHeadersHeight = 25
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColUser, Me.ColCoputerID, Me.ColUserType, Me.ColIPAddress, Me.ColStatus, Me.ColTake, Me.ColSend, Me.ColIdentifier, Me.ColPrimary})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(4, 19)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(712, 266)
        Me.DataGridView1.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColUser
        '
        Me.ColUser.HeaderText = "User"
        Me.ColUser.Name = "ColUser"
        Me.ColUser.ReadOnly = True
        Me.ColUser.Width = 63
        '
        'ColCoputerID
        '
        Me.ColCoputerID.HeaderText = "Computer ID"
        Me.ColCoputerID.Name = "ColCoputerID"
        Me.ColCoputerID.ReadOnly = True
        Me.ColCoputerID.Width = 111
        '
        'ColUserType
        '
        Me.ColUserType.HeaderText = "User Type"
        Me.ColUserType.Name = "ColUserType"
        Me.ColUserType.Visible = False
        Me.ColUserType.Width = 99
        '
        'ColIPAddress
        '
        Me.ColIPAddress.HeaderText = "IP Address"
        Me.ColIPAddress.Name = "ColIPAddress"
        Me.ColIPAddress.ReadOnly = True
        Me.ColIPAddress.Width = 101
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.Width = 73
        '
        'ColTake
        '
        Me.ColTake.HeaderText = "Take Control"
        Me.ColTake.Name = "ColTake"
        Me.ColTake.Width = 95
        '
        'ColSend
        '
        Me.ColSend.HeaderText = "Send Control"
        Me.ColSend.Name = "ColSend"
        Me.ColSend.Width = 96
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.ReadOnly = True
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 87
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.ReadOnly = True
        Me.ColPrimary.Visible = False
        Me.ColPrimary.Width = 81
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.Location = New System.Drawing.Point(192, 137)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(11, 300)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(4)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(616, 25)
        Me.ProgressBar1.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(635, 300)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(83, 26)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Refresh"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        '
        'FrmTransaksiMonitor_Production
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(739, 332)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(747, 369)
        Me.Name = "FrmTransaksiMonitor_Production"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production Planning Monitor"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCoputerID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUserType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIPAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTake As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColSend As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
