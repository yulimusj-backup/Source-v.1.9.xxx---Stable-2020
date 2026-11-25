<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_SPL_Authorize_Access
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnExit = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.LlbUser = New System.Windows.Forms.LinkLabel
        Me.dgvUser = New System.Windows.Forms.DataGridView
        Me.col2No = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col2UserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col2Username = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.col2PrimaryUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.dgvManager = New System.Windows.Forms.DataGridView
        Me.ChkActive = New System.Windows.Forms.CheckBox
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ApplyToAllUserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ApplyToUserGroupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUserCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colUsername = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colAccessType = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.colPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifier = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvManager, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnExit)
        Me.GroupBox1.Controls.Add(Me.btnCancel)
        Me.GroupBox1.Controls.Add(Me.btnOK)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.dgvManager)
        Me.GroupBox1.Controls.Add(Me.ChkActive)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(424, 417)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'btnExit
        '
        Me.btnExit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExit.Location = New System.Drawing.Point(340, 387)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(75, 23)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "&Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(263, 387)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOK.Location = New System.Drawing.Point(186, 387)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(75, 23)
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.LlbUser)
        Me.GroupBox2.Controls.Add(Me.dgvUser)
        Me.GroupBox2.Location = New System.Drawing.Point(6, 191)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(412, 193)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "User"
        '
        'LlbUser
        '
        Me.LlbUser.ActiveLinkColor = System.Drawing.Color.Blue
        Me.LlbUser.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUser.AutoSize = True
        Me.LlbUser.Enabled = False
        Me.LlbUser.Location = New System.Drawing.Point(362, 0)
        Me.LlbUser.Name = "LlbUser"
        Me.LlbUser.Size = New System.Drawing.Size(50, 13)
        Me.LlbUser.TabIndex = 2
        Me.LlbUser.TabStop = True
        Me.LlbUser.Text = "Advance"
        Me.LlbUser.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'dgvUser
        '
        Me.dgvUser.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvUser.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvUser.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUser.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col2No, Me.col2UserCode, Me.col2Username, Me.col2PrimaryUser})
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvUser.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvUser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvUser.Enabled = False
        Me.dgvUser.Location = New System.Drawing.Point(3, 16)
        Me.dgvUser.Name = "dgvUser"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvUser.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvUser.Size = New System.Drawing.Size(406, 174)
        Me.dgvUser.TabIndex = 0
        '
        'col2No
        '
        Me.col2No.HeaderText = "No"
        Me.col2No.Name = "col2No"
        Me.col2No.Width = 46
        '
        'col2UserCode
        '
        Me.col2UserCode.HeaderText = "User Code"
        Me.col2UserCode.Name = "col2UserCode"
        Me.col2UserCode.Width = 82
        '
        'col2Username
        '
        Me.col2Username.HeaderText = "User Name"
        Me.col2Username.Name = "col2Username"
        Me.col2Username.Width = 85
        '
        'col2PrimaryUser
        '
        Me.col2PrimaryUser.HeaderText = "Primary User"
        Me.col2PrimaryUser.Name = "col2PrimaryUser"
        Me.col2PrimaryUser.Visible = False
        Me.col2PrimaryUser.Width = 91
        '
        'dgvManager
        '
        Me.dgvManager.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvManager.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvManager.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvManager.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvManager.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvManager.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colUserCode, Me.colUsername, Me.colAccessType, Me.colPrimary, Me.ColIdentifier})
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvManager.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvManager.Location = New System.Drawing.Point(6, 35)
        Me.dgvManager.Name = "dgvManager"
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvManager.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvManager.Size = New System.Drawing.Size(412, 150)
        Me.dgvManager.TabIndex = 1
        '
        'ChkActive
        '
        Me.ChkActive.AutoSize = True
        Me.ChkActive.Location = New System.Drawing.Point(6, 12)
        Me.ChkActive.Name = "ChkActive"
        Me.ChkActive.Size = New System.Drawing.Size(230, 17)
        Me.ChkActive.TabIndex = 0
        Me.ChkActive.Text = "Activate Overtime Order Access Restriction"
        Me.ChkActive.UseVisualStyleBackColor = True
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ApplyToAllUserToolStripMenuItem, Me.ApplyToUserGroupToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(182, 48)
        '
        'ApplyToAllUserToolStripMenuItem
        '
        Me.ApplyToAllUserToolStripMenuItem.Name = "ApplyToAllUserToolStripMenuItem"
        Me.ApplyToAllUserToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.ApplyToAllUserToolStripMenuItem.Text = "Apply to All User"
        '
        'ApplyToUserGroupToolStripMenuItem
        '
        Me.ApplyToUserGroupToolStripMenuItem.Name = "ApplyToUserGroupToolStripMenuItem"
        Me.ApplyToUserGroupToolStripMenuItem.Size = New System.Drawing.Size(181, 22)
        Me.ApplyToUserGroupToolStripMenuItem.Text = "Apply to User Group"
        '
        'BackgroundWorker1
        '
        '
        'colNo
        '
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.Width = 46
        '
        'colUserCode
        '
        Me.colUserCode.HeaderText = "User Code"
        Me.colUserCode.Name = "colUserCode"
        Me.colUserCode.Width = 82
        '
        'colUsername
        '
        Me.colUsername.HeaderText = "User Name"
        Me.colUsername.Name = "colUsername"
        Me.colUsername.Width = 85
        '
        'colAccessType
        '
        Me.colAccessType.HeaderText = "Access Type"
        Me.colAccessType.Items.AddRange(New Object() {"Full Access", "Selected User Only"})
        Me.colAccessType.Name = "colAccessType"
        Me.colAccessType.Width = 75
        '
        'colPrimary
        '
        Me.colPrimary.HeaderText = "Primary User"
        Me.colPrimary.Name = "colPrimary"
        Me.colPrimary.Visible = False
        Me.colPrimary.Width = 91
        '
        'ColIdentifier
        '
        Me.ColIdentifier.HeaderText = "Identifier"
        Me.ColIdentifier.Name = "ColIdentifier"
        Me.ColIdentifier.Visible = False
        Me.ColIdentifier.Width = 72
        '
        'FrmTrans_SPL_Authorize_Access
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(424, 417)
        Me.Controls.Add(Me.GroupBox1)
        Me.MinimumSize = New System.Drawing.Size(440, 455)
        Me.Name = "FrmTrans_SPL_Authorize_Access"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Overtime Order Auth Access"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvManager, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkActive As System.Windows.Forms.CheckBox
    Friend WithEvents dgvManager As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents LlbUser As System.Windows.Forms.LinkLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ApplyToAllUserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents ApplyToUserGroupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents col2No As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2UserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2Username As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents col2PrimaryUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUserCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colUsername As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colAccessType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifier As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
