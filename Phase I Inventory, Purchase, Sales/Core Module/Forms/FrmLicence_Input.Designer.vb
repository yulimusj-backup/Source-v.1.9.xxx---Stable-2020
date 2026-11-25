<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLicence_Input
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.BtnImport = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.BtnExecute = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsBlackList = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAddLicence = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColComputerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductKey = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLicenceKey = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReplaceFrom = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReplaceFromCompName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewLicenceKey = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNewExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLastLoginUser = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLastLoginTimestamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.dgvItem)
        Me.GroupBox1.Controls.Add(Me.txtFilePath)
        Me.GroupBox1.Controls.Add(Me.BtnImport)
        Me.GroupBox1.Controls.Add(Me.btnClose)
        Me.GroupBox1.Controls.Add(Me.BtnExecute)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(878, 533)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 17)
        Me.Label2.TabIndex = 25
        Me.Label2.Text = "File Path"
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPrimary, Me.ColIsBlackList, Me.ColAddLicence, Me.ColComputerName, Me.ColProductKey, Me.ColLicenceKey, Me.ColExpiryDate, Me.ColReplaceFrom, Me.ColReplaceFromCompName, Me.ColNewLicenceKey, Me.ColNewExpiryDate, Me.ColLastLoginUser, Me.ColLastLoginTimestamp})
        Me.dgvItem.Location = New System.Drawing.Point(6, 45)
        Me.dgvItem.MultiSelect = False
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(858, 482)
        Me.dgvItem.TabIndex = 0
        '
        'txtFilePath
        '
        Me.txtFilePath.Location = New System.Drawing.Point(76, 15)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.ReadOnly = True
        Me.txtFilePath.Size = New System.Drawing.Size(497, 22)
        Me.txtFilePath.TabIndex = 23
        '
        'BtnImport
        '
        Me.BtnImport.Location = New System.Drawing.Point(579, 15)
        Me.BtnImport.Name = "BtnImport"
        Me.BtnImport.Size = New System.Drawing.Size(91, 24)
        Me.BtnImport.TabIndex = 22
        Me.BtnImport.Text = "&Import"
        Me.BtnImport.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(773, 15)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(91, 24)
        Me.btnClose.TabIndex = 6
        Me.btnClose.Text = "&Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'BtnExecute
        '
        Me.BtnExecute.Location = New System.Drawing.Point(676, 15)
        Me.BtnExecute.Name = "BtnExecute"
        Me.BtnExecute.Size = New System.Drawing.Size(91, 24)
        Me.BtnExecute.TabIndex = 5
        Me.BtnExecute.Text = "&Execute"
        Me.BtnExecute.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Visible = False
        '
        'ColIsBlackList
        '
        Me.ColIsBlackList.HeaderText = "Is Black List"
        Me.ColIsBlackList.Name = "ColIsBlackList"
        Me.ColIsBlackList.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColIsBlackList.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'ColAddLicence
        '
        Me.ColAddLicence.HeaderText = "Add Licence"
        Me.ColAddLicence.Name = "ColAddLicence"
        Me.ColAddLicence.Visible = False
        '
        'ColComputerName
        '
        Me.ColComputerName.HeaderText = "Computer Name"
        Me.ColComputerName.Name = "ColComputerName"
        '
        'ColProductKey
        '
        Me.ColProductKey.HeaderText = "Product Key"
        Me.ColProductKey.Name = "ColProductKey"
        '
        'ColLicenceKey
        '
        Me.ColLicenceKey.HeaderText = "Licence Key"
        Me.ColLicenceKey.Name = "ColLicenceKey"
        '
        'ColExpiryDate
        '
        Me.ColExpiryDate.HeaderText = "Expiry Date"
        Me.ColExpiryDate.Name = "ColExpiryDate"
        '
        'ColReplaceFrom
        '
        Me.ColReplaceFrom.HeaderText = "Replace From"
        Me.ColReplaceFrom.Name = "ColReplaceFrom"
        '
        'ColReplaceFromCompName
        '
        Me.ColReplaceFromCompName.HeaderText = "Computer Name"
        Me.ColReplaceFromCompName.Name = "ColReplaceFromCompName"
        '
        'ColNewLicenceKey
        '
        Me.ColNewLicenceKey.HeaderText = "New Licence Key"
        Me.ColNewLicenceKey.Name = "ColNewLicenceKey"
        '
        'ColNewExpiryDate
        '
        Me.ColNewExpiryDate.HeaderText = "New Expiry Date"
        Me.ColNewExpiryDate.Name = "ColNewExpiryDate"
        '
        'ColLastLoginUser
        '
        Me.ColLastLoginUser.FillWeight = 150.0!
        Me.ColLastLoginUser.HeaderText = "Last Login User"
        Me.ColLastLoginUser.Name = "ColLastLoginUser"
        Me.ColLastLoginUser.Width = 150
        '
        'ColLastLoginTimestamp
        '
        Me.ColLastLoginTimestamp.FillWeight = 170.0!
        Me.ColLastLoginTimestamp.HeaderText = "Last Login Timestamp"
        Me.ColLastLoginTimestamp.Name = "ColLastLoginTimestamp"
        Me.ColLastLoginTimestamp.Width = 170
        '
        'FrmLicence_Input
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(897, 542)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FrmLicence_Input"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Input Licence"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents BtnExecute As System.Windows.Forms.Button
    Friend WithEvents BtnImport As System.Windows.Forms.Button
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsBlackList As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAddLicence As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColComputerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductKey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLicenceKey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReplaceFrom As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReplaceFromCompName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewLicenceKey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNewExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLastLoginUser As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLastLoginTimestamp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
