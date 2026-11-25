<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLicence_Add
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
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.BtnLoad = New System.Windows.Forms.Button
        Me.LlbUncheckAll = New System.Windows.Forms.LinkLabel
        Me.BtnExport = New System.Windows.Forms.Button
        Me.LlbCheckAll = New System.Windows.Forms.LinkLabel
        Me.BtnClose = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.LblCount = New System.Windows.Forms.Label
        Me.ColPrimary = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIsBlackList = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColAddLicence = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColComputerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColProductKey = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLicenceKey = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExpiryDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReplaceFrom = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReplaceFromCompName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLastUserLogin = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLastLoginTimestamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.GroupBox1.SuspendLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.DgvItem)
        Me.GroupBox1.Controls.Add(Me.BtnLoad)
        Me.GroupBox1.Controls.Add(Me.LlbUncheckAll)
        Me.GroupBox1.Controls.Add(Me.BtnExport)
        Me.GroupBox1.Controls.Add(Me.LlbCheckAll)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(965, 516)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'DgvItem
        '
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColPrimary, Me.ColIsBlackList, Me.ColAddLicence, Me.ColComputerName, Me.ColProductKey, Me.ColLicenceKey, Me.ColExpiryDate, Me.ColReplaceFrom, Me.ColReplaceFromCompName, Me.ColLastUserLogin, Me.ColLastLoginTimestamp})
        Me.DgvItem.Location = New System.Drawing.Point(9, 44)
        Me.DgvItem.MultiSelect = False
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowTemplate.Height = 24
        Me.DgvItem.Size = New System.Drawing.Size(950, 466)
        Me.DgvItem.TabIndex = 0
        '
        'BtnLoad
        '
        Me.BtnLoad.Location = New System.Drawing.Point(9, 11)
        Me.BtnLoad.Name = "BtnLoad"
        Me.BtnLoad.Size = New System.Drawing.Size(91, 24)
        Me.BtnLoad.TabIndex = 22
        Me.BtnLoad.Text = "&Load "
        Me.BtnLoad.UseVisualStyleBackColor = True
        '
        'LlbUncheckAll
        '
        Me.LlbUncheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbUncheckAll.AutoSize = True
        Me.LlbUncheckAll.Location = New System.Drawing.Point(884, 24)
        Me.LlbUncheckAll.Name = "LlbUncheckAll"
        Me.LlbUncheckAll.Size = New System.Drawing.Size(75, 17)
        Me.LlbUncheckAll.TabIndex = 2
        Me.LlbUncheckAll.TabStop = True
        Me.LlbUncheckAll.Text = "Unchek All"
        '
        'BtnExport
        '
        Me.BtnExport.Location = New System.Drawing.Point(106, 11)
        Me.BtnExport.Name = "BtnExport"
        Me.BtnExport.Size = New System.Drawing.Size(91, 24)
        Me.BtnExport.TabIndex = 5
        Me.BtnExport.Text = "&Export"
        Me.BtnExport.UseVisualStyleBackColor = True
        '
        'LlbCheckAll
        '
        Me.LlbCheckAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LlbCheckAll.AutoSize = True
        Me.LlbCheckAll.Location = New System.Drawing.Point(812, 24)
        Me.LlbCheckAll.Name = "LlbCheckAll"
        Me.LlbCheckAll.Size = New System.Drawing.Size(66, 17)
        Me.LlbCheckAll.TabIndex = 1
        Me.LlbCheckAll.TabStop = True
        Me.LlbCheckAll.Text = "Check All"
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(886, 525)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(91, 24)
        Me.BtnClose.TabIndex = 6
        Me.BtnClose.Text = "&Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "csv"
        Me.SaveFileDialog1.Filter = "csv files|*.csv"
        '
        'LblCount
        '
        Me.LblCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblCount.AutoSize = True
        Me.LblCount.Location = New System.Drawing.Point(12, 525)
        Me.LblCount.Name = "LblCount"
        Me.LblCount.Size = New System.Drawing.Size(139, 17)
        Me.LblCount.TabIndex = 23
        Me.LblCount.Text = "Number of license : -"
        '
        'ColPrimary
        '
        Me.ColPrimary.HeaderText = "Primary"
        Me.ColPrimary.Name = "ColPrimary"
        Me.ColPrimary.Visible = False
        '
        'ColIsBlackList
        '
        Me.ColIsBlackList.HeaderText = "Is BlackList"
        Me.ColIsBlackList.Name = "ColIsBlackList"
        '
        'ColAddLicence
        '
        Me.ColAddLicence.HeaderText = "Add Licence"
        Me.ColAddLicence.Name = "ColAddLicence"
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
        'ColLastUserLogin
        '
        Me.ColLastUserLogin.FillWeight = 150.0!
        Me.ColLastUserLogin.HeaderText = "Last Login User"
        Me.ColLastUserLogin.Name = "ColLastUserLogin"
        Me.ColLastUserLogin.Width = 150
        '
        'ColLastLoginTimestamp
        '
        Me.ColLastLoginTimestamp.FillWeight = 170.0!
        Me.ColLastLoginTimestamp.HeaderText = "Last Login Timestamp"
        Me.ColLastLoginTimestamp.Name = "ColLastLoginTimestamp"
        Me.ColLastLoginTimestamp.Width = 170
        '
        'FrmLicence_Add
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(989, 556)
        Me.Controls.Add(Me.LblCount)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnClose)
        Me.Name = "FrmLicence_Add"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Licence"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LlbUncheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents LlbCheckAll As System.Windows.Forms.LinkLabel
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents BtnExport As System.Windows.Forms.Button
    Friend WithEvents BtnLoad As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents LblCount As System.Windows.Forms.Label
    Friend WithEvents ColPrimary As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIsBlackList As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColAddLicence As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColComputerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColProductKey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLicenceKey As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExpiryDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReplaceFrom As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReplaceFromCompName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLastUserLogin As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLastLoginTimestamp As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
