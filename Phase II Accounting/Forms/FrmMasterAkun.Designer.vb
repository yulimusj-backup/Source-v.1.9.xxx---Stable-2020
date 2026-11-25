<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmMasterAkun
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvAkun = New System.Windows.Forms.DataGridView
        Me.btnUp = New System.Windows.Forms.Button
        Me.btnDown = New System.Windows.Forms.Button
        Me.BtnRight = New System.Windows.Forms.Button
        Me.btnLeft = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnBudget = New System.Windows.Forms.Button
        Me.btnEdit = New System.Windows.Forms.Button
        Me.BtnPrint = New System.Windows.Forms.Button
        Me.ColAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDebet = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColKredit = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColHeader = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkun = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColUrutan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLvl1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLvl2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLvl3 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColLevel = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipePerkiraan = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSaldoNormal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeAkun2 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeAkun2Str = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLvl4 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLvl5 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLvl6 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryLvl7 = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvAkun, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvAkun
        '
        Me.dgvAkun.AllowUserToAddRows = False
        Me.dgvAkun.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.dgvAkun.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvAkun.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAkun.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvAkun.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvAkun.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dgvAkun.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAkun.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColAccount, Me.ColDebet, Me.ColKredit, Me.ColHeader, Me.ColPrimaryAkun, Me.ColUrutan, Me.ColPrimaryLvl1, Me.ColPrimaryLvl2, Me.ColPrimaryLvl3, Me.ColLevel, Me.ColTipePerkiraan, Me.ColSaldoNormal, Me.ColTipeAkun2, Me.ColTipeAkun2Str, Me.ColPrimaryLvl4, Me.ColPrimaryLvl5, Me.ColPrimaryLvl6, Me.ColPrimaryLvl7})
        Me.dgvAkun.Location = New System.Drawing.Point(12, 12)
        Me.dgvAkun.MultiSelect = False
        Me.dgvAkun.Name = "dgvAkun"
        Me.dgvAkun.ReadOnly = True
        Me.dgvAkun.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dgvAkun.RowTemplate.Height = 24
        Me.dgvAkun.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvAkun.Size = New System.Drawing.Size(667, 466)
        Me.dgvAkun.TabIndex = 0
        '
        'btnUp
        '
        Me.btnUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnUp.BackgroundImage = Global.AccModule.My.Resources.Resources.Up_Arrow
        Me.btnUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnUp.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.btnUp.Location = New System.Drawing.Point(12, 484)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(35, 23)
        Me.btnUp.TabIndex = 1
        Me.btnUp.UseVisualStyleBackColor = True
        '
        'btnDown
        '
        Me.btnDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDown.BackgroundImage = Global.AccModule.My.Resources.Resources.Down_Arrow
        Me.btnDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnDown.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.btnDown.Location = New System.Drawing.Point(53, 484)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(35, 23)
        Me.btnDown.TabIndex = 2
        Me.btnDown.UseVisualStyleBackColor = True
        '
        'BtnRight
        '
        Me.BtnRight.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnRight.BackgroundImage = Global.AccModule.My.Resources.Resources.Right_Arrow
        Me.BtnRight.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.BtnRight.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.BtnRight.Location = New System.Drawing.Point(135, 484)
        Me.BtnRight.Name = "BtnRight"
        Me.BtnRight.Size = New System.Drawing.Size(35, 23)
        Me.BtnRight.TabIndex = 4
        Me.BtnRight.UseVisualStyleBackColor = True
        '
        'btnLeft
        '
        Me.btnLeft.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnLeft.BackgroundImage = Global.AccModule.My.Resources.Resources.Left_Arrow
        Me.btnLeft.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnLeft.FlatAppearance.BorderColor = System.Drawing.SystemColors.Control
        Me.btnLeft.Location = New System.Drawing.Point(94, 484)
        Me.btnLeft.Name = "btnLeft"
        Me.btnLeft.Size = New System.Drawing.Size(35, 23)
        Me.btnLeft.TabIndex = 3
        Me.btnLeft.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(636, 484)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(43, 23)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnNew
        '
        Me.btnNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNew.Location = New System.Drawing.Point(176, 484)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(78, 23)
        Me.btnNew.TabIndex = 5
        Me.btnNew.Text = "New"
        Me.btnNew.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(344, 484)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(78, 23)
        Me.btnDelete.TabIndex = 7
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnBudget
        '
        Me.btnBudget.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnBudget.Location = New System.Drawing.Point(520, 484)
        Me.btnBudget.Name = "btnBudget"
        Me.btnBudget.Size = New System.Drawing.Size(86, 23)
        Me.btnBudget.TabIndex = 8
        Me.btnBudget.Text = "Go to Budget"
        Me.btnBudget.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnEdit.Location = New System.Drawing.Point(260, 484)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(78, 23)
        Me.btnEdit.TabIndex = 6
        Me.btnEdit.Text = "Edit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'BtnPrint
        '
        Me.BtnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnPrint.Location = New System.Drawing.Point(428, 484)
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(86, 23)
        Me.BtnPrint.TabIndex = 10
        Me.BtnPrint.Text = "Print"
        Me.BtnPrint.UseVisualStyleBackColor = True
        '
        'ColAccount
        '
        Me.ColAccount.HeaderText = "Account"
        Me.ColAccount.Name = "ColAccount"
        Me.ColAccount.ReadOnly = True
        Me.ColAccount.Width = 72
        '
        'ColDebet
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDebet.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColDebet.HeaderText = "Debet"
        Me.ColDebet.Name = "ColDebet"
        Me.ColDebet.ReadOnly = True
        Me.ColDebet.Visible = False
        Me.ColDebet.Width = 61
        '
        'ColKredit
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColKredit.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColKredit.HeaderText = "Kredit"
        Me.ColKredit.Name = "ColKredit"
        Me.ColKredit.ReadOnly = True
        Me.ColKredit.Visible = False
        Me.ColKredit.Width = 59
        '
        'ColHeader
        '
        Me.ColHeader.HeaderText = "Header"
        Me.ColHeader.Name = "ColHeader"
        Me.ColHeader.ReadOnly = True
        Me.ColHeader.Visible = False
        Me.ColHeader.Width = 67
        '
        'ColPrimaryAkun
        '
        Me.ColPrimaryAkun.HeaderText = "Primary"
        Me.ColPrimaryAkun.Name = "ColPrimaryAkun"
        Me.ColPrimaryAkun.ReadOnly = True
        Me.ColPrimaryAkun.Visible = False
        Me.ColPrimaryAkun.Width = 66
        '
        'ColUrutan
        '
        Me.ColUrutan.HeaderText = "Urutan"
        Me.ColUrutan.Name = "ColUrutan"
        Me.ColUrutan.ReadOnly = True
        Me.ColUrutan.Visible = False
        Me.ColUrutan.Width = 64
        '
        'ColPrimaryLvl1
        '
        Me.ColPrimaryLvl1.HeaderText = "PrimaryLvl1"
        Me.ColPrimaryLvl1.Name = "ColPrimaryLvl1"
        Me.ColPrimaryLvl1.ReadOnly = True
        Me.ColPrimaryLvl1.Visible = False
        Me.ColPrimaryLvl1.Width = 86
        '
        'ColPrimaryLvl2
        '
        Me.ColPrimaryLvl2.HeaderText = "PrimaryLvl2"
        Me.ColPrimaryLvl2.Name = "ColPrimaryLvl2"
        Me.ColPrimaryLvl2.ReadOnly = True
        Me.ColPrimaryLvl2.Visible = False
        Me.ColPrimaryLvl2.Width = 86
        '
        'ColPrimaryLvl3
        '
        Me.ColPrimaryLvl3.HeaderText = "PrimaryLvl3"
        Me.ColPrimaryLvl3.Name = "ColPrimaryLvl3"
        Me.ColPrimaryLvl3.ReadOnly = True
        Me.ColPrimaryLvl3.Visible = False
        Me.ColPrimaryLvl3.Width = 86
        '
        'ColLevel
        '
        Me.ColLevel.HeaderText = "Level"
        Me.ColLevel.Name = "ColLevel"
        Me.ColLevel.ReadOnly = True
        Me.ColLevel.Visible = False
        Me.ColLevel.Width = 58
        '
        'ColTipePerkiraan
        '
        Me.ColTipePerkiraan.HeaderText = "Tipe_Perkiraan"
        Me.ColTipePerkiraan.Name = "ColTipePerkiraan"
        Me.ColTipePerkiraan.ReadOnly = True
        Me.ColTipePerkiraan.Visible = False
        Me.ColTipePerkiraan.Width = 104
        '
        'ColSaldoNormal
        '
        Me.ColSaldoNormal.HeaderText = "SaldoNormal"
        Me.ColSaldoNormal.Name = "ColSaldoNormal"
        Me.ColSaldoNormal.ReadOnly = True
        Me.ColSaldoNormal.Visible = False
        Me.ColSaldoNormal.Width = 92
        '
        'ColTipeAkun2
        '
        Me.ColTipeAkun2.HeaderText = "TipeAkun"
        Me.ColTipeAkun2.Name = "ColTipeAkun2"
        Me.ColTipeAkun2.ReadOnly = True
        Me.ColTipeAkun2.Visible = False
        Me.ColTipeAkun2.Width = 78
        '
        'ColTipeAkun2Str
        '
        Me.ColTipeAkun2Str.HeaderText = "Account Type"
        Me.ColTipeAkun2Str.Name = "ColTipeAkun2Str"
        Me.ColTipeAkun2Str.ReadOnly = True
        Me.ColTipeAkun2Str.Width = 99
        '
        'ColPrimaryLvl4
        '
        Me.ColPrimaryLvl4.HeaderText = "Primary Lvl 4"
        Me.ColPrimaryLvl4.Name = "ColPrimaryLvl4"
        Me.ColPrimaryLvl4.ReadOnly = True
        Me.ColPrimaryLvl4.Visible = False
        Me.ColPrimaryLvl4.Width = 92
        '
        'ColPrimaryLvl5
        '
        Me.ColPrimaryLvl5.HeaderText = "Primary Lvl 5"
        Me.ColPrimaryLvl5.Name = "ColPrimaryLvl5"
        Me.ColPrimaryLvl5.ReadOnly = True
        Me.ColPrimaryLvl5.Visible = False
        Me.ColPrimaryLvl5.Width = 92
        '
        'ColPrimaryLvl6
        '
        Me.ColPrimaryLvl6.HeaderText = "Primary Lvl 6"
        Me.ColPrimaryLvl6.Name = "ColPrimaryLvl6"
        Me.ColPrimaryLvl6.ReadOnly = True
        Me.ColPrimaryLvl6.Visible = False
        Me.ColPrimaryLvl6.Width = 92
        '
        'ColPrimaryLvl7
        '
        Me.ColPrimaryLvl7.HeaderText = "Primary Lvl 7"
        Me.ColPrimaryLvl7.Name = "ColPrimaryLvl7"
        Me.ColPrimaryLvl7.ReadOnly = True
        Me.ColPrimaryLvl7.Visible = False
        Me.ColPrimaryLvl7.Width = 92
        '
        'FrmMasterAkun
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(699, 519)
        Me.Controls.Add(Me.BtnPrint)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnBudget)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.BtnRight)
        Me.Controls.Add(Me.btnLeft)
        Me.Controls.Add(Me.btnDown)
        Me.Controls.Add(Me.btnUp)
        Me.Controls.Add(Me.dgvAkun)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(709, 556)
        Me.Name = "FrmMasterAkun"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Chart Of Account"
        CType(Me.dgvAkun, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents BtnRight As System.Windows.Forms.Button
    Friend WithEvents btnLeft As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnBudget As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents dgvAkun As System.Windows.Forms.DataGridView
    Friend WithEvents BtnPrint As System.Windows.Forms.Button
    Friend WithEvents ColAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDebet As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColKredit As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColHeader As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkun As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColUrutan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLvl1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLvl2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLvl3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLevel As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipePerkiraan As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSaldoNormal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeAkun2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeAkun2Str As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLvl4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLvl5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLvl6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryLvl7 As System.Windows.Forms.DataGridViewTextBoxColumn

End Class
