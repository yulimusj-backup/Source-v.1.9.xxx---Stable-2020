<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_Gaji_OtherDeduction
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
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.btnExit = New System.Windows.Forms.Button
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSlipDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColComponentName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeKomponen = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColTaxElement = New System.Windows.Forms.DataGridViewComboBoxColumn
        Me.ColExpenseAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOverride = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.ColOverrideReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOriginalValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExplain = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySlipDesc = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColIdentifierKomponen = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkunExpense = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExplaination = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 30
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColSlipDesc, Me.ColComponentName, Me.ColTipeKomponen, Me.ColTaxElement, Me.ColExpenseAccount, Me.ColAmount, Me.ColOverride, Me.ColOverrideReason, Me.ColOriginalValue, Me.ColExplain, Me.ColRemarks, Me.ColPrimarySlipDesc, Me.ColIdentifierKomponen, Me.ColPrimaryAkunExpense, Me.ColExplaination})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(732, 358)
        Me.dgvItem.TabIndex = 2
        '
        'btnExit
        '
        Me.btnExit.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(326, 166)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 26)
        Me.btnExit.TabIndex = 7
        Me.btnExit.Text = "E&xit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 46
        '
        'ColSlipDesc
        '
        Me.ColSlipDesc.HeaderText = "Slip Description"
        Me.ColSlipDesc.Name = "ColSlipDesc"
        Me.ColSlipDesc.ReadOnly = True
        Me.ColSlipDesc.Width = 105
        '
        'ColComponentName
        '
        Me.ColComponentName.HeaderText = "Component Name"
        Me.ColComponentName.Name = "ColComponentName"
        Me.ColComponentName.ReadOnly = True
        Me.ColComponentName.Width = 117
        '
        'ColTipeKomponen
        '
        Me.ColTipeKomponen.HeaderText = "Type"
        Me.ColTipeKomponen.Items.AddRange(New Object() {"Other Deduction"})
        Me.ColTipeKomponen.Name = "ColTipeKomponen"
        Me.ColTipeKomponen.ReadOnly = True
        Me.ColTipeKomponen.Width = 37
        '
        'ColTaxElement
        '
        Me.ColTaxElement.HeaderText = "Tax Element"
        Me.ColTaxElement.Items.AddRange(New Object() {"GAJI/PENSIUN ATAU THT/JHT", "TUNJANGAN LAINNYA, UANG LEMBUR DAN SEBAGAINYA", "HONORARIUM DAN IMBALAN LAIN SEJENISNYA", "PREMI ASURANSI YANG DIBAYARKAN PEMBERI KERJA", "PENERIMAAN DALAM BENTUK NATURA DAN KENIKMATAN LAINNYA YANG DIKENAKAN PEMOTONGAN P" & _
                        "Ph PASAL 21", "TANTIEM, BONUS, GRATIFIKASI, JASA PRODUKSI DAN THR", "IURAN PENSIUN ATAU IURAN THT/JHT"})
        Me.ColTaxElement.Name = "ColTaxElement"
        Me.ColTaxElement.ReadOnly = True
        Me.ColTaxElement.Visible = False
        Me.ColTaxElement.Width = 72
        '
        'ColExpenseAccount
        '
        Me.ColExpenseAccount.HeaderText = "Expense Account"
        Me.ColExpenseAccount.Name = "ColExpenseAccount"
        Me.ColExpenseAccount.ReadOnly = True
        Me.ColExpenseAccount.Visible = False
        Me.ColExpenseAccount.Width = 116
        '
        'ColAmount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.Width = 68
        '
        'ColOverride
        '
        Me.ColOverride.HeaderText = "Override"
        Me.ColOverride.Name = "ColOverride"
        Me.ColOverride.ReadOnly = True
        Me.ColOverride.Width = 53
        '
        'ColOverrideReason
        '
        Me.ColOverrideReason.HeaderText = "Override Reason"
        Me.ColOverrideReason.Name = "ColOverrideReason"
        Me.ColOverrideReason.ReadOnly = True
        Me.ColOverrideReason.Width = 112
        '
        'ColOriginalValue
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOriginalValue.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColOriginalValue.HeaderText = "Original Value"
        Me.ColOriginalValue.Name = "ColOriginalValue"
        Me.ColOriginalValue.ReadOnly = True
        Me.ColOriginalValue.Visible = False
        Me.ColOriginalValue.Width = 97
        '
        'ColExplain
        '
        Me.ColExplain.HeaderText = "Explain"
        Me.ColExplain.Name = "ColExplain"
        Me.ColExplain.ReadOnly = True
        Me.ColExplain.Width = 47
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Remarks"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.Width = 74
        '
        'ColPrimarySlipDesc
        '
        Me.ColPrimarySlipDesc.HeaderText = "Primary Slip Description"
        Me.ColPrimarySlipDesc.Name = "ColPrimarySlipDesc"
        Me.ColPrimarySlipDesc.ReadOnly = True
        Me.ColPrimarySlipDesc.Visible = False
        Me.ColPrimarySlipDesc.Width = 142
        '
        'ColIdentifierKomponen
        '
        Me.ColIdentifierKomponen.HeaderText = "Identifier Salary Component"
        Me.ColIdentifierKomponen.Name = "ColIdentifierKomponen"
        Me.ColIdentifierKomponen.ReadOnly = True
        Me.ColIdentifierKomponen.Visible = False
        Me.ColIdentifierKomponen.Width = 161
        '
        'ColPrimaryAkunExpense
        '
        Me.ColPrimaryAkunExpense.HeaderText = "Primary Akun Expense"
        Me.ColPrimaryAkunExpense.Name = "ColPrimaryAkunExpense"
        Me.ColPrimaryAkunExpense.ReadOnly = True
        Me.ColPrimaryAkunExpense.Visible = False
        Me.ColPrimaryAkunExpense.Width = 138
        '
        'ColExplaination
        '
        Me.ColExplaination.HeaderText = "Explaination"
        Me.ColExplaination.Name = "ColExplaination"
        Me.ColExplaination.ReadOnly = True
        Me.ColExplaination.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ColExplaination.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExplaination.Visible = False
        Me.ColExplaination.Width = 70
        '
        'FrmTrans_Gaji_OtherDeduction
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnExit
        Me.ClientSize = New System.Drawing.Size(732, 358)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.btnExit)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(748, 396)
        Me.Name = "FrmTrans_Gaji_OtherDeduction"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Other Deduction"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSlipDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColComponentName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeKomponen As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColTaxElement As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents ColExpenseAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOverride As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents ColOverrideReason As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOriginalValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExplain As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySlipDesc As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColIdentifierKomponen As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkunExpense As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExplaination As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
