<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiSO_ReminderAR
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
        Me.btnNo = New System.Windows.Forms.Button
        Me.btnYes = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrencyOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDueDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnNo
        '
        Me.btnNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNo.Location = New System.Drawing.Point(705, 367)
        Me.btnNo.Name = "btnNo"
        Me.btnNo.Size = New System.Drawing.Size(75, 23)
        Me.btnNo.TabIndex = 0
        Me.btnNo.Text = "No"
        Me.btnNo.UseVisualStyleBackColor = True
        '
        'btnYes
        '
        Me.btnYes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnYes.Location = New System.Drawing.Point(624, 367)
        Me.btnYes.Name = "btnYes"
        Me.btnYes.Size = New System.Drawing.Size(75, 23)
        Me.btnYes.TabIndex = 1
        Me.btnYes.Text = "Yes"
        Me.btnYes.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 30
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColInvNo, Me.ColCurrencyOriginal, Me.ColBalanceOriginal, Me.ColBalance, Me.ColDueDate})
        Me.dgvItem.Location = New System.Drawing.Point(12, 35)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(768, 326)
        Me.dgvItem.TabIndex = 2
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.Width = 51
        '
        'ColInvNo
        '
        Me.ColInvNo.HeaderText = "Invoice Number"
        Me.ColInvNo.Name = "ColInvNo"
        Me.ColInvNo.ReadOnly = True
        Me.ColInvNo.Width = 131
        '
        'ColCurrencyOriginal
        '
        Me.ColCurrencyOriginal.HeaderText = ""
        Me.ColCurrencyOriginal.Name = "ColCurrencyOriginal"
        Me.ColCurrencyOriginal.ReadOnly = True
        Me.ColCurrencyOriginal.Width = 19
        '
        'ColBalanceOriginal
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalanceOriginal.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColBalanceOriginal.HeaderText = "Balance Original"
        Me.ColBalanceOriginal.Name = "ColBalanceOriginal"
        Me.ColBalanceOriginal.ReadOnly = True
        Me.ColBalanceOriginal.Width = 137
        '
        'ColBalance
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalance.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColBalance.HeaderText = "Balance"
        Me.ColBalance.Name = "ColBalance"
        Me.ColBalance.ReadOnly = True
        Me.ColBalance.Width = 84
        '
        'ColDueDate
        '
        Me.ColDueDate.HeaderText = "Due Date"
        Me.ColDueDate.Name = "ColDueDate"
        Me.ColDueDate.ReadOnly = True
        Me.ColDueDate.Width = 93
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Label1"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 369)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(106, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Continue save?"
        '
        'FrmTransaksiSO_ReminderAR
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(792, 393)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.btnYes)
        Me.Controls.Add(Me.btnNo)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(600, 400)
        Me.Name = "FrmTransaksiSO_ReminderAR"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Account Receivable Reminder"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnNo As System.Windows.Forms.Button
    Friend WithEvents btnYes As System.Windows.Forms.Button
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrencyOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDueDate As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
