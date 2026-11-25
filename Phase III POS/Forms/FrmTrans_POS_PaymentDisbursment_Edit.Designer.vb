<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTrans_POS_PaymentDisbursment_Edit
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
        Me.spc = New System.Windows.Forms.SplitContainer
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboPaymentStatus = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransactionNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAccountNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColApprovalCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBankCharge = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBankPayment = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDifference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.spc.Panel1.SuspendLayout()
        Me.spc.Panel2.SuspendLayout()
        Me.spc.SuspendLayout()
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'spc
        '
        Me.spc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.spc.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.spc.Location = New System.Drawing.Point(0, 0)
        Me.spc.Name = "spc"
        Me.spc.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'spc.Panel1
        '
        Me.spc.Panel1.Controls.Add(Me.btnCancel)
        Me.spc.Panel1.Controls.Add(Me.btnSave)
        Me.spc.Panel1.Controls.Add(Me.dtpFrom)
        Me.spc.Panel1.Controls.Add(Me.Label1)
        Me.spc.Panel1.Controls.Add(Me.cboPaymentStatus)
        Me.spc.Panel1.Controls.Add(Me.Label4)
        '
        'spc.Panel2
        '
        Me.spc.Panel2.Controls.Add(Me.dgvItem)
        Me.spc.Size = New System.Drawing.Size(799, 362)
        Me.spc.SplitterDistance = 63
        Me.spc.TabIndex = 1
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(291, 38)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(291, 12)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(75, 23)
        Me.btnSave.TabIndex = 11
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dtpFrom
        '
        Me.dtpFrom.CustomFormat = "dd MMM yyyy"
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(96, 12)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(189, 20)
        Me.dtpFrom.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "From"
        '
        'cboPaymentStatus
        '
        Me.cboPaymentStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPaymentStatus.FormattingEnabled = True
        Me.cboPaymentStatus.Items.AddRange(New Object() {"Open", "Succeed", "Failed"})
        Me.cboPaymentStatus.Location = New System.Drawing.Point(96, 38)
        Me.cboPaymentStatus.Name = "cboPaymentStatus"
        Me.cboPaymentStatus.Size = New System.Drawing.Size(189, 21)
        Me.cboPaymentStatus.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(9, 41)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Payment Status"
        '
        'dgvItem
        '
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColTransactionNumber, Me.ColDate, Me.ColAccountNumber, Me.ColApprovalCode, Me.ColAmount, Me.ColBankCharge, Me.ColBankPayment, Me.ColDifference, Me.ColStatus})
        Me.dgvItem.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvItem.Location = New System.Drawing.Point(0, 0)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.Size = New System.Drawing.Size(799, 295)
        Me.dgvItem.TabIndex = 0
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.Width = 46
        '
        'ColTransactionNumber
        '
        Me.ColTransactionNumber.HeaderText = "Transaction Number"
        Me.ColTransactionNumber.Name = "ColTransactionNumber"
        Me.ColTransactionNumber.Width = 128
        '
        'ColDate
        '
        Me.ColDate.HeaderText = "Date"
        Me.ColDate.Name = "ColDate"
        Me.ColDate.Width = 55
        '
        'ColAccountNumber
        '
        Me.ColAccountNumber.HeaderText = "Account Number"
        Me.ColAccountNumber.Name = "ColAccountNumber"
        Me.ColAccountNumber.Width = 112
        '
        'ColApprovalCode
        '
        Me.ColApprovalCode.HeaderText = "Approval Code"
        Me.ColApprovalCode.Name = "ColApprovalCode"
        Me.ColApprovalCode.Width = 102
        '
        'ColAmount
        '
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.Width = 68
        '
        'ColBankCharge
        '
        Me.ColBankCharge.HeaderText = "Charge"
        Me.ColBankCharge.Name = "ColBankCharge"
        Me.ColBankCharge.Width = 66
        '
        'ColBankPayment
        '
        Me.ColBankPayment.HeaderText = "Bank Payment"
        Me.ColBankPayment.Name = "ColBankPayment"
        Me.ColBankPayment.Width = 101
        '
        'ColDifference
        '
        Me.ColDifference.HeaderText = "Difference"
        Me.ColDifference.Name = "ColDifference"
        Me.ColDifference.Width = 81
        '
        'ColStatus
        '
        Me.ColStatus.HeaderText = "Status"
        Me.ColStatus.Name = "ColStatus"
        Me.ColStatus.Width = 62
        '
        'FrmTrans_POS_PaymentDisbursment_Edit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(799, 362)
        Me.Controls.Add(Me.spc)
        Me.MinimumSize = New System.Drawing.Size(815, 400)
        Me.Name = "FrmTrans_POS_PaymentDisbursment_Edit"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Data"
        Me.spc.Panel1.ResumeLayout(False)
        Me.spc.Panel1.PerformLayout()
        Me.spc.Panel2.ResumeLayout(False)
        Me.spc.ResumeLayout(False)
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents spc As System.Windows.Forms.SplitContainer
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dtpFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboPaymentStatus As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransactionNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAccountNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColApprovalCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBankCharge As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBankPayment As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDifference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColStatus As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
