<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiImport_DebitNote
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.btnFindFile = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.dgvItem = New System.Windows.Forms.DataGridView
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboSeparator = New System.Windows.Forms.ComboBox
        Me.bgw1 = New System.ComponentModel.BackgroundWorker
        Me.LblProgress = New System.Windows.Forms.Label
        Me.prgBar1 = New System.Windows.Forms.ProgressBar
        Me.bgw2 = New System.ComponentModel.BackgroundWorker
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colTipePartner = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colJournalNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colDescription = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colContactPerson = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colAddress = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colReferedDocNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDueDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrency = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColAmount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExchangeRate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColARAccount = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColException = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.colIdentifierAddress = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 12)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Source File"
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Location = New System.Drawing.Point(73, 10)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(262, 20)
        Me.TextBox1.TabIndex = 1
        '
        'btnFindFile
        '
        Me.btnFindFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFindFile.Location = New System.Drawing.Point(338, 10)
        Me.btnFindFile.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindFile.Name = "btnFindFile"
        Me.btnFindFile.Size = New System.Drawing.Size(56, 20)
        Me.btnFindFile.TabIndex = 2
        Me.btnFindFile.Text = "Find"
        Me.btnFindFile.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStart.Location = New System.Drawing.Point(338, 33)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(2)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(56, 20)
        Me.btnStart.TabIndex = 4
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'dgvItem
        '
        Me.dgvItem.AllowUserToAddRows = False
        Me.dgvItem.AllowUserToDeleteRows = False
        Me.dgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvItem.ColumnHeadersHeight = 25
        Me.dgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.colTipePartner, Me.ColCustomerCode, Me.ColCustName, Me.ColInvoiceNo, Me.ColInvoiceDate, Me.colJournalNo, Me.colDescription, Me.colContactPerson, Me.colAddress, Me.colReferedDocNo, Me.ColDueDate, Me.ColCurrency, Me.ColAmount, Me.ColExchangeRate, Me.ColARAccount, Me.ColRemarks, Me.ColException, Me.colIdentifierAddress})
        Me.dgvItem.Location = New System.Drawing.Point(9, 58)
        Me.dgvItem.Margin = New System.Windows.Forms.Padding(2)
        Me.dgvItem.Name = "dgvItem"
        Me.dgvItem.ReadOnly = True
        Me.dgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvItem.RowTemplate.Height = 24
        Me.dgvItem.Size = New System.Drawing.Size(718, 438)
        Me.dgvItem.TabIndex = 6
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "csv files|*.csv"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 37)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Separator"
        '
        'cboSeparator
        '
        Me.cboSeparator.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSeparator.FormattingEnabled = True
        Me.cboSeparator.Items.AddRange(New Object() {",", ";"})
        Me.cboSeparator.Location = New System.Drawing.Point(73, 33)
        Me.cboSeparator.Margin = New System.Windows.Forms.Padding(2)
        Me.cboSeparator.Name = "cboSeparator"
        Me.cboSeparator.Size = New System.Drawing.Size(262, 21)
        Me.cboSeparator.TabIndex = 8
        '
        'bgw1
        '
        Me.bgw1.WorkerReportsProgress = True
        '
        'LblProgress
        '
        Me.LblProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LblProgress.AutoSize = True
        Me.LblProgress.Location = New System.Drawing.Point(399, 37)
        Me.LblProgress.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LblProgress.Name = "LblProgress"
        Me.LblProgress.Size = New System.Drawing.Size(16, 13)
        Me.LblProgress.TabIndex = 9
        Me.LblProgress.Text = "..."
        '
        'prgBar1
        '
        Me.prgBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgBar1.Location = New System.Drawing.Point(399, 10)
        Me.prgBar1.Margin = New System.Windows.Forms.Padding(2)
        Me.prgBar1.Name = "prgBar1"
        Me.prgBar1.Size = New System.Drawing.Size(328, 19)
        Me.prgBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.prgBar1.TabIndex = 10
        '
        'bgw2
        '
        Me.bgw2.WorkerReportsProgress = True
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'colTipePartner
        '
        Me.colTipePartner.HeaderText = "Type"
        Me.colTipePartner.Name = "colTipePartner"
        Me.colTipePartner.ReadOnly = True
        Me.colTipePartner.Width = 56
        '
        'ColCustomerCode
        '
        Me.ColCustomerCode.HeaderText = "Code"
        Me.ColCustomerCode.Name = "ColCustomerCode"
        Me.ColCustomerCode.ReadOnly = True
        Me.ColCustomerCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustomerCode.Width = 38
        '
        'ColCustName
        '
        Me.ColCustName.HeaderText = "Name"
        Me.ColCustName.Name = "ColCustName"
        Me.ColCustName.ReadOnly = True
        Me.ColCustName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustName.Width = 41
        '
        'ColInvoiceNo
        '
        Me.ColInvoiceNo.HeaderText = "Document No"
        Me.ColInvoiceNo.Name = "ColInvoiceNo"
        Me.ColInvoiceNo.ReadOnly = True
        Me.ColInvoiceNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColInvoiceNo.Width = 79
        '
        'ColInvoiceDate
        '
        Me.ColInvoiceDate.HeaderText = "Document Date"
        Me.ColInvoiceDate.Name = "ColInvoiceDate"
        Me.ColInvoiceDate.ReadOnly = True
        Me.ColInvoiceDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColInvoiceDate.Width = 88
        '
        'colJournalNo
        '
        Me.colJournalNo.HeaderText = "Journal No"
        Me.colJournalNo.Name = "colJournalNo"
        Me.colJournalNo.ReadOnly = True
        Me.colJournalNo.Width = 83
        '
        'colDescription
        '
        Me.colDescription.HeaderText = "Journal Description"
        Me.colDescription.Name = "colDescription"
        Me.colDescription.ReadOnly = True
        Me.colDescription.Width = 122
        '
        'colContactPerson
        '
        Me.colContactPerson.HeaderText = "Contact Person"
        Me.colContactPerson.Name = "colContactPerson"
        Me.colContactPerson.ReadOnly = True
        Me.colContactPerson.Width = 105
        '
        'colAddress
        '
        Me.colAddress.HeaderText = "Address"
        Me.colAddress.Name = "colAddress"
        Me.colAddress.ReadOnly = True
        Me.colAddress.Width = 70
        '
        'colReferedDocNo
        '
        Me.colReferedDocNo.HeaderText = "Refered Document No"
        Me.colReferedDocNo.Name = "colReferedDocNo"
        Me.colReferedDocNo.ReadOnly = True
        Me.colReferedDocNo.Width = 139
        '
        'ColDueDate
        '
        Me.ColDueDate.HeaderText = "Due Date"
        Me.ColDueDate.Name = "ColDueDate"
        Me.ColDueDate.ReadOnly = True
        Me.ColDueDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColDueDate.Width = 59
        '
        'ColCurrency
        '
        Me.ColCurrency.HeaderText = "Currency"
        Me.ColCurrency.Name = "ColCurrency"
        Me.ColCurrency.ReadOnly = True
        Me.ColCurrency.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCurrency.Width = 55
        '
        'ColAmount
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColAmount.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColAmount.HeaderText = "Amount"
        Me.ColAmount.Name = "ColAmount"
        Me.ColAmount.ReadOnly = True
        Me.ColAmount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColAmount.Width = 49
        '
        'ColExchangeRate
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColExchangeRate.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColExchangeRate.HeaderText = "Exchange Rate"
        Me.ColExchangeRate.Name = "ColExchangeRate"
        Me.ColExchangeRate.ReadOnly = True
        Me.ColExchangeRate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColExchangeRate.Width = 87
        '
        'ColARAccount
        '
        Me.ColARAccount.HeaderText = "AR Account"
        Me.ColARAccount.Name = "ColARAccount"
        Me.ColARAccount.ReadOnly = True
        Me.ColARAccount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColARAccount.Width = 71
        '
        'ColRemarks
        '
        Me.ColRemarks.HeaderText = "Reason For Claim"
        Me.ColRemarks.Name = "ColRemarks"
        Me.ColRemarks.ReadOnly = True
        Me.ColRemarks.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColRemarks.Width = 96
        '
        'ColException
        '
        Me.ColException.HeaderText = "Exception"
        Me.ColException.Name = "ColException"
        Me.ColException.ReadOnly = True
        Me.ColException.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColException.Width = 60
        '
        'colIdentifierAddress
        '
        Me.colIdentifierAddress.HeaderText = "Identifier Address"
        Me.colIdentifierAddress.Name = "colIdentifierAddress"
        Me.colIdentifierAddress.ReadOnly = True
        Me.colIdentifierAddress.Visible = False
        Me.colIdentifierAddress.Width = 113
        '
        'FrmTransaksiImport_DebitNote
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(744, 519)
        Me.Controls.Add(Me.prgBar1)
        Me.Controls.Add(Me.LblProgress)
        Me.Controls.Add(Me.cboSeparator)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.dgvItem)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnFindFile)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(752, 550)
        Me.Name = "FrmTransaksiImport_DebitNote"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Import Beginning Debit Note"
        CType(Me.dgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents btnFindFile As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents dgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cboSeparator As System.Windows.Forms.ComboBox
    Friend WithEvents bgw1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents LblProgress As System.Windows.Forms.Label
    Friend WithEvents prgBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents bgw2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTipePartner As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colJournalNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDescription As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colContactPerson As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colAddress As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colReferedDocNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDueDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrency As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColAmount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExchangeRate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColARAccount As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColException As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colIdentifierAddress As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
