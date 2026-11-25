<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmLaporanVAT2025_Import
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
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.DgvItem = New System.Windows.Forms.DataGridView
        Me.ColNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColNPWP = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCustomerName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTransactionCode = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxInvoiceNumber = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxPeriod = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxYear = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxStatus = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTaxBase = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColOtherTaxBase = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVAT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPPnBM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSignedBy = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReference = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReportedBySeller = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReportedByVATCollector = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvalidReason = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.lblProgress = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.BtnClose = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.TxtFilePath = New System.Windows.Forms.TextBox
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DgvItem
        '
        Me.DgvItem.AllowUserToAddRows = False
        Me.DgvItem.AllowUserToDeleteRows = False
        Me.DgvItem.AllowUserToResizeColumns = False
        Me.DgvItem.AllowUserToResizeRows = False
        Me.DgvItem.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DgvItem.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DgvItem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.DgvItem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColNo, Me.ColNPWP, Me.ColCustomerName, Me.ColTransactionCode, Me.ColTaxInvoiceNumber, Me.ColTaxDate, Me.ColTaxPeriod, Me.ColTaxYear, Me.ColTaxStatus, Me.ColTaxBase, Me.ColOtherTaxBase, Me.ColVAT, Me.ColPPnBM, Me.ColSignedBy, Me.ColReference, Me.ColReportedBySeller, Me.ColReportedByVATCollector, Me.ColInvalidReason})
        Me.DgvItem.Location = New System.Drawing.Point(4, 36)
        Me.DgvItem.Name = "DgvItem"
        Me.DgvItem.RowHeadersWidth = 24
        Me.DgvItem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DgvItem.Size = New System.Drawing.Size(589, 361)
        Me.DgvItem.TabIndex = 1
        '
        'ColNo
        '
        Me.ColNo.HeaderText = "No"
        Me.ColNo.Name = "ColNo"
        Me.ColNo.ReadOnly = True
        Me.ColNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNo.Width = 27
        '
        'ColNPWP
        '
        Me.ColNPWP.HeaderText = "Customer NPWP / Other Identification"
        Me.ColNPWP.Name = "ColNPWP"
        Me.ColNPWP.ReadOnly = True
        Me.ColNPWP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColNPWP.Width = 193
        '
        'ColCustomerName
        '
        Me.ColCustomerName.HeaderText = "Customer Name"
        Me.ColCustomerName.Name = "ColCustomerName"
        Me.ColCustomerName.ReadOnly = True
        Me.ColCustomerName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColCustomerName.Width = 88
        '
        'ColTransactionCode
        '
        Me.ColTransactionCode.HeaderText = "Transaction Code"
        Me.ColTransactionCode.Name = "ColTransactionCode"
        Me.ColTransactionCode.ReadOnly = True
        Me.ColTransactionCode.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTransactionCode.Width = 97
        '
        'ColTaxInvoiceNumber
        '
        Me.ColTaxInvoiceNumber.HeaderText = "Tax Invoice Number"
        Me.ColTaxInvoiceNumber.Name = "ColTaxInvoiceNumber"
        Me.ColTaxInvoiceNumber.ReadOnly = True
        Me.ColTaxInvoiceNumber.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxInvoiceNumber.Width = 109
        '
        'ColTaxDate
        '
        Me.ColTaxDate.HeaderText = "Tax Date"
        Me.ColTaxDate.Name = "ColTaxDate"
        Me.ColTaxDate.ReadOnly = True
        Me.ColTaxDate.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxDate.Width = 57
        '
        'ColTaxPeriod
        '
        Me.ColTaxPeriod.HeaderText = "Tax Period"
        Me.ColTaxPeriod.Name = "ColTaxPeriod"
        Me.ColTaxPeriod.ReadOnly = True
        Me.ColTaxPeriod.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxPeriod.Width = 64
        '
        'ColTaxYear
        '
        Me.ColTaxYear.HeaderText = "Tax Year"
        Me.ColTaxYear.Name = "ColTaxYear"
        Me.ColTaxYear.ReadOnly = True
        Me.ColTaxYear.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxYear.Width = 56
        '
        'ColTaxStatus
        '
        Me.ColTaxStatus.HeaderText = "Tax Status"
        Me.ColTaxStatus.Name = "ColTaxStatus"
        Me.ColTaxStatus.ReadOnly = True
        Me.ColTaxStatus.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxStatus.Width = 64
        '
        'ColTaxBase
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTaxBase.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColTaxBase.HeaderText = "Tax Base"
        Me.ColTaxBase.Name = "ColTaxBase"
        Me.ColTaxBase.ReadOnly = True
        Me.ColTaxBase.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColTaxBase.Width = 58
        '
        'ColOtherTaxBase
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColOtherTaxBase.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColOtherTaxBase.HeaderText = "Other Tax Base"
        Me.ColOtherTaxBase.Name = "ColOtherTaxBase"
        Me.ColOtherTaxBase.ReadOnly = True
        Me.ColOtherTaxBase.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColOtherTaxBase.Width = 87
        '
        'ColVAT
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColVAT.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColVAT.HeaderText = "VAT"
        Me.ColVAT.Name = "ColVAT"
        Me.ColVAT.ReadOnly = True
        Me.ColVAT.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColVAT.Width = 34
        '
        'ColPPnBM
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColPPnBM.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColPPnBM.HeaderText = "Luxury Goods Sales Tax (LST)"
        Me.ColPPnBM.Name = "ColPPnBM"
        Me.ColPPnBM.ReadOnly = True
        Me.ColPPnBM.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColPPnBM.Width = 157
        '
        'ColSignedBy
        '
        Me.ColSignedBy.HeaderText = "Signed By"
        Me.ColSignedBy.Name = "ColSignedBy"
        Me.ColSignedBy.ReadOnly = True
        Me.ColSignedBy.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColSignedBy.Width = 61
        '
        'ColReference
        '
        Me.ColReference.HeaderText = "Reference"
        Me.ColReference.Name = "ColReference"
        Me.ColReference.ReadOnly = True
        Me.ColReference.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReference.Width = 63
        '
        'ColReportedBySeller
        '
        Me.ColReportedBySeller.HeaderText = "Reported by Seller"
        Me.ColReportedBySeller.Name = "ColReportedBySeller"
        Me.ColReportedBySeller.ReadOnly = True
        Me.ColReportedBySeller.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ColReportedByVATCollector
        '
        Me.ColReportedByVATCollector.HeaderText = "Reported by the VAT Collector"
        Me.ColReportedByVATCollector.Name = "ColReportedByVATCollector"
        Me.ColReportedByVATCollector.ReadOnly = True
        Me.ColReportedByVATCollector.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColReportedByVATCollector.Width = 157
        '
        'ColInvalidReason
        '
        Me.ColInvalidReason.HeaderText = "Invalid Reason"
        Me.ColInvalidReason.Name = "ColInvalidReason"
        Me.ColInvalidReason.ReadOnly = True
        Me.ColInvalidReason.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ColInvalidReason.Width = 84
        '
        'BackgroundWorker1
        '
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 12)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(599, 23)
        Me.ProgressBar1.TabIndex = 2
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Location = New System.Drawing.Point(12, 40)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(57, 13)
        Me.lblProgress.TabIndex = 3
        Me.lblProgress.Text = "Progress : "
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.DgvItem)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 60)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(599, 403)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(171, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "The following data failed to import :"
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(536, 463)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(75, 23)
        Me.BtnClose.TabIndex = 5
        Me.BtnClose.Text = "&Close"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'TxtFilePath
        '
        Me.TxtFilePath.Location = New System.Drawing.Point(295, 238)
        Me.TxtFilePath.Name = "TxtFilePath"
        Me.TxtFilePath.Size = New System.Drawing.Size(33, 20)
        Me.TxtFilePath.TabIndex = 23
        Me.TxtFilePath.TabStop = False
        Me.TxtFilePath.Visible = False
        '
        'FrmLaporanVAT2025_Import
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(631, 73)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.BtnClose)
        Me.Controls.Add(Me.TxtFilePath)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(639, 100)
        Me.Name = "FrmLaporanVAT2025_Import"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Progress Import"
        CType(Me.DgvItem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DgvItem As System.Windows.Forms.DataGridView
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TxtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents ColNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColNPWP As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCustomerName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTransactionCode As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxInvoiceNumber As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxPeriod As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxYear As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxStatus As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTaxBase As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColOtherTaxBase As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVAT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPPnBM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSignedBy As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReference As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReportedBySeller As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReportedByVATCollector As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvalidReason As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
