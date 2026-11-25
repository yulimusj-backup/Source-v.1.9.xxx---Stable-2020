<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTransaksiJournal_Sales
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.dgvSalesBalance = New System.Windows.Forms.DataGridView
        Me.ColInvoiceNo = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvoiceDate = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDueDateInvoice = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColInvRemarks = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColSalesValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColVatSalesValue = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColDutystamp = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTotalSales_VAT_Value = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrencyBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivedOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivementOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColExhangeRateBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPayInv = New System.Windows.Forms.DataGridViewButtonColumn
        Me.ColBalanceSales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceived = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColReceivement = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimarySalesBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryDocumenAcuan_Sales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColTipeDocumentAcuan_Sales = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCust = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryCurrencyBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColCurrencyDefPrec = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalance1 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceOriginal01 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColBalanceOriginal02 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColPrimaryAkunBalance = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.dgvSalesBalance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvSalesBalance
        '
        Me.dgvSalesBalance.AllowUserToAddRows = False
        Me.dgvSalesBalance.AllowUserToDeleteRows = False
        Me.dgvSalesBalance.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSalesBalance.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSalesBalance.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ColInvoiceNo, Me.ColInvoiceDate, Me.ColDueDateInvoice, Me.ColInvRemarks, Me.ColSalesValue, Me.ColVatSalesValue, Me.ColDutystamp, Me.ColTotalSales_VAT_Value, Me.ColCurrencyBalance, Me.ColBalanceOriginal, Me.ColReceivedOriginal, Me.ColReceivementOriginal, Me.ColExhangeRateBalance, Me.ColPayInv, Me.ColBalanceSales, Me.ColReceived, Me.ColReceivement, Me.ColPrimarySalesBalance, Me.ColPrimaryDocumenAcuan_Sales, Me.ColTipeDocumentAcuan_Sales, Me.ColPrimaryCust, Me.ColPrimaryCurrencyBalance, Me.ColCurrencyDefPrec, Me.ColBalance1, Me.ColBalanceOriginal01, Me.ColBalanceOriginal02, Me.ColPrimaryAkunBalance})
        Me.dgvSalesBalance.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvSalesBalance.Location = New System.Drawing.Point(0, 0)
        Me.dgvSalesBalance.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvSalesBalance.MultiSelect = False
        Me.dgvSalesBalance.Name = "dgvSalesBalance"
        Me.dgvSalesBalance.RowTemplate.Height = 24
        Me.dgvSalesBalance.Size = New System.Drawing.Size(292, 263)
        Me.dgvSalesBalance.TabIndex = 3
        '
        'ColInvoiceNo
        '
        Me.ColInvoiceNo.Frozen = True
        Me.ColInvoiceNo.HeaderText = "Invoice Number"
        Me.ColInvoiceNo.Name = "ColInvoiceNo"
        Me.ColInvoiceNo.Width = 131
        '
        'ColInvoiceDate
        '
        Me.ColInvoiceDate.HeaderText = "Invoice Date"
        Me.ColInvoiceDate.Name = "ColInvoiceDate"
        Me.ColInvoiceDate.Visible = False
        Me.ColInvoiceDate.Width = 93
        '
        'ColDueDateInvoice
        '
        Me.ColDueDateInvoice.HeaderText = "Due Date"
        Me.ColDueDateInvoice.Name = "ColDueDateInvoice"
        Me.ColDueDateInvoice.Width = 93
        '
        'ColInvRemarks
        '
        Me.ColInvRemarks.HeaderText = "Remarks"
        Me.ColInvRemarks.Name = "ColInvRemarks"
        Me.ColInvRemarks.Width = 89
        '
        'ColSalesValue
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColSalesValue.DefaultCellStyle = DataGridViewCellStyle1
        Me.ColSalesValue.HeaderText = "Sales"
        Me.ColSalesValue.Name = "ColSalesValue"
        Me.ColSalesValue.Visible = False
        Me.ColSalesValue.Width = 58
        '
        'ColVatSalesValue
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColVatSalesValue.DefaultCellStyle = DataGridViewCellStyle2
        Me.ColVatSalesValue.HeaderText = "VAT"
        Me.ColVatSalesValue.Name = "ColVatSalesValue"
        Me.ColVatSalesValue.Visible = False
        Me.ColVatSalesValue.Width = 53
        '
        'ColDutystamp
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColDutystamp.DefaultCellStyle = DataGridViewCellStyle3
        Me.ColDutystamp.HeaderText = "Duty stamp"
        Me.ColDutystamp.Name = "ColDutystamp"
        Me.ColDutystamp.Visible = False
        Me.ColDutystamp.Width = 85
        '
        'ColTotalSales_VAT_Value
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColTotalSales_VAT_Value.DefaultCellStyle = DataGridViewCellStyle4
        Me.ColTotalSales_VAT_Value.HeaderText = "Sales + VAT + Duty stamp"
        Me.ColTotalSales_VAT_Value.Name = "ColTotalSales_VAT_Value"
        Me.ColTotalSales_VAT_Value.Visible = False
        Me.ColTotalSales_VAT_Value.Width = 156
        '
        'ColCurrencyBalance
        '
        Me.ColCurrencyBalance.HeaderText = "Currency"
        Me.ColCurrencyBalance.Name = "ColCurrencyBalance"
        Me.ColCurrencyBalance.Width = 90
        '
        'ColBalanceOriginal
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalanceOriginal.DefaultCellStyle = DataGridViewCellStyle5
        Me.ColBalanceOriginal.HeaderText = "Balance"
        Me.ColBalanceOriginal.Name = "ColBalanceOriginal"
        Me.ColBalanceOriginal.Width = 84
        '
        'ColReceivedOriginal
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceivedOriginal.DefaultCellStyle = DataGridViewCellStyle6
        Me.ColReceivedOriginal.HeaderText = "Received"
        Me.ColReceivedOriginal.Name = "ColReceivedOriginal"
        Me.ColReceivedOriginal.Visible = False
        Me.ColReceivedOriginal.Width = 78
        '
        'ColReceivementOriginal
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceivementOriginal.DefaultCellStyle = DataGridViewCellStyle7
        Me.ColReceivementOriginal.HeaderText = "Pay"
        Me.ColReceivementOriginal.Name = "ColReceivementOriginal"
        Me.ColReceivementOriginal.Width = 57
        '
        'ColExhangeRateBalance
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColExhangeRateBalance.DefaultCellStyle = DataGridViewCellStyle8
        Me.ColExhangeRateBalance.HeaderText = "Exchange Rate"
        Me.ColExhangeRateBalance.Name = "ColExhangeRateBalance"
        Me.ColExhangeRateBalance.Width = 129
        '
        'ColPayInv
        '
        Me.ColPayInv.HeaderText = ""
        Me.ColPayInv.Name = "ColPayInv"
        Me.ColPayInv.Text = "Pay"
        Me.ColPayInv.Width = 5
        '
        'ColBalanceSales
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColBalanceSales.DefaultCellStyle = DataGridViewCellStyle9
        Me.ColBalanceSales.HeaderText = "Balance IDR"
        Me.ColBalanceSales.Name = "ColBalanceSales"
        Me.ColBalanceSales.Visible = False
        Me.ColBalanceSales.Width = 93
        '
        'ColReceived
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceived.DefaultCellStyle = DataGridViewCellStyle10
        Me.ColReceived.HeaderText = "Received IDR"
        Me.ColReceived.Name = "ColReceived"
        Me.ColReceived.Visible = False
        '
        'ColReceivement
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ColReceivement.DefaultCellStyle = DataGridViewCellStyle11
        Me.ColReceivement.HeaderText = "Receivement IDR"
        Me.ColReceivement.Name = "ColReceivement"
        Me.ColReceivement.Visible = False
        Me.ColReceivement.Width = 117
        '
        'ColPrimarySalesBalance
        '
        Me.ColPrimarySalesBalance.HeaderText = "PrimarySalesBalance"
        Me.ColPrimarySalesBalance.Name = "ColPrimarySalesBalance"
        Me.ColPrimarySalesBalance.Visible = False
        Me.ColPrimarySalesBalance.Width = 131
        '
        'ColPrimaryDocumenAcuan_Sales
        '
        Me.ColPrimaryDocumenAcuan_Sales.HeaderText = "Primary Documen Acuan"
        Me.ColPrimaryDocumenAcuan_Sales.Name = "ColPrimaryDocumenAcuan_Sales"
        Me.ColPrimaryDocumenAcuan_Sales.Visible = False
        Me.ColPrimaryDocumenAcuan_Sales.Width = 149
        '
        'ColTipeDocumentAcuan_Sales
        '
        Me.ColTipeDocumentAcuan_Sales.HeaderText = "TipeDocumenAcuan"
        Me.ColTipeDocumentAcuan_Sales.Name = "ColTipeDocumentAcuan_Sales"
        Me.ColTipeDocumentAcuan_Sales.Visible = False
        Me.ColTipeDocumentAcuan_Sales.Width = 130
        '
        'ColPrimaryCust
        '
        Me.ColPrimaryCust.HeaderText = "Primary Cust"
        Me.ColPrimaryCust.Name = "ColPrimaryCust"
        Me.ColPrimaryCust.Visible = False
        Me.ColPrimaryCust.Width = 90
        '
        'ColPrimaryCurrencyBalance
        '
        Me.ColPrimaryCurrencyBalance.HeaderText = "Primary Currency"
        Me.ColPrimaryCurrencyBalance.Name = "ColPrimaryCurrencyBalance"
        Me.ColPrimaryCurrencyBalance.Visible = False
        Me.ColPrimaryCurrencyBalance.Width = 111
        '
        'ColCurrencyDefPrec
        '
        Me.ColCurrencyDefPrec.HeaderText = "Default Prec Cur"
        Me.ColCurrencyDefPrec.Name = "ColCurrencyDefPrec"
        Me.ColCurrencyDefPrec.Visible = False
        Me.ColCurrencyDefPrec.Width = 110
        '
        'ColBalance1
        '
        Me.ColBalance1.HeaderText = "Balance1"
        Me.ColBalance1.Name = "ColBalance1"
        Me.ColBalance1.Visible = False
        Me.ColBalance1.Width = 77
        '
        'ColBalanceOriginal01
        '
        Me.ColBalanceOriginal01.HeaderText = "BalanceOriginal01"
        Me.ColBalanceOriginal01.Name = "ColBalanceOriginal01"
        Me.ColBalanceOriginal01.Visible = False
        Me.ColBalanceOriginal01.Width = 118
        '
        'ColBalanceOriginal02
        '
        Me.ColBalanceOriginal02.HeaderText = "Balance Original 02"
        Me.ColBalanceOriginal02.Name = "ColBalanceOriginal02"
        Me.ColBalanceOriginal02.Visible = False
        Me.ColBalanceOriginal02.Width = 124
        '
        'ColPrimaryAkunBalance
        '
        Me.ColPrimaryAkunBalance.HeaderText = "Primary Akun Balance"
        Me.ColPrimaryAkunBalance.Name = "ColPrimaryAkunBalance"
        Me.ColPrimaryAkunBalance.Visible = False
        Me.ColPrimaryAkunBalance.Width = 136
        '
        'FrmTransaksiJournal_Sales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 263)
        Me.Controls.Add(Me.dgvSalesBalance)
        Me.Name = "FrmTransaksiJournal_Sales"
        Me.Text = "FrmTransaksiJournal_Sales"
        CType(Me.dgvSalesBalance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents dgvSalesBalance As System.Windows.Forms.DataGridView
    Friend WithEvents ColInvoiceNo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvoiceDate As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDueDateInvoice As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColInvRemarks As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColSalesValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColVatSalesValue As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColDutystamp As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTotalSales_VAT_Value As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrencyBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivedOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivementOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColExhangeRateBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPayInv As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents ColBalanceSales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceived As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColReceivement As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimarySalesBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryDocumenAcuan_Sales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColTipeDocumentAcuan_Sales As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCust As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryCurrencyBalance As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColCurrencyDefPrec As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalance1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceOriginal01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColBalanceOriginal02 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColPrimaryAkunBalance As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
