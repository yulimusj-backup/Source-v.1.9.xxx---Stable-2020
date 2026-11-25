Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesBalance
    Public Enum EnumTipePayment As Integer
        SalesInvoice_SalesInvoiceDP
        SalesReturn
        AOP_AOPDP
        Purchase_Return
    End Enum
    Public Enum EnumTipeSalesBalance As Integer
        Automatic
        Manual
        Manual_Transfer_Invoice
        '2 berikut ini baru ditambahkan di v.1.0.5, berkaitan penggabungan jurnal sales dan jurnal sales return.
        Automatic_SR
        Manual_SR
    End Enum

    Public Sub IsikanListCustDenganBalance(ByVal vListPrimary As List(Of Double))

        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' fungsi ini digunakan form sales journal untuk mengisikan daftar balance invoice
    ''' </summary>
    ''' <param name="vDGV"></param>
    ''' <remarks></remarks>
    Public Sub IsiDGVBalanceTanpaJurnal(ByVal vDGV As DataGridView, ByVal TBL As DseTransaksiSalesBalance.balance_invoice_viewDataTable, _
    ByVal ij As FrmTransaksiJournal_Sales, ByVal ObjCD As Core_Module.ClsCekData, ByVal vPrimaryCust As Double, _
    ByVal Default_Prec As Integer, ByVal TBLSR As DseTransaksiSalesBalance.balance_sr_viewDataTable)
        Dim DTA As New DseTransaksiSalesBalanceTableAdapters.balance_invoice_viewTableAdapter
        Dim TBLDummy As New DseTransaksiSalesBalance.balance_invoice_viewDataTable
        Dim Filter As String = ""
        Dim TipePayment As EnumTipePayment = EnumTipePayment.SalesInvoice_SalesInvoiceDP
        Dim row As DseTransaksiSalesBalance.balance_invoice_viewRow
        Dim rowdgv As DataGridViewRow
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ketemu As Boolean = False

        Dim TipeBalance1 As EnumTipeSalesBalance = EnumTipeSalesBalance.Manual_SR
        Dim TipeBalance2 As EnumTipeSalesBalance = EnumTipeSalesBalance.Automatic_SR
        Dim ObjSRBalance As New ClsTransaksiSalesReturnBalance
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim AdaNonIDR As Boolean = False

        Try
            'isikan data dari database
            DTA.ClearBeforeFill = False
            TBL.Rows.Clear()

            If vPrimaryCust = -1 Then
                DTA.FillBy_All(TBL)
            Else
                DTA.FillBy_PrimaryCust(TBL, vPrimaryCust)
            End If

            'khusus untuk sales return maka isikan no transaksi invoice, no sr, dan remarksnya
            For Each row In TBL.Rows
                If row.Tipe = TipeBalance1 Then
                    'manual sr, yang muncul pada no invoice adalah no transaksi sales return balance
                    ObjSRBalance.GetNoTransaksi_Remarks_Tanggal(row.Invoice_Number, row.Remarks, row.Invoice_Date, row.Primary_Invoice)
                End If
                If row.Tipe = TipeBalance2 Then
                    ObjSR.GetNoTransaksi_Remarks_NoInvoice(row.Invoice_Number, row.Remarks, row.Primary_Invoice, row.Invoice_Date)
                End If
            Next

            'isikan ke dalam datagridview
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.Rows.Clear()
            TBL.DefaultView.RowFilter = ""
            TBL.DefaultView.Sort = "[Invoice Date] , [Invoice Number], primary"

            For Each dv As Data.DataRowView In TBL.DefaultView
                row = dv.Row
                If row.Balance <> 0 Or row.Receivement <> 0 Then
                    vDGV.Rows.Add()
                    rowdgv = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowdgv
                        If row.primary_currency <> Core_Module.PrimaryBaseCurrency Then
                            AdaNonIDR = True
                        End If
                        If row.Sales = 0 Then
                            row.Invoice_Number = row.Invoice_Number + " (Tax + Dutystamp)"
                        End If
                        .Cells(ij.ColInvoiceNo.Index).Value = row.Invoice_Number
                        .Cells(ij.ColInvoiceDate.Index).Value = Format(row.Invoice_Date, "dd MMM yyyy")
                        .Cells(ij.ColSalesValue.Index).Value = ObjCD.DigitGrouping(row.Sales, Default_Prec)
                        .Cells(ij.ColVatSalesValue.Index).Value = ObjCD.DigitGrouping(row.Sales_VAT, Default_Prec)
                        .Cells(ij.ColDutystamp.Index).Value = ObjCD.DigitGrouping(row.Duty_stamp, Default_Prec)
                        .Cells(ij.ColTotalSales_VAT_Value.Index).Value = ObjCD.DigitGrouping(row.Sales + row.Sales_VAT + row.Duty_stamp, Default_Prec)
                        .Cells(ij.ColReceived.Index).Value = ObjCD.DigitGrouping(row.Received, Default_Prec)
                        .Cells(ij.ColBalanceSales.Index).Value = ObjCD.DigitGrouping(row.Balance, Default_Prec)
                        .Cells(ij.ColReceivement.Index).Value = ObjCD.DigitGrouping(Math.Round(row.Receivement_Original * row.exchange_rate, Default_Prec), Default_Prec)
                        .Cells(ij.ColPrimarySalesBalance.Index).Value = row.Primary
                        .Cells(ij.ColDueDateInvoice.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
                        .Cells(ij.ColPrimaryDocumenAcuan_Sales.Index).Value = row.Primary_Invoice
                        .Cells(ij.ColTipeDocumentAcuan_Sales.Index).Value = row.Tipe
                        .Cells(ij.ColInvRemarks.Index).Value = row.Remarks
                        .Cells(ij.ColPayInv.Index).Value = "Full Payment"
                        .Cells(ij.ColPrimaryCust.Index).Value = row.Primary_Customer
                        .Cells(ij.ColCurrencyBalance.Index).Value = row.Currency_Code
                        .Cells(ij.ColCurrencyDefPrec.Index).Value = row.Default_Precision
                        .Cells(ij.ColPrimaryCurrencyBalance.Index).Value = row.primary_currency
                        .Cells(ij.ColBalanceOriginal.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                        .Cells(ij.ColReceivementOriginal.Index).Value = ObjCD.DigitGrouping(row.Receivement_Original, row.Default_Precision)
                        .Cells(ij.ColReceivedOriginal.Index).Value = ObjCD.DigitGrouping(row.received_original, row.Default_Precision)
                        .Cells(ij.ColExhangeRateBalance.Index).Value = ObjCD.DigitGrouping(row.exchange_rate, Core_Module.DefPrecExchangeRate)
                        .Cells(ij.ColBalance1.Index).Value = CDec(.Cells(ij.ColBalanceSales.Index).Value) + _
                        Math.Round(CDec(row.Receivement_Original * row.exchange_rate), Default_Prec) + _
                        CDec(.Cells(ij.ColReceived.Index).Value)
                        ' nilai diatas ini sebenarnya sama dengan received, 
                        'hanya saja menjadi salah untuk journal yang sudah direaluasi kemudian belum dibetulkan, diganti untuk mengakomodasi proses pembetulan
                        .Cells(ij.ColBalanceOriginal02.Index).Value = .Cells(ij.ColBalanceOriginal.Index).Value
                        .Cells(ij.ColBalanceOriginal01.Index).Value = CDec(.Cells(ij.ColBalanceOriginal.Index).Value) + _
                        CDec(.Cells(ij.ColReceivementOriginal.Index).Value) + CDec(.Cells(ij.ColReceivedOriginal.Index).Value)
                        .Cells(ij.ColPrimaryAkunBalance.Index).Value = row.primary_akun_AR
                    End With
                End If
            Next

            'jika semua baris currencynya IDR, maka exchage rate tidak perlu ada
            If AdaNonIDR = True Then
                ij.ColExhangeRateBalance.Visible = True
            Else
                ij.ColExhangeRateBalance.Visible = False
            End If

            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Sub
End Class
