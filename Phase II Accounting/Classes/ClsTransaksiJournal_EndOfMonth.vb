Option Explicit On
Option Strict Off

Public Class ClsTransaksiJournal_EndOfMonth
    Public Sub IsikanDGVCurrencyExchangeRate(ByVal vDGV As DataGridView, ByVal vFrm As FrmTransaksiJournal_EndOfMonth, ByVal vTglTransaksi As Date)
        Dim DTA As New DseTransaksiJournal_EndOfMonthTableAdapters.tblmaster_currencyTableAdapter
        Dim TBL As New DseTransaksiJournal_EndOfMonth.tblmaster_currencyDataTable
        Dim row As DseTransaksiJournal_EndOfMonth.tblmaster_currencyRow
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA1 As New DseTransaksiJournal_EndOfMonthTableAdapters.tbltrans_salesinvoice_taxrate_itemTableAdapter
        Dim TBL1 As New DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoice_taxrate_itemDataTable
        Dim row1 As DseTransaksiJournal_EndOfMonth.tbltrans_salesinvoice_taxrate_itemRow
        Dim Filter As String = ""

        Try
            'kosongkan dgv
            vDGV.Rows.Clear()

            'isikan berdasarkan default currency
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)
            For Each row In TBL.Rows
                If row.Primary <> Core_Module.PrimaryBaseCurrency Then
                    vDGV.Rows.Add()
                    rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)

                    rowDGV.Cells(vFrm.ColItemExc_Currency.Index).Value = row.Currency_Code
                    rowDGV.Cells(vFrm.ColItemExc_PrimaryCurrency.Index).Value = row.Primary
                    If row.Default_Exchange_Rate = 0 Then
                        row.Default_Exchange_Rate = 1
                    End If
                    rowDGV.Cells(vFrm.ColItemExc_ExchangeRate.Index).Value = ObjCD.DigitGrouping(row.Default_Exchange_Rate, Core_Module.DefPrecExchangeRate)
                    rowDGV.Cells(vFrm.ColItemExc_Remarks.Index).Value = ""
                End If
            Next

            'isikan berdasarkan tax exchange rate, bila ada 
            DTA1.Fill(TBL1, vTglTransaksi)
            For Each rowDGV In vDGV.Rows
                Filter = "[primary currency] = " + rowDGV.Cells(vFrm.ColItemExc_PrimaryCurrency.Index).Value.ToString
                TBL1.DefaultView.RowFilter = Filter

                If TBL1.DefaultView.Count > 0 Then
                    row1 = TBL1.DefaultView(0).Row
                    If row1.rate <> 1 And row1.rate <> 0 Then
                        rowDGV.Cells(vFrm.ColItemExc_ExchangeRate.Index).Value = ObjCD.DigitGrouping(row1.rate, Core_Module.DefPrecExchangeRate)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub
End Class
