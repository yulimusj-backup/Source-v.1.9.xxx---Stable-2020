Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesInvoice
    Dim ObjCD As New Core_Module.ClsCekData

    'elemen yang digunakan pada automatic invoice transfer
    Public Enum EnumTipeSalesBalance As Integer
        Automatic
        Manual
        Manual_Transfer '--> tidak digunakan lagi, hanya pada cmsi saja
        Automatic_SR
        Manual_SR
    End Enum

    Public Enum EnumTipeJournal As Integer
        GeneralJournal
        AdjustmenJournal
        ClosingJournal
        SalesJournal
        SalesReturnJournal
        SalesDownPayment
        PurchaseJournal
        PurchaseReturnJournal
        PurchaseDownPaymentJournal
        SalesReceivementJournal
        SalesReturnPaymentJournal
        PurchasePaymentJournal
        PurchaseReturnReceivement
    End Enum

    Public Enum Enum_Tipe_Setting
        None
        Debet
        Kredit
    End Enum

    Public Enum EnumTipeInvoice As Integer
        InvoiceDP
        SalesInvoice
        Credit_Note
        Debet_Note
    End Enum
    Public Overridable Function IsiDGV(ByVal Criteria As String, _
     ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesInvoice.tbltrans_salesinvoice_viewDataTable, _
     ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal TipeInvoice As EnumTipeInvoice, ByVal vIsQuick As Boolean, _
     Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
     Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_viewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalTransaksi(TBL, TipeInvoice, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.TaxInvoiceNo Then
                QueryResult = DTA.FillBy_NoTaxInvoiceLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                QueryResult = DTA.FillBy_CustName(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpec(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDeliveryOrderNo Then
                QueryResult = DTA.FillBy_DONumberLIKE(TBL, TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then
                If vIsQuick = True Then
                    QueryResult = DTA.FillBy_SalesPersonCodeLIKE1(TBL, Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        QueryResult = DTA.FillBy_SalesPersonCodeLIKE2(TBL, Criteria + "%")
                    Else
                        QueryResult = DTA.FillBy_SalesPersonCodeLIKE3(TBL, Criteria + "%")
                    End If
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                If vIsQuick = True Then
                    QueryResult = DTA.FillBy_SalesPersonDescLIKE1(TBL, Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        QueryResult = DTA.FillBy_SalesPersonDescLIKE2(TBL, Criteria + "%")
                    Else
                        QueryResult = DTA.FillBy_SalesPersonDescLIKE3(TBL, Criteria + "%")
                    End If
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + ObjCD.DigitGrouping(TS.TotalMilliseconds, 2, MidpointRounding.AwayFromZero) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function
    Public Function GetJml(ByVal Criteria As String, _
      ByVal SearchCriteria As EnumSearchCriteria, ByVal TipeInvoice As EnumTipeInvoice, ByVal vIsQuick As Boolean, Optional _
      ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_viewTableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%", TipeInvoice)
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalTransaksi(TipeInvoice, vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(TipeInvoice, Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDescLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%", TipeInvoice)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%", TipeInvoice)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%", TipeInvoice)
            ElseIf SearchCriteria = EnumSearchCriteria.TaxInvoiceNo Then
                Return DTA.ScalarQuery_GetJmlByTaxInvoiceLIKE(Criteria + "%", TipeInvoice)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDeliveryOrderNo Then
                Return DTA.ScalarQuery_GetJmlBy_DONumberLIKE(TipeInvoice, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then
                If vIsQuick = True Then
                    'quikc invoice, cari langsung dari header
                    Return DTA.ScalarQuery_GetJmlBySalesPersonCode1(Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        Return DTA.ScalarQuery_GetJmlBySalesPersonCode2(Criteria + "%")
                    Else
                        Return DTA.ScalarQuery_GetJmlBySalesPersonCode3(Criteria + "%")
                    End If
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                If vIsQuick = True Then
                    'quikc invoice, cari langsung dari header
                    Return DTA.ScalarQuery_GetJmlBySalesPersonDesc1(Criteria + "%")
                Else
                    If TipeInvoice = EnumTipeInvoice.InvoiceDP Then
                        Return DTA.ScalarQuery_GetJmlBySalesPersonDesc2(Criteria + "%")
                    Else
                        Return DTA.ScalarQuery_GetJmlBySalesPersonDesc3(Criteria + "%")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetInvoiceNumber(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoiceTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoInvoice(vPrimary)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoInvoice(vPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_view1TableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByNoInvoice(vNoTransaksi).HasValue Then
                If DTA.ScalarQuery_GetJmlByNoInvoice(vNoTransaksi).Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryInvoice As String) As Double
        Dim DTA As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_view1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryInvoice).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryInvoice).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vnoTrans As String) As Double
        Dim dta As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_view1TableAdapter
        Dim RetVal As Nullable(Of Double)

        Try
            RetVal = dta.ScalarQuery_getprimary(vnoTrans)
            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetAmount(ByVal vPrimary As Double) As Double
        Dim dta As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_view1TableAdapter
        Try
            Return dta.ScalarQuery_GetAmount(vPrimary).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCurrency(ByVal vPrimaryInv As Double) As String
        Dim dta As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_view1TableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetCurrencyCode(vPrimaryInv)) = False Then
                Return dta.ScalarQuery_GetCurrencyCode(vPrimaryInv)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetInvoiceDate(ByVal vPrimary As Double) As Date
        Dim dta As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihan_view1TableAdapter
        Try
            If dta.ScalarQuery_GetInvoiceDate(vPrimary).HasValue Then
                Return dta.ScalarQuery_GetInvoiceDate(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetDueDate(ByVal vPrimary As Double) As Date
        Dim dta As New DseTransaksiTagihanTableAdapters.tbltrans_surattagihan_view1TableAdapter
        Try
            If dta.ScalarQuery_GetDueDate(vPrimary).HasValue Then
                Return dta.ScalarQuery_GetDueDate(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCustomerCode(ByVal vPrimaryInvoice As Double) As String
        Dim dta As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_view1TableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetCustCode(vPrimaryInvoice)) = False Then
                Return dta.ScalarQuery_GetCustCode(vPrimaryInvoice)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetCustomerName(ByVal vPrimaryInvoice As Double) As String
        Dim dta As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice_view1TableAdapter
        Try
            If String.IsNullOrEmpty(dta.ScalarQuery_GetCustName(vPrimaryInvoice)) = False Then
                Return dta.ScalarQuery_GetCustName(vPrimaryInvoice)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryAkunAR(ByVal vPrimaryInv As Double) As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            Return ObjCur.GetPrimaryAkunAR(Me.GetPrimaryCurrency(vPrimaryInv))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Function

End Class
