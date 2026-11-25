Option Explicit On
Option Strict Off

Public Class ClsTransaksiAOP
    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumTipePurchaseBalance As Integer
        Automatic
        Manual
        Manual_Transfer_Invoice ' tidak terpakai lagi
        '2 berikut ini baru ditambahkan di v.1.0.5, berkaitan penggabungan jurnal sales dan jurnal sales return.
        Automatic_SR
        Manual_SR
        '2 berikut ini baru ditambahkan di v.1.6.0(end or beginning v.1.7.0), berkaitan dengan credit note dan debit note
        Credit_Note
        Debet_Note
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

    Public Enum EnumTaxExclude As Integer
        Exclude
        Include
    End Enum

    Public Enum EnumCalculatesBy As Integer
        AOP ''Invoice
        GRN ''Delivery_Order
        PO ''Sales_Order
    End Enum

    Public Enum EnumTipeAop As Integer
        AOP
        AOP_DP
    End Enum

    Public Enum EnumLokal As Integer
        Lokal
        Import
    End Enum

    Public Function GetSuppInvoiceNo(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiAOP1TableAdapters.tbltrans_aop_view0TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoTransaksi(vPrimary)) = False Then
                Return dta.ScalarQuery_GetNoTransaksi(vPrimary)
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

    Public Function GetJml(ByVal Criteria As String, _
       ByVal SearchCriteria As EnumSearchCriteria, ByVal TipeAOP As EnumTipeAop, ByVal ForeignSupplier As Integer, Optional _
       ByVal vTanggalDari As Date = (#1/1/1950#), Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If TipeAOP = EnumTipeAop.AOP_DP Then
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    Return DTA.ScalarQuery_GetjmlByNoTransaksiLIKE(Criteria + "%", TipeAOP).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    Return DTA.ScalarQuery_GetJmlbyTgltransaksi(TipeAOP, vTanggalDari, vTanggalSampai).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(TipeAOP, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(TipeAOP, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                    Return DTA.ScalarQuery_GetJmlProductSpecLIKE(TipeAOP, Criteria + "%").Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                    Return DTA.ScalarQuery_GetJmlBySuppCodeLIKE(Criteria + "%", TipeAOP).Value
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    Return DTA.ScalarQuery_GetJmlBySuppNameLIKE(TipeAOP, Criteria + "%").value
                ElseIf SearchCriteria = EnumSearchCriteria.By_IPO_Number Then
                    Return DTA.ScalarQuery_GetJmlByPONO_LIKE(Criteria + "%").Value
                End If
            Else
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE2(Criteria & "%", TipeAOP, ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    Return DTA.ScalarQuery_GetJmlByTanggalTransaksi2(TipeAOP, vTanggalDari, vTanggalSampai, ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                    Return DTA.ScalarQuery_GetJmlByProductSpecLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                    Return DTA.ScalarQuery_GetJmlBySuppCodeLIKE2(Criteria + "%", TipeAOP, ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    Return DTA.ScalarQuery_GetJmlBySuppNameLIKE2(TipeAOP, Criteria + "%", ForeignSupplier).Value
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    Return DTA.ScalarQuery_GetJmlByGRNNoLike2(TipeAOP, ForeignSupplier, Criteria + "%").Value
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
      ByVal DGV As DataGridView, ByVal TBL As DseTransaksiAOP.tbltrans_aop_view0DataTable, ByVal TipeAOP As EnumTipeAop, _
      ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal vForeignSupplier As Integer, ByVal vQuickPurchase As Boolean, _
      Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
      Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiAOPTableAdapters.tbltrans_aop_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim IdxColPO As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If TipeAOP = EnumTipeAop.AOP_DP Then
                If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                    QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%", TipeAOP)
                ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                    QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vTanggalDari, vTanggalSampai, TipeAOP)
                ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                    QueryResult = DTA.FillBy_SupplierCodeLIKE(TBL, Criteria + "%", TipeAOP)
                ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                    QueryResult = DTA.FillBy_SupplierNameLIKE(TBL, Criteria + "%", TipeAOP)
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                    QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%", TipeAOP)
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                    QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%", TipeAOP)
                ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                    QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%", TipeAOP)
                ElseIf SearchCriteria = EnumSearchCriteria.By_IPO_Number Then
                    QueryResult = DTA.FillBy_PONumberLIKE(TBL, Criteria + "%", TipeAOP)
                End If
            Else
                If vQuickPurchase = False Then
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA.FillBy_NoTransaksiLIKE2(TBL, Criteria + "%", TipeAOP, vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA.FillBy_TanggalDariSampai2(TBL, vTanggalDari, vTanggalSampai, TipeAOP, vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                        QueryResult = DTA.FillBy_SupplierCodeLIKE2(TBL, Criteria + "%", TipeAOP, vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                        QueryResult = DTA.FillBy_SupplierNameLIKE2(TBL, Criteria + "%", TipeAOP, vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA.FillBy_ProductCodeLIKE2(TBL, Criteria + "%", TipeAOP, vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA.FillBy_ProductDescLIKE2(TBL, Criteria + "%", TipeAOP, vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_ProductSpecLIKE2(TBL, TipeAOP, Criteria + "%", vForeignSupplier)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                        QueryResult = DTA.FillBy_NoGRNLike2(TBL, TipeAOP, vForeignSupplier, Criteria + "%")
                    End If
                Else
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA.FillBy_NoTransaksiLIKE3(TBL, Criteria + "%", TipeAOP)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA.FillBy_TanggalDariSampai3(TBL, vTanggalDari, vTanggalSampai, TipeAOP)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                        QueryResult = DTA.FillBy_SupplierCodeLIKE3(TBL, Criteria + "%", TipeAOP)
                    ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                        QueryResult = DTA.FillBy_SupplierNameLIKE3(TBL, Criteria + "%", TipeAOP)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA.FillBy_ProductCodeLIKE3(TBL, Criteria + "%", TipeAOP)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA.FillBy_ProductDescLIKE3(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_ProductSpecLIKE3(TBL, TipeAOP, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                        QueryResult = DTA.FillBy_NoGRNLike3(TBL, TipeAOP, Criteria + "%")
                    End If
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                IdxColPO = DGV.ColumnCount - 1
                If TipeAOP = EnumTipeAop.AOP Then
                    DGV.Columns(IdxColPO).Visible = False
                Else
                    DGV.Columns(IdxColPO).Visible = True
                End If
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

    Public Function GetPrimaryAkunAP(ByVal vPrimaryAOP As Double) As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            Return ObjCur.GetPrimaryAkunAP(Me.GetPrimaryCurrency(vPrimaryAOP))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCur = Nothing
        End Try
    End Function

    Public Function GetSuppCode(ByVal vPrimaryAOP As Double) As String
        Dim DTA As New DseTransaksiAOP1TableAdapters.tbltrans_aop_view0TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSuppCode(vPrimaryAOP)) = False Then
                Return DTA.ScalarQuery_GetSuppCode(vPrimaryAOP)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSuppName(ByVal vPrimaryAOP As Double) As String
        Dim DTA As New DseTransaksiAOP1TableAdapters.tbltrans_aop_view0TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSuppName(vPrimaryAOP)) = False Then
                Return DTA.ScalarQuery_GetSuppName(vPrimaryAOP)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim dta As New DseTransaksiAOP1TableAdapters.tbltrans_aop_view0TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).HasValue Then
                If dta.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value > 0 Then
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
            dta.Dispose()
        End Try
    End Function

    Public Function GetAOPDate(ByVal vPrimary As Double) As Date
        Dim dta As New DseTransaksiAOP1TableAdapters.tbltrans_aop_view0TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetAOPDate(vPrimary).HasValue Then
                Return dta.ScalarQuery_GetAOPDate(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryAOP As Double) As Double
        Dim dta As New DseTransaksiAOP1TableAdapters.tbltrans_aopTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimaryCurrency(vPrimaryAOP).HasValue Then
                Return dta.ScalarQuery_GetPrimaryCurrency(vPrimaryAOP).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vAOPNumber As String) As Double
        Dim dta As New DseTransaksiAOP1TableAdapters.tbltrans_aop_view0TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimary(vAOPNumber).HasValue Then
                Return dta.ScalarQuery_GetPrimary(vAOPNumber).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim dta As New DseTransaksiAOP1TableAdapters.tbltrans_aop_view0TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoTransaksi(vPrimary)) = False Then
                Return dta.ScalarQuery_GetNoTransaksi(vPrimary)
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
End Class
