Option Explicit On
Option Strict Off

Public Class ClsTrans_CutOffDB_Core
    Public Function GetPrimaryUpdate(ByVal vPrimaryOld As Double, _
    ByVal vTblPrimary As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable) As Double
        Dim rowSO As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateRow

        Try
            vTblPrimary.DefaultView.RowFilter = "[Primary Old] = " & vPrimaryOld
            If vTblPrimary.DefaultView.Count > 0 Then
                rowSO = vTblPrimary.DefaultView(0).Row
                Return rowSO.Primary_New
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub AddPrimaryUpdate(ByVal vPrimaryOld As Double, ByVal vPrimaryNew As Double, _
    ByRef vTblPrimary As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateDataTable)
        Dim row As Core_Module.DseTrans_CutOffDB_Core.Tbl_PrimaryUpdateRow

        Try
            row = vTblPrimary.NewTbl_PrimaryUpdateRow
            row.Primary_Old = vPrimaryOld
            row.Primary_New = vPrimaryNew
            vTblPrimary.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetAopNo(ByVal vPrimaryAop As Double, _
        ByVal vTblAopNo As Core_Module.DseTrans_CutOffDB_Core.tbltrans_aopDataTable, _
        Optional ByRef vTglAop As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim rowAop As Core_Module.DseTrans_CutOffDB_Core.tbltrans_aopRow

        Try
            vTblAopNo.DefaultView.RowFilter = "[Primary] = " & vPrimaryAop
            If vTblAopNo.DefaultView.Count > 0 Then
                rowAop = vTblAopNo.DefaultView(0).Row
                vTglAop = rowAop.Tanggal_Transaksi
                vPayTerm = rowAop.Payment_Term
                Return rowAop.No_Transaksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetPurchReturnNo(ByVal vPrimaryPR As Double, _
    ByVal vTblPRNo As Core_Module.DseTrans_CutOffDB_Core.tbltrans_purchasereturnDataTable, _
    Optional ByRef vTglPR As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim rowPR As Core_Module.DseTrans_CutOffDB_Core.tbltrans_purchasereturnRow

        Try
            vTblPRNo.DefaultView.RowFilter = "[Primary] = " & vPrimaryPR
            If vTblPRNo.DefaultView.Count > 0 Then
                rowPR = vTblPRNo.DefaultView(0).Row
                vTglPR = rowPR.Tanggal_Transaksi
                vPayTerm = rowPR.Payment_Term
                Return rowPR.No_Transaksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetInvNo(ByVal vPrimaryInv As Double, _
    ByVal vTblInvNo As Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoDataTable, _
    Optional ByRef vTglInv As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim rowInv As Core_Module.DseTrans_CutOffDB_Core.tbl_InvNoRow

        Try
            vTblInvNo.DefaultView.RowFilter = "[Primary] = " & vPrimaryInv
            If vTblInvNo.DefaultView.Count > 0 Then
                rowInv = vTblInvNo.DefaultView(0).Row
                vTglInv = rowInv.Tanggal_Transaksi
                vPayTerm = rowInv.Payment_Term
                Return rowInv.No_Transaksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetSalesReturnNo(ByVal vPrimarySR As Double, _
    ByVal vTblSRNo As Core_Module.DseTrans_CutOffDB_Core.tbl_SRNoDataTable, _
    Optional ByRef vTglSR As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim rowSR As Core_Module.DseTrans_CutOffDB_Core.tbl_SRNoRow

        Try
            vTblSRNo.DefaultView.RowFilter = "[Primary] = " & vPrimarySR
            If vTblSRNo.DefaultView.Count > 0 Then
                rowSR = vTblSRNo.DefaultView(0).Row
                vTglSR = rowSR.Tanggal_Transaksi
                vPayTerm = rowSR.Payment_Term
                Return rowSR.No_Transaksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetCDNoteNo(ByVal vPrimaryCD As Double, _
    ByVal vTblCDNo As Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoDataTable, _
    Optional ByRef vTglCDNote As Date = #1/1/1950#, Optional ByRef vPayTerm As Integer = 0) As String
        Dim rowCD As Core_Module.DseTrans_CutOffDB_Core.tbl_CDNoRow

        Try
            vTblCDNo.DefaultView.RowFilter = "[Primary] = " & vPrimaryCD
            If vTblCDNo.DefaultView.Count > 0 Then
                rowCD = vTblCDNo.DefaultView(0).Row
                vTglCDNote = rowCD.tgl
                vPayTerm = rowCD.Payment_term
                Return rowCD.no_transksi
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    ' Public Function GetPrimaryBalanceOrAcuanSales(ByVal vPrimaryCust As Double, ByVal vPrimaryInv As Double, _
    ' ByVal vTipe As Integer, ByVal vTblBalSales As Core_Module.DseTrans_CutOffDB_Core.tbltrans_balance_salesDataTable, _
    'Optional ByVal vNoTransAcuan As String = "", Optional ByRef PrimaryBalance As Double = -1) As Double
    '     Dim rowBalSales As Core_Module.DseTrans_CutOffDB_Core.tbltrans_balance_salesRow

    '     Try
    '         If vNoTransAcuan <> "" Then
    '             vTblBalSales.DefaultView.RowFilter = "[Invoice Number Manual] = '" & vNoTransAcuan.ToString & "' AND [Tipe] = " & vTipe
    '             If vTblBalSales.DefaultView.Count > 0 Then
    '                 rowBalSales = vTblBalSales.DefaultView(0).Row
    '                 PrimaryBalance = rowBalSales.Primary
    '                 Return rowBalSales.Primary_Invoice
    '             Else
    '                 Return -1
    '             End If
    '         Else
    '             vTblBalSales.DefaultView.RowFilter = "[Primary Customer] = " & vPrimaryCust & " AND [Primary Invoice] = " & vPrimaryInv & " AND [Tipe] = " & vTipe
    '             If vTblBalSales.DefaultView.Count > 0 Then
    '                 rowBalSales = vTblBalSales.DefaultView(0).Row
    '                 Return rowBalSales.Primary
    '             Else
    '                 Return -1
    '             End If
    '         End If
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     End Try
    ' End Function

    Public Function GetPrimaryGrn(ByVal vPrimaryStockInPurchase As Double, _
    ByVal vTblPrimary As DseTrans_CutOffDB_Core.tbltrans_grnDataTable) As Double
        Dim row As Core_Module.DseTrans_CutOffDB_Core.tbltrans_grnRow

        Try
            vTblPrimary.DefaultView.RowFilter = "[Primary Stock In Purchase] = " & vPrimaryStockInPurchase
            If vTblPrimary.DefaultView.Count > 0 Then
                row = vTblPrimary.DefaultView(0).Row
                Return row.Primary
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNo_DOByPrimaryInv(ByVal TblDO_No As Core_Module.DseTrans_CutOffDB_Core.tbl_do_noDataTable, _
    ByVal vPrimaryQuickSalesInv As Double) As String
        Dim rowDO As Core_Module.DseTrans_CutOffDB_Core.tbl_do_noRow

        Try
            TblDO_No.DefaultView.RowFilter = Nothing
            TblDO_No.DefaultView.RowFilter = "[Primary Invoice] = " & vPrimaryQuickSalesInv
            If TblDO_No.DefaultView.Count > 0 Then
                rowDO = TblDO_No.DefaultView(0).Row
                Return rowDO.No_Transaksi
            Else
                Return "-"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNo_InOutByPrimaryInv(ByVal TblInOut_No As Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noDataTable, _
    ByVal vPrimaryInv As Double) As String
        Dim rowInOut As Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noRow
        Try
            TblInOut_No.DefaultView.RowFilter = Nothing
            TblInOut_No.DefaultView.RowFilter = "[Primary Invoice] = " & vPrimaryInv
            If TblInOut_No.DefaultView.Count > 0 Then
                rowInOut = TblInOut_No.DefaultView(0).Row
                Return rowInOut.No_Transaksi
            Else
                Return "-"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoInOut(ByVal vTblInOut As Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noDataTable, ByVal vPrimary As Double) As String
        Dim rowInOut As Core_Module.DseTrans_CutOffDB_Core.tbl_stockinout_noRow

        Try
            vTblInOut.DefaultView.RowFilter = "[Primary] = " & vPrimary
            If vTblInOut.DefaultView.Count > 0 Then
                rowInOut = vTblInOut.DefaultView(0).Row
                Return rowInOut.No_Transaksi
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Class
