Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseReturn
    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumTipe_PurchaseReturn As Integer
        Revisi_AOP
        Nota_Retur
    End Enum

    Public Function IsiDGV(ByVal vCriteria As String, _
          ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPurchaseReturn.DataTable3DataTable, _
          ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, Optional ByVal vStartDate As Date = (#1/1/1950#), Optional ByVal vEndDate As Date = (#1/1/2100#)) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim TipePR As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Nota_Retur

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalDariSampai(TBL, vStartDate, vEndDate, TipePR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, vCriteria + "%", TipePR)
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                QueryResult = DTA.FillBy_SuppNameLIKE(TBL, vCriteria + "%", TipePR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                QueryResult = DTA.FillBy_SuppCodeLIKE(TBL, vCriteria + "%", TipePR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                QueryResult = DTA.FillBy_NoGRNLIKE(TBL, vCriteria + "%", TipePR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                QueryResult = DTA.FillBy_NoStockInPurchaseLIKE(TBL, vCriteria + "%", TipePR)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns("Tanggal Transaksi").DefaultCellStyle = ij
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
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable31TableAdapter
        Dim TipePR As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Nota_Retur

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.ScalarQuery_NoTransaksiLIKE(Criteria + "%", TipePR).HasValue = True Then
                    Return DTA.ScalarQuery_NoTransaksiLIKE(Criteria + "%", TipePR)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalDari, vTanggalSampai, TipePR).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalDari, vTanggalSampai, TipePR).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                If DTA.ScalarQuery_GetJmlBySuppCodeLIKE(Criteria + "%", TipePR).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlBySuppCodeLIKE(Criteria + "%", TipePR)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                If DTA.ScalarQuery_GetJmlBySuppNameLIKE(Criteria + "%", TipePR).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlBySuppNameLIKE(Criteria + "%", TipePR)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                If DTA.ScalarQuery_GetJmlByNoGRN(Criteria + "%", TipePR).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByNoGRN(Criteria + "%", TipePR)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                If DTA.ScalarQuery_NoStockInPurchaseLIKE(Criteria + "%", TipePR).HasValue = True Then
                    Return DTA.ScalarQuery_NoStockInPurchaseLIKE(Criteria + "%", TipePR)
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSuppCode(ByVal vPrimaryPR As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSuppCode(vPrimaryPR)) = False Then
                Return DTA.ScalarQuery_GetSuppCode(vPrimaryPR)
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

    Public Function GetSuppName(ByVal vPrimaryPR As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSuppName(vPrimaryPR)) = False Then
                Return DTA.ScalarQuery_GetSuppName(vPrimaryPR)
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
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).HasValue Then
                If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi).Value > 0 Then
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

    Public Function GetPRDate(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Try
            If DTA.ScalarQuery_GetPRDate(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPRDate(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryPR As String) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Try
            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryPR).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryPR).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vPRNumber As String) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Try
            If DTA.ScalarQuery_GetPrimary(vPRNumber).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vPRNumber).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'mengembalikan no transaksi purchase return yang merujuk aop yang diberikan primary-nya
    'tipe retur hanya yang tipenya nota return pajak
    Public Function GetNoTransaksi_NotaRetur_ByAOP(ByVal vPrimaryAOP As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.DataTable3DataTable
        Dim row As DseTransaksiPurchaseReturn.DataTable3Row
        Dim TipeRetur As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Nota_Retur
        Dim ListNoPR As New List(Of String)
        Dim RetVal As String = ""

        Try
            DTA.FillBy_PrimaryAOP(TBL, CInt(TipeRetur), vPrimaryAOP)
            For Each row In TBL.Rows
                If ListNoPR.Contains(row.No_Transaksi) = False Then
                    If RetVal = "" Then
                        RetVal = row.No_Transaksi
                    Else
                        RetVal = RetVal + row.No_Transaksi
                    End If
                    ListNoPR.Add(row.No_Transaksi)
                End If
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ListNoPR = Nothing
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPRNumber(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetPRNumber(vPrimary)
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

    Public Sub GetNoTransaksi_Remarks_NoInvoice(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByVal vPrimary As Double, ByRef vDate As Date, _
    ByRef vSuppInvNumber As String)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable1TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.DataTable1DataTable
        Dim row As DseTransaksiPurchaseReturn.DataTable1Row

        Try
            DTA.Fill(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vRemarks = row.Remarks
                vNoTransaksi = row.No_Transaksi
                vSuppInvNumber = row.Supplier_Invoice_No
                vDate = row.Tanggal_Transaksi
            Else
                vRemarks = ""
                vNoTransaksi = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub
End Class
