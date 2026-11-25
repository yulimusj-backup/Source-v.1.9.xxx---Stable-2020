Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesReturn
    Dim ObjCD As New Core_Module.ClsCekData

    Public Enum EnumTipeSR As Integer
        Ganti_Barang
        Revisi_Invoice
        Nota_Retur_Pajak
    End Enum

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
       ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSalesReturn.tbltrans_salesreturn_view0DataTable, _
       ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
       Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
       Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter

        Dim QueryResult As Double
        Dim StartT As Date
        Dim TipeSR As EnumTipeSR = EnumTipeSR.Nota_Retur_Pajak

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari, TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                QueryResult = DTA.FillBy_SONumberLIKE(TBL, Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                QueryResult = DTA.FillBy_CustNameLIKE(TBL, Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInSalesReturn Then
                QueryResult = DTA.FillBy_NoStockInSalesReturnLIKE(TBL, Criteria + "%", TipeSR)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
                DGV.Columns(2).HeaderText = "Date"
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
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TipeSR As EnumTipeSR = EnumTipeSR.Nota_Retur_Pajak

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalSampaiDari(TipeSR, vTanggalSampai, vTanggalDari).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%", TipeSR).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerCode Then
                Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerName Then
                Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesOrderNo Then
                Return DTA.ScalarQuery_GetJmlByNoSOLIKE(Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDeliveryOrderNo Then
                Return DTA.ScalarQuery_GetJmlByDONumberLIKE(Criteria + "%", TipeSR)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInSalesReturn Then
                Return DTA.ScalarQuery_GetJmlByStockInSRLIKE(Criteria + "%", TipeSR)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustCode(ByVal vPrimarySR As Double) As String
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCustCode(vPrimarySR)) = False Then
                Return DTA.ScalarQuery_GetCustCode(vPrimarySR)
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

    Public Function GetCustName(ByVal vPrimarySR As Double) As String
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCustName(vPrimarySR)) = False Then
                Return DTA.ScalarQuery_GetCustName(vPrimarySR)
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

    Public Function ApakahSudahAda(ByVal vNoTransaksi As String, ByVal vTIpe As Integer) As Boolean
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter
        Try
            If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi, vTIpe).HasValue Then
                If DTA.ScalarQuery_GetJmlByNoTransaksi(vNoTransaksi, vTIpe).Value > 0 Then
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

    Public Function GetSRDate(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter
        Try
            If DTA.ScalarQuery_GetSRDate(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetSRDate(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimarySR As String) As Double
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimarySR).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimarySR).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vSRNumber As String) As Double
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter
        Try
            If DTA.ScalarQuery_GetPrimary(vSRNumber).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vSRNumber).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'mengembalikan no invoice jika diberikan no transaksi-nya
    'tidak berlaku untuk quick sales return
    Public Function GetInvoiceNumber(ByVal vNoSR As String, ByVal DTA2 As DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturnTableAdapter, _
    ByVal DTA1 As DseTransaksiDOTableAdapters.DataTable1TableAdapter, ByVal TBL1 As DseTransaksiDO.DataTable1DataTable, _
    ByVal ObjDO As ClsTransaksiDO) As String

        Dim PrimaryDO As Double = -1

        Try
            If DTA2.ScalarQuery_GetPrimaryDO_NonQuick(vNoSR).HasValue = True Then
                PrimaryDO = DTA2.ScalarQuery_GetPrimaryDO_NonQuick(vNoSR).Value
                Return ObjDO.GetNoInvoice(PrimaryDO, DTA1, TBL1)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn_view0TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSales_Return_Number(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetSales_Return_Number(vPrimary)
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

    'mengisikan data no invoice pada pada parameter no transaksi, tanggal pada parameter date, dan no sr + remarks pada parameter remarks
    Public Sub GetNoTransaksi_Remarks_NoInvoice(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByVal vPrimary As Double, ByRef vDate As Date)
        Dim DTA As New DseTransaksiSalesReturnTableAdapters.tbltrans_salesreturn1TableAdapter
        Dim TBL As New DseTransaksiSalesReturn.tbltrans_salesreturn1DataTable
        Dim row As DseTransaksiSalesReturn.tbltrans_salesreturn1Row

        Try
            DTA.Fill(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.No_Invoice <> "" Then
                    If row.Remarks <> "" Then
                        vRemarks = "SR Number: " + row.No_Transaksi + " - " + row.Remarks
                    Else
                        vRemarks = "SR Number: " + row.No_Transaksi
                    End If
                    vNoTransaksi = row.No_Invoice
                Else
                    vRemarks = row.Remarks
                    vNoTransaksi = row.No_Transaksi
                End If
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
