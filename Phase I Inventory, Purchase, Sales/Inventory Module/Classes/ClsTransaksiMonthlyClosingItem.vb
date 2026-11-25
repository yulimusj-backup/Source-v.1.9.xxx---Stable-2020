Option Explicit On
Option Strict Off

Public Class ClsTransaksiMonthlyClosingItem
    Public Enum EnumTipeDokumenAcuan As Integer
        StockInAdjust
        StockTakingInAdjustment
        StockBundlingAdjustment
        StockInSalesReturnPeriodeSama
    End Enum

    Public Enum EnumTipeItemDetail As Integer
        BeginingBalanceItem
        InItem
        OutItem
        EndingBalanceItem
    End Enum

    Public Enum EnumTipeTranskasiAsal As Integer
        InPurchase '0
        InSR '1
        InAdjust '2
        InStockTakeAdjust '3
        OutBundling '4
        OutSales '5
        OutPR '6
        OutStockTakeAdjust '7
        OutAdjust '8
        InBundling '9
        OverheadCost '10
    End Enum

    Public Function GetPrimaryClosing(ByVal vClosingPeriod As Date) As Double
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vClosingPeriod).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vClosingPeriod).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetInAmount_Purchase_SalesReturnBedaPeriode(ByVal vPrimaryCodeNo As Double, _
   ByVal vTanggalSampai As Date, ByVal vTanggalDari As Date, _
   ByVal DTA As DseTransaksiMonthlyClosingTableAdapters.DataTable25TableAdapter, _
   ByVal TBL As DseTransaksiMonthlyClosing.DataTable25DataTable, ByVal ObjCD As Core_Module.ClsCekData, _
    ByVal vObjCur As Core_Module.clsMasterCurrency) As Double
        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable4TableAdapter
        Dim TBL1 As New DseTransaksiMonthlyClosing.DataTable4DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable4Row
        Dim Dummy As Double = 0
        Dim TBLSR As New DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1DataTable
        Dim DTASR As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item9_1TableAdapter
        Dim rowSR As DseTransaksiMonthlyClosing.tbltrans_purchaseclosing_item9_1Row


        Try
            'in purchase
            DTA1.FillBy(TBL1, vTanggalSampai, vTanggalDari, vPrimaryCodeNo)
            If TBL1.Rows.Count > 0 Then
                row = TBL1.Rows(0)
                Dummy += row.Purchase_Amount
            End If
            'in sales return dari do yang beda periode sehingga bisa langsung diambil price at closingnya
            DTASR.Fill(TBLSR, vPrimaryCodeNo, vTanggalSampai, vTanggalDari)
            For Each rowSR In TBLSR.Rows
                If rowSR.DO_Date >= vTanggalDari And rowSR.DO_Date <= vTanggalSampai Then
                    rowSR.Unit_Price = 0
                ElseIf rowSR.DO_Date < vTanggalDari And rowSR.Unit_Price = 0 Then
                    rowSR.Unit_Price = Me.GetUnitPrice(vPrimaryCodeNo, ObjCD.GetFirstDateofTheMonth(rowSR.DO_Date), _
                    DTA, TBL, vObjCur.GetDecimal_DefaultCurrency)
                End If
                Dummy += (rowSR.Unit_Price * rowSR.Total_Qty)
            Next

            Return Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
        End Try
    End Function

    Private Function GetUnitPrice(ByVal vPrimaryProduct As Double, ByVal vClosingPeriode As Date, _
   ByVal DTA As DseTransaksiMonthlyClosingTableAdapters.DataTable25TableAdapter, _
   ByVal TBL As DseTransaksiMonthlyClosing.DataTable25DataTable, ByVal DefaultPrec As Integer) As Decimal
        Dim rowClosing As DseTransaksiMonthlyClosing.DataTable25Row

        Try
            DTA.Fill(TBL, vClosingPeriode, vPrimaryProduct)

            If TBL.Rows.Count = 0 Then
                Return 0
            Else
                rowClosing = TBL.Rows(0)
                If rowClosing.Ending_Qty = 0 Then
                    If rowClosing.Out_Qty = 0 Then
                        Return 0
                    Else
                        Return Math.Round(rowClosing.Out_Amount / rowClosing.Out_Qty, DefaultPrec, MidpointRounding.AwayFromZero)
                    End If
                Else
                    Return Math.Round(rowClosing.Ending_Amount / rowClosing.Ending_Qty, DefaultPrec, MidpointRounding.AwayFromZero)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Sub Get_BeginAmount_BeginQty_InQty_IAmount(ByVal vPrimaryProduct As Double, ByRef vBeginAmount As Double, _
       ByRef vBeginQty As Double, ByRef vInQty As Double, ByVal vTanggalSampai As Date, _
       ByVal vTanggalDari As Date, ByVal ObjCD As Core_Module.ClsCekData)
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchaseclosing_item2TableAdapter
        Dim DTA1 As New DseTransaksiMonthlyClosingTableAdapters.DataTable16TableAdapter
        Dim TBL1 As New DseTransaksiMonthlyClosing.DataTable16DataTable
        Dim row1 As DseTransaksiMonthlyClosing.DataTable16Row

        Try
            If DTA.ScalarQuery(vPrimaryProduct, vTanggalSampai, vTanggalDari).HasValue = True Then
                vInQty = DTA.ScalarQuery(vPrimaryProduct, vTanggalSampai, vTanggalDari).Value
            Else
                vInQty = 0
            End If

            'mencari data begin amount n begin qty dari bulan sebelumnya
            vTanggalSampai = vTanggalDari - TimeSpan.FromDays(1)
            vTanggalDari = ObjCD.GetFirstDateofTheMonth(vTanggalSampai)
            If DTA1.FillBy_CLosingPeriod(TBL1, vTanggalDari, vPrimaryProduct) > 0 Then
                row1 = TBL1.Rows(0)
                vBeginQty = row1.Ending_Qty
                vBeginAmount = row1.Ending_Amount
            Else
                vBeginQty = 0
                vBeginAmount = 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Public Function GetLastClosingPeriod() As Date
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiMonthlyClosing.DataTable2DataTable
        Dim row As DseTransaksiMonthlyClosing.DataTable2Row

        Try
            If DTA.Fill(TBL) > 0 Then
                row = TBL.Rows(0)
                If row.IsMaxClosingPeriodNull = True Then
                    Return #1/1/1950#
                End If
                Return row.MaxClosingPeriod
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetJml() As Double
        Dim DTA As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetJml

            If RetVal.HasValue = True Then
                Return RetVal.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
