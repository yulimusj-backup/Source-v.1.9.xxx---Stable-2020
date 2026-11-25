Option Explicit On
Option Strict Off

Public Class FrmCekMovement
    Dim StartPeriod, EndPeriod As Date
    Dim ObjCD As New Core_Module.ClsCekData

    Private Sub CekOut(ByVal TBLMovement As DseTransaksiMonthlyClosing.tbltrans_puchaseclosing_item0DataTable, _
    ByVal TBLSales As Sales_Module.DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable, _
    ByVal TBLOutAdjustment As Inventory_Module.dsLaporanAdjustmentList.laporan_AdjustmentList_view0DataTable, _
    ByVal TBLPurchaseReturn As DsePurchaseReturnList.DataTable1DataTable)
        Dim rowMovement As DseTransaksiMonthlyClosing.tbltrans_puchaseclosing_item0Row
        Dim JmlSales, JmlPurchaseReturn, JmlOutAdjust As Double

        Dim TotalSales, TotalPR, TotalOutAdjust As Double

        Try
            Me.ProgressBar1.Value = 0
            Me.ProgressBar1.Maximum = TBLMovement.Rows.Count + 1

            For Each rowMovement In TBLMovement.Rows
                Me.ProgressBar1.Value += 1
                JmlSales = Me.GetCostSales(rowMovement.Product_Code, TBLSales)
                JmlOutAdjust = Me.GetCostAdjust(rowMovement.Primary_Product, TBLOutAdjustment)
                JmlPurchaseReturn = Me.GetCostPR(rowMovement.Primary_Product, TBLPurchaseReturn)

                If Math.Abs(rowMovement.Out_Amount - (JmlSales + JmlPurchaseReturn + JmlOutAdjust)) > 500 Then
                    MsgBox(rowMovement.Product_Code)
                End If

                TotalSales += JmlSales
                TotalPR += JmlPurchaseReturn
                TotalOutAdjust += JmlOutAdjust
            Next

            MsgBox("Sales Cost:" + Me.ObjCD.DigitGrouping(TotalSales) + "; Purchase Return: " + _
            Me.ObjCD.DigitGrouping(TotalPR) + "; Out Adjust: " + Me.ObjCD.DigitGrouping(TotalOutAdjust))
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetCostSales(ByVal vProductCode As String, _
    ByVal TBLSales As Sales_Module.DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable) As Double
        Dim rowSales As Sales_Module.DseLaporanSalesAnalysis.Group_By_Product_BrandRow
        Dim retVal As Double = 0

        Try
            For Each rowSales In TBLSales.Rows
                If rowSales.Product_Brand = vProductCode Then
                    retVal += rowSales.Cost_Of_Sales
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetCostAdjust(ByVal vPrimaryProduct As Double, _
ByVal TBLAdjust As Inventory_Module.dsLaporanAdjustmentList.laporan_AdjustmentList_view0DataTable) As Double
        Dim rowAdjust As Inventory_Module.dsLaporanAdjustmentList.laporan_AdjustmentList_view0Row
        Dim retVal As Double = 0

        Try
            For Each rowAdjust In TBLAdjust.Rows
                If rowAdjust.Primary_Product = vPrimaryProduct Then
                    retVal += rowAdjust.Cost_Amount
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetCostPR(ByVal vPrimaryProduct As Double, _
ByVal TBLPR As DsePurchaseReturnList.DataTable1DataTable) As Double
        Dim rowPR As DsePurchaseReturnList.DataTable1Row
        Dim retVal As Double = 0

        Try
            For Each rowPR In TBLPR.Rows
                If rowPR.Primary_Product = vPrimaryProduct Then
                    retVal += rowPR.Cost
                End If
            Next

            Return retVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Sub PersiapanDataOut()
        Dim FrmSales As New Sales_Module.FrmLaporanSalesAnalysis
        Dim TBLSales As Sales_Module.DseLaporanSalesAnalysis.Group_By_Product_BrandDataTable = FrmSales.TBLGroupByBrand
        Dim FrmAdjust As New Inventory_Module.frmLaporanAdjustmentList
        Dim TBLOutAdjust As Inventory_Module.dsLaporanAdjustmentList.laporan_AdjustmentList_view0DataTable
        Dim FrmPR As New FrmLaporanPurchaseReturnList
        Dim TBLPR As DsePurchaseReturnList.DataTable1DataTable
        Dim TBLMovement As New DseTransaksiMonthlyClosing.tbltrans_puchaseclosing_item0DataTable
        Dim DTAMovement As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_puchaseclosing_item0TableAdapter

        Try
            DTAMovement.Fill(TBLMovement, 22)
            MsgBox("Movements: " + TBLMovement.Rows.Count.ToString)
            Me.StartPeriod = #12/1/2010#
            Me.EndPeriod = #12/31/2010#
            'isikan data sales
            FrmSales.CekOut(Me.StartPeriod, Me.EndPeriod)
            MsgBox("Sales: " + TBLSales.Rows.Count.ToString)
            'isikan data out adjust
            'FrmAdjust.CekOut(Me.StartPeriod, Me.EndPeriod)
            'TBLOutAdjust = FrmAdjust.DSCekOut.laporan_AdjustmentList_view0
            'MsgBox("Adjust: " + TBLOutAdjust.Rows.Count.ToString)
            'isikan out purchase return
            FrmPR.CekOut1(Me.StartPeriod, Me.EndPeriod)
            TBLPR = FrmPR.DSOut.DataTable1
            MsgBox("PR: " + TBLPR.Rows.Count.ToString)
            Me.CekOut(TBLMovement, TBLSales, TBLOutAdjust, TBLPR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            'Dim Th As New Threading.Thread(AddressOf PersiapanDataOut)
            'Control.CheckForIllegalCrossThreadCalls = False
            'Th.Start()
            Me.PersiapanDataOut()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub


End Class