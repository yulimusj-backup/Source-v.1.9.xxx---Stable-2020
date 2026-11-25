Option Explicit On
Option Strict Off

Public Class ClsTrans_ManufacturingOrder_Item
    Public Sub IsiDataItemWO(ByVal vIdentifierJO As String, ByVal vTblitemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable)
        Dim DtaItemWO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter
        Try
            DtaItemWO.Connection.ConnectionString = Core_Module.GetConstring
            DtaItemWO.FillBy_IdentiferJO(vTblitemWO, vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetBalanceMP_SimpanPerubahan(ByVal vPrimaryMO As Double, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable)
        Dim DTAItemWO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim TBLItemWO As New DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable
        Dim rowItemWO, rowItemWO1 As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim QtyUsed, QtyBalance, a As Decimal
        Dim vRowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim DTA1 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling_item_before6TableAdapter
        Dim TBL1 As New DseTrans_StockBundling.tbltransaksi_stockbundling_item_before6DataTable
        Dim row1 As DseTrans_StockBundling.tbltransaksi_stockbundling_item_before6Row
        Dim DTA2 As New DseTrans_StockBundlingTableAdapters.tblgets_rmTableAdapter
        Dim GP_Actual, QtyRefProcess, QtyRM, QtyActualProduce, QtyUsedWIP, QtyBalanceWIP As Decimal
        Dim DTA4 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_bomTableAdapter
        Dim TBL4 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable
        Dim row4 As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim DTA5 As New DseTrans_StockBundlingTableAdapters.tbltransaksi_stockbundling7TableAdapter
        Dim TBL5 As New DseTrans_StockBundling.tbltransaksi_stockbundling7DataTable
        Dim row5 As DseTrans_StockBundling.tbltransaksi_stockbundling7Row
        Dim SudahAdaSchedule As Boolean = False
        Dim vRowItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipRow

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA4.Connection.ConnectionString = Core_Module.GetConstring
            DTA5.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemWO.Connection.ConnectionString = Core_Module.GetConstring

            '1. item wo
            'Status, Actual_StartTime, Actual_EndTime, Actual_SetupTime, Actual_WorkHour, Actual_TotalTime, _
            'Actual_Downtime, Actual_ProdRate, Actual_Produce, Actual_Result, Actual_ProjectResult, Actual_GoodPercentage, _
            'Balance_Produce, Balance_Result, PrimaryWU, Balance_MP
            'berdasarkan kombinasi identifier jo dan identifier wo
            DTAItemWO.FillBy_PrimaryMO(TBLItemWO, vPrimaryMO)
            For Each rowItemWO1 In vTblItemWO.Rows
                For Each rowItemWO In TBLItemWO.Rows
                    If rowItemWO.Identifier_JO = rowItemWO1.Identifier_JO And _
                    rowItemWO.Identifier = rowItemWO1.Identifier Then
                        With rowItemWO1
                            .Status = rowItemWO.Status
                            .Actual_Start_Time = rowItemWO.Actual_Start_Time
                            .Actual_End_Time = rowItemWO.Actual_End_Time
                            .Actual_Setup_Time = rowItemWO.Actual_Setup_Time
                            .Actual_Work_Hour = rowItemWO.Actual_Work_Hour
                            .Actual_Total_Time = rowItemWO.Actual_Total_Time
                            .Actual_Down_Time = rowItemWO.Actual_Down_Time
                            .Actual_Prod_Rate = rowItemWO.Actual_Prod_Rate
                            .Qty_Actual_Prod = rowItemWO.Qty_Actual_Prod
                            .Qty_Actual_Result = rowItemWO.Qty_Actual_Result
                            .Qty_Projected_Result = rowItemWO.Qty_Projected_Result
                            .Actual_Good_Perct = rowItemWO.Actual_Good_Perct
                            .Balance_Produce = rowItemWO.Balance_Produce
                            .Balance_Result = rowItemWO.Balance_Result
                            .Primary_Work_Unit = rowItemWO.Primary_Work_Unit
                            If SudahAdaSchedule = True Then
                                'qty balance mp, tidak akan berubah karena perubahan qty wip, pasti sama dengan balance mp lama
                                .Qty_Balance_MP = rowItemWO.Qty_Balance_MP
                            Else
                                'jika belum ada schedule, maka belum ada bundling, maka ada kemungkinan balance MP berubah
                                'karena user mengubah qty wip, 
                                'tapi karena kondisi belum ada schedule dan belum ada bundling maka bisa dipastikan qty balance mp = qty target result
                                .Qty_Balance_MP = .Qty_Target_Result
                            End If
                        End With

                        Exit For
                    End If
                Next
            Next

            'untuk setiap baris pada db, dari data yang bersesuaian, update nilainya
            'hitung ulang kembali semuanya


            '2. item rm wo: qty used, qty balance, gp actual
            'kumpulkan item rm untuk kombinasi identifier jo + wo
            DTA1.Fill(TBL1, vPrimaryMO)
            DTA4.FillBy_PrimaryMO(TBL4, vPrimaryMO)
            DTA5.Fill(TBL5, vPrimaryMO)

            'kumpulkan item before untuk kombinasi identifier jo + wo
            For Each vRowItemRM In vTblItemRM.Rows
                'untuk setiap baris item rm
                'cari qty used
                QtyUsed = 0

                For Each row1 In TBL1.Rows
                    If row1.Primary_Product = vRowItemRM.Primary_Product_Code_RM And _
                    row1.identifier_jo = vRowItemRM.Identifier_JO And _
                    row1.identifier_wo = vRowItemRM.Identifier_WO Then
                        QtyUsed += row1.Qty
                        Exit For
                    End If
                Next

                'set nilai balance
                If vRowItemRM.is_wip = 1 Then
                    'set nilai balance-> khusus untuk item wip, asumsi menggunakan wip terlebih dahulu, sehingga  qty balance 
                    'rumusnya berbeda jika qty used < qty wip vs qty used => qty wip
                    If vRowItemRM.WIP_Qty - QtyUsed < 0 Then
                        'ada alocated rm yang digunakan
                        QtyBalance = vRowItemRM.Target_Qty + vRowItemRM.Additional_Qty - vRowItemRM.WIP_Qty + (vRowItemRM.WIP_Qty - QtyUsed)
                        'semua item wip sudah digunakan
                        QtyUsedWIP = vRowItemRM.WIP_Qty
                        QtyBalanceWIP = 0
                    Else
                        'allocated rm belum digunakan
                        QtyBalance = vRowItemRM.Target_Qty + vRowItemRM.Additional_Qty - vRowItemRM.WIP_Qty
                        'wip baru digunakan sebagian
                        QtyUsedWIP = QtyUsed
                        QtyBalanceWIP = vRowItemRM.WIP_Qty - QtyUsedWIP
                    End If
                    'update nilai qty used dan qty balance pada item wip
                    If vRowItemRM.WIP_Qty > 0 Then
                        For Each vRowItemWIP In vTblItemWIP.Rows
                            If vRowItemWIP.Primary_Product_Code_WIP = vRowItemRM.Primary_Product_Code_RM And _
                            vRowItemRM.Identifier_JO = vRowItemWIP.Identifier_JO And _
                            vRowItemRM.Identifier_WO = vRowItemWIP.Identifier_WO Then
                                vRowItemWIP.Qty_Balance = QtyBalanceWIP
                                vRowItemWIP.Qty_Used = QtyUsedWIP
                            End If
                        Next
                    End If
                Else
                    QtyBalance = vRowItemRM.Target_Qty + vRowItemRM.Additional_Qty - vRowItemRM.Replace_Qty - QtyUsed
                End If

                'set nilai gp aktual
                For Each row4 In TBL4.Rows
                    If row4.Primary_Product_RM = vRowItemRM.Primary_Product_Code_RM And _
                    row4.Identifier_JO = vRowItemRM.Identifier_JO And _
                    row4.Identifier_WO = vRowItemRM.Identifier_WO Then
                        QtyRefProcess = row4.Qty_Ref_Process
                        QtyRM = row4.Qty_RM
                        Exit For
                    End If
                Next

                'get nilai qty actual produce
                QtyActualProduce = 0
                For Each row5 In TBL5.Rows
                    If row5.identifier_jo = vRowItemRM.Identifier_JO And _
                    row5.identifier_wo = vRowItemRM.Identifier_WO Then
                        QtyActualProduce += row5.total_produce
                        Exit For
                    End If
                Next

                If QtyUsed = 0 Then
                    GP_Actual = 0
                Else
                    If QtyRefProcess > 0 Then a = QtyRM / QtyRefProcess Else a = 0
                    GP_Actual = ObjCD.DigitGrouping(((a * QtyActualProduce) / QtyUsed) * 100, 2)
                End If

                vRowItemRM.Used_Qty = QtyUsed
                vRowItemRM.Balance_Qty = QtyBalance
                vRowItemRM.GP_Actual = GP_Actual
            Next

            '3. rm detail + wip detail
            Me.Update_RM_WIP_Detail(vPrimaryMO, vTblItemWO, vTblItemRM, vTblItemWIP)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Update_RM_WIP_Detail(ByVal vPrimaryMO As Double, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable)
        Dim DTARMItemDetail As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rm_detailTableAdapter
        Dim TBLRMDetail As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable
        Dim DTAWIPItemDetail As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wip_detailTableAdapter
        Dim TBLWIPDetail As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable
        Dim ListIdentifierJO As New List(Of String)
        Dim RowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim ObjMO As New ClsTrans_ManufacturingOrder
        Dim TBLUpdate As New DseTrans_ManufacturingOrder.DataTable6DataTable


        Try
            'set connection string dahulu
            DTARMItemDetail.Connection.ConnectionString = Core_Module.GetConstring
            DTAWIPItemDetail.Connection.ConnectionString = Core_Module.GetConstring

            'load item rm detail + wip detail
            DTARMItemDetail.FillBy_PrimaryMO(TBLRMDetail, vPrimaryMO)
            DTAWIPItemDetail.FillBy_PrimaryMO(TBLWIPDetail, vPrimaryMO)
            TBLUpdate.Rows.Clear()

            'kumpulan identifier jo
            ListIdentifierJO.Clear()
            For Each RowItemWO In vTblItemWO.Rows
                If ListIdentifierJO.Contains(RowItemWO.Identifier_JO) = False Then
                    ListIdentifierJO.Add(RowItemWO.Identifier_JO)
                End If
            Next

            For Each IdentifierJO As String In ListIdentifierJO
                TBLUpdate.Rows.Clear()

                If ObjMO.IsScheduled(vPrimaryMO, IdentifierJO) = True Then
                    'untuk setiap identifier jo yang sudah ada jadwal, mulai perhitungan update qty rm + wip detail
                    '1. untuk setiap wo pada jo ybs, cari kombinasi jo + wo, yang qty targetnya berbeda antara rm + wip detail vs rm + rm add + wip allocated
                    Me.GetKombinasiTidakCocok(TBLUpdate, IdentifierJO, vTblItemWO, vTblItemRM, TBLRMDetail, TBLWIPDetail)

                    '2. untuk setiap kombinasi yang tidak sesuai, update perhitungannya
                    Me.UpdatePerhitungan(TBLUpdate, IdentifierJO, vTblItemWO, vTblItemRM, TBLRMDetail, TBLWIPDetail, _
                    vPrimaryMO, DTARMItemDetail, DTAWIPItemDetail)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub UpdatePerhitungan(ByVal vTBLUpdate As DseTrans_ManufacturingOrder.DataTable6DataTable, _
    ByVal vIdentifierJO As String, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vTBLRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable, _
    ByVal vTBLWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable, ByVal vPrimaryMO As Double, _
    ByVal vDTARMItemDetail As DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rm_detailTableAdapter, _
    ByVal vDTAWIPItemDetail As DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wip_detailTableAdapter)
        Dim DTAWU As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wo_item_sch1TableAdapter
        Dim TBLWU As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo_item_sch1DataTable
        Dim rowWU As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo_item_sch1Row
        Dim JmlDetail As Decimal = 0
        Dim AllocatedQty As Decimal = 0
        Dim TotalQtyProduce As Decimal = 0
        Dim DefPrec As Integer = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct
        DTAWU.Connection.ConnectionString = Core_Module.GetConstring
        Dim rowUpdate As DseTrans_ManufacturingOrder.DataTable6Row
        Dim rowRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailRow
        Dim rowWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailRow
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim RowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow

        Try
            If vTBLUpdate.Rows.Count > 0 Then
                'a. kumpulkan primary wu yang dialokasikan
                If TBLWU.Rows.Count = 0 Then
                    DTAWU.Fill(TBLWU, vPrimaryMO)
                End If
                'b. buang item rm detail yang tidak ada di item rm
                For Each rowUpdate In vTBLUpdate.Rows
                    For ij As Double = vTBLRMDetail.Rows.Count - 1 To 0 Step -1
                        rowRMDetail = vTBLRMDetail.Rows(ij)
                        If rowRMDetail.Identifier_JO = vIdentifierJO And rowRMDetail.Identifier_WO = rowUpdate.Identifier_WO Then
                            If Me.GetCountRM(rowRMDetail.Primary_Product_Code_RM, rowRMDetail.Identifier_JO, rowRMDetail.Identifier_WO, vTblItemRM) = 0 Then
                                vDTARMItemDetail.DeleteQuery1(rowRMDetail.Primary)
                                vTBLRMDetail.Rows.Remove(rowRMDetail)
                            End If
                        End If
                    Next
                Next
                'c. tambahkan item rm bila tidak ada di rm detail dan update qty target + qty balance pada rm detail
                For Each rowUpdate In vTBLUpdate.Rows
                    For ij As Double = vTblItemRM.Rows.Count - 1 To 0 Step -1
                        RowItemRM = vTblItemRM.Rows(ij)

                        If RowItemRM.Identifier_JO = rowUpdate.Identifier_JO And RowItemRM.Identifier_WO = rowUpdate.Identifier_WO Then
                            JmlDetail = Me.GetQtyRMDetail(RowItemRM.Primary_Product_Code_RM, RowItemRM.Identifier_JO, RowItemRM.Identifier_WO, vTBLRMDetail)
                            If JmlDetail <> RowItemRM.Required_Qty + RowItemRM.Additional_Qty - RowItemRM.Replace_Qty Then
                                AllocatedQty = 0
                                TotalQtyProduce = 0
                                DefPrec = ObjProduct.GetAngkaDecimal(RowItemRM.Primary_Product_Code_RM)

                                TBLWU.DefaultView.RowFilter = "[identifier jo] = '" + rowUpdate.Identifier_JO + "' AND [identifier wo] = '" + rowUpdate.Identifier_WO + "'"
                                For Each DV As Data.DataRowView In TBLWU.DefaultView
                                    rowWU = DV.Row
                                    TotalQtyProduce += rowWU.qty_produce
                                Next
                                If JmlDetail = 0 Then
                                    'tambahkan baris di item detail + simpan langsung ke database
                                    For mn As Integer = 0 To TBLWU.DefaultView.Count - 1
                                        rowWU = TBLWU.DefaultView(mn).Row
                                        rowRMDetail = vTBLRMDetail.Newtbltrans_mo_item_rm_detailRow
                                        With rowRMDetail
                                            .Identifier_JO = rowUpdate.Identifier_JO
                                            .Identifier_WO = rowUpdate.Identifier_WO
                                            .is_wip = 0
                                            .Primary = -1
                                            .Primary_Input_Location = rowWU.primary_input_location
                                            .Primary_MO = vPrimaryMO
                                            .Primary_Product_Code_RM = RowItemRM.Primary_Product_Code_RM
                                            .Primary_WU = rowWU.primary_work_unit
                                            .Remarks = ""
                                            If mn = TBLWU.DefaultView.Count - 1 Then
                                                .Target_Qty = RowItemRM.Required_Qty - RowItemRM.Additional_Qty - RowItemRM.Replace_Qty - AllocatedQty
                                            Else
                                                .Target_Qty = Math.Round(rowWU.qty_produce * _
                                                (RowItemRM.Required_Qty - RowItemRM.Additional_Qty - RowItemRM.Replace_Qty) / TotalQtyProduce, DefPrec, _
                                                MidpointRounding.AwayFromZero)
                                                AllocatedQty += .Target_Qty
                                            End If

                                            .Used_Qty = 0
                                            .Balance_Qty = .Target_Qty
                                            vTBLRMDetail.Rows.Add(rowRMDetail)
                                            'insert ke dalam db
                                            vDTARMItemDetail.InsertQuery(.Primary_MO, .Identifier_JO, .Identifier_WO, .Primary_WU, _
                                            .Primary_Product_Code_RM, .Target_Qty, .Used_Qty, .Balance_Qty, .Remarks, .Primary_Input_Location, .is_wip)
                                        End With
                                    Next
                                Else
                                    'sesuaikan qty target dan qty balance berdasarkan weighing average
                                    For mn As Integer = 0 To TBLWU.DefaultView.Count - 1
                                        rowWU = TBLWU.DefaultView(mn).Row
                                        For Each rowRMDetail In vTBLRMDetail.Rows
                                            If rowRMDetail.Identifier_WO = rowUpdate.Identifier_WO And rowRMDetail.Identifier_JO = rowUpdate.Identifier_JO _
                                            And rowRMDetail.Primary_Product_Code_RM = RowItemRM.Primary_Product_Code_RM And _
                                            rowRMDetail.Primary_WU = rowWU.primary_work_unit Then
                                                With rowRMDetail
                                                    If mn = TBLWU.DefaultView.Count - 1 Then
                                                        .Target_Qty = RowItemRM.Required_Qty - RowItemRM.Additional_Qty - RowItemRM.Replace_Qty - AllocatedQty
                                                    Else
                                                        .Target_Qty = Math.Round(rowWU.qty_produce * _
                                                        (RowItemRM.Required_Qty - RowItemRM.Additional_Qty - RowItemRM.Replace_Qty) / TotalQtyProduce, DefPrec, _
                                                        MidpointRounding.AwayFromZero)
                                                        AllocatedQty += .Target_Qty
                                                    End If
                                                    .Balance_Qty = .Target_Qty - .Used_Qty
                                                    If .Balance_Qty < 0 Then
                                                        .Balance_Qty = 0
                                                    End If
                                                    'd. simpan perubahan ke dalam db untuk: hapus, simpan baru, update qty target + qty balance
                                                    vDTARMItemDetail.UpdateQuery(.Target_Qty, .Balance_Qty, .Primary)
                                                End With
                                                Exit For
                                            End If
                                        Next
                                    Next
                                End If
                            End If
                        End If
                    Next
                Next
                'e. ulang untuk qty wip vs qty wip detail
                For Each rowUpdate In vTBLUpdate.Rows
                    For Each RowItemWO In vTblItemWO.Rows
                        If RowItemWO.Identifier = rowUpdate.Identifier_WO And RowItemWO.Identifier_JO = rowUpdate.Identifier_JO Then
                            JmlDetail = Me.GetQtyWIPDetail(RowItemWO.Primary_Product, RowItemWO.Identifier_JO, RowItemWO.Identifier, vTBLWIPDetail)
                            If JmlDetail <> RowItemWO.Qty_WIP Then
                                TBLWU.DefaultView.RowFilter = "[identifier jo] = '" + rowUpdate.Identifier_JO + "' AND [identifier wo] = '" + rowUpdate.Identifier_WO + "'"
                                TotalQtyProduce = 0
                                AllocatedQty = 0
                                DefPrec = ObjProduct.GetAngkaDecimal(RowItemWO.Primary_Product)

                                For Each DV As Data.DataRowView In TBLWU.DefaultView
                                    rowWU = DV.Row
                                    TotalQtyProduce += rowWU.qty_produce
                                Next
                                If JmlDetail = 0 Then
                                    'tambahkan ke db
                                    For mn As Integer = 0 To TBLWU.DefaultView.Count - 1
                                        rowWU = TBLWU.DefaultView(mn).Row
                                        rowWIPDetail = vTBLWIPDetail.Newtbltrans_mo_item_wip_detailRow
                                        With rowWIPDetail
                                            .Identifier_JO = rowUpdate.Identifier_JO
                                            .Identifier_WO = rowUpdate.Identifier_WO
                                            .Primary = -1
                                            .Primary_Input_Location = rowWU.primary_input_location
                                            .Primary_MO = vPrimaryMO
                                            .Primary_Product_Code_WIP = RowItemWO.Primary_Product
                                            .Primary_WU = rowWU.primary_work_unit
                                            .Qty = 0
                                            If mn = TBLWU.DefaultView.Count - 1 Then
                                                .Qty = RowItemWO.Qty_WIP - AllocatedQty
                                            Else
                                                .Qty = Math.Round((rowWU.qty_produce * RowItemWO.Qty_WIP) / TotalQtyProduce, DefPrec, _
                                                MidpointRounding.AwayFromZero)
                                                AllocatedQty += .Qty
                                            End If
                                            .Used_Qty = 0
                                            .Balance_Qty = .Qty
                                            vTBLWIPDetail.Rows.Add(rowWIPDetail)
                                            vDTAWIPItemDetail.InsertQuery(.Primary_MO, .Primary_WU, .Identifier_JO, .Identifier_WO, .Primary_Product_Code_WIP, _
                                            .Qty, .Used_Qty, .Balance_Qty, .Primary_Input_Location)
                                        End With
                                    Next
                                Else
                                    'edit qty dan balance qty
                                    For mn As Integer = 0 To TBLWU.DefaultView.Count - 1
                                        rowWU = TBLWU.DefaultView(mn).Row
                                        For Each rowWIPDetail In vTBLWIPDetail.Rows
                                            If rowWIPDetail.Identifier_WO = rowUpdate.Identifier_WO And rowWIPDetail.Identifier_JO _
                                            And rowWIPDetail.Primary_Product_Code_WIP = RowItemWO.Primary_Product And _
                                            rowWIPDetail.Primary_WU = rowWU.primary_work_unit Then
                                                With rowWIPDetail
                                                    If mn = TBLWU.DefaultView.Count - 1 Then
                                                        .Qty = RowItemWO.Qty_WIP - AllocatedQty
                                                    Else
                                                        .Qty = Math.Round((rowWU.qty_produce * RowItemWO.Qty_WIP) / TotalQtyProduce, DefPrec, _
                                                        MidpointRounding.AwayFromZero)
                                                        AllocatedQty += .Qty
                                                    End If
                                                    .Balance_Qty = .Qty - .Used_Qty
                                                    If .Balance_Qty < 0 Then
                                                        .Balance_Qty = 0
                                                    End If
                                                    vDTAWIPItemDetail.UpdateQuery(.Qty, .Balance_Qty, .Primary)
                                                End With
                                            End If
                                        Next
                                    Next
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetKombinasiTidakCocok(ByVal vTBLUpdate As DseTrans_ManufacturingOrder.DataTable6DataTable, _
    ByVal vIdentifierJO As String, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vTBLRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable, _
    ByVal vTBLWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable)
        Dim RowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailRow
        Dim rowRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailRow
        Dim Cocok As Boolean = True
        Dim rowUpdate As DseTrans_ManufacturingOrder.DataTable6Row

        Try
            For Each RowItemWO In vTblItemWO.Rows
                If RowItemWO.Identifier_JO = vIdentifierJO Then
                    Cocok = True

                    'pastikan semua item di rm detail dan wip detail ada di tbl rm dan wip allocated
                    For Each rowRMDetail In vTBLRMDetail.Rows
                        If rowRMDetail.Identifier_JO = vIdentifierJO And rowRMDetail.Identifier_WO = RowItemWO.Identifier Then
                            If Me.GetCountRM(rowRMDetail.Primary_Product_Code_RM, rowRMDetail.Identifier_JO, rowRMDetail.Identifier_WO, vTblItemRM) = 0 Then
                                Cocok = False
                                Exit For
                            End If
                        End If
                    Next
                    For Each rowWIPDetail In vTBLWIPDetail.Rows
                        If rowWIPDetail.Identifier_JO = vIdentifierJO And rowWIPDetail.Identifier_WO = RowItemWO.Identifier Then
                            If RowItemWO.Qty_WIP = 0 Then
                                Cocok = False
                                Exit For
                            End If
                        End If
                    Next
                    'cocokan qty rm vs rm detail
                    If Cocok = True Then
                        For Each RowItemRM In vTblItemRM.Rows
                            If RowItemRM.Identifier_WO = RowItemWO.Identifier And RowItemRM.Identifier_JO = vIdentifierJO Then
                                If RowItemRM.Required_Qty + RowItemRM.Additional_Qty - RowItemRM.Replace_Qty <> _
                                Me.GetQtyRMDetail(RowItemRM.Primary_Product_Code_RM, RowItemRM.Identifier_JO, RowItemRM.Identifier_WO, vTBLRMDetail) Then
                                    Cocok = False
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                    'cocokan qty wip vs wip detail
                    If Cocok = True Then
                        If Me.GetQtyWIPDetail(RowItemWO.Primary_Product, vIdentifierJO, RowItemWO.Identifier, vTBLWIPDetail) <> RowItemWO.Qty_WIP Then
                            Cocok = False
                        End If
                    End If
                    'jika tidak cocok, catat kombinasinya
                    If Cocok = False Then
                        rowUpdate = vTBLUpdate.NewDataTable6Row
                        rowUpdate.Identifier_JO = vIdentifierJO
                        rowUpdate.Identifier_WO = RowItemWO.Identifier
                        vTBLUpdate.Rows.Add(rowUpdate)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetCountRM(ByVal vPrimaryRM As Double, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable) As Integer
        Dim RowItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            For Each RowItemRM In vTblItemRM.Rows
                If RowItemRM.Primary_Product_Code_RM = vPrimaryRM And RowItemRM.Identifier_JO = vIdentifierJO And _
                RowItemRM.Identifier_WO = vIdentifierWO Then
                    Return 1
                End If
            Next
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Private Function GetQtyRMDetail(ByVal vPrimaryRM As Double, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, _
    ByVal vTBLRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailDataTable) As Decimal
        Dim rowRMDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_detailRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowRMDetail In vTBLRMDetail.Rows
                If rowRMDetail.Identifier_JO = vIdentifierJO And rowRMDetail.Identifier_WO = vIdentifierWO And _
                rowRMDetail.Primary_Product_Code_RM = vPrimaryRM Then
                    RetVal += rowRMDetail.Target_Qty
                End If
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetQtyWIPDetail(ByVal vPrimaryWIP As Double, ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, _
    ByVal vTBLWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailDataTable) As Decimal
        Dim rowWIPDetail As DseTrans_ManufacturingOrder.tbltrans_mo_item_wip_detailRow
        Dim RetVal As Decimal = 0

        Try
            For Each rowWIPDetail In vTBLWIPDetail.Rows
                If rowWIPDetail.Identifier_JO = vIdentifierJO And rowWIPDetail.Identifier_WO = vIdentifierWO And _
                rowWIPDetail.Primary_Product_Code_WIP = vPrimaryWIP Then
                    RetVal += rowWIPDetail.Qty
                End If
            Next
            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetBalanceMP_SimpanBaru(ByVal vIdentifierJO As String)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetQtyBalanceMP(vIdentifierJO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'mengembalikan qty balance result untuk perhitungan mp
    'qty balance = qty balance item jo yang prosesnya adalah proses terakhir
    Public Sub SetMOBalance(ByVal vIdentifierJO As String)
        Dim IdentifierWO As String = ""
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wo4TableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo4DataTable
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo4Row
        Dim dta1 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            dta1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'ambil qty balance mp pada item wo yang merupakan proses terakhir yang menghasilkan mo
            DTA.Fill(TBL, vIdentifierJO)
            If TBL.Rows.Count = 1 Then
                'update nilai balancenya pada item mo
                row = TBL.Rows(0)
                dta1.UpdateQuery_SetQtyBalance(row.Qty_Balance_MP, vIdentifierJO)
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function ApakahAdaYg_Ready_BelumTerjadwal() As Boolean
        Dim DTA1 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wo2TableAdapter
        Dim TBL1 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo2DataTable

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Fill(TBL1, 0)
            If TBL1.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Function

    Public Sub IsiDGVJadwalBaru(ByVal vDGV As DataGridView, ByVal i As FrmTrans_ProductionSchedule_NewSch)
        Dim DTA1 As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wo2TableAdapter
        Dim TBL1 As New DseTrans_ManufacturingOrder.tbltrans_mo_item_wo2DataTable
        Dim row1 As DseTrans_ManufacturingOrder.tbltrans_mo_item_wo2Row
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item2TableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltrans_mo_item2DataTable
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item2Row
        Dim PrimaryMax, PrimaryMin As Double
        Dim rowDGV As DataGridViewRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim filter As String = ""
        Dim primaryMO As Double = -1

        Try


            vDGV.Rows.Clear()
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            'cari jo baru yang butuh diberesin
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Connection.ConnectionString = DTA1.Connection.ConnectionString

            DTA1.Fill(TBL1, 0)
            If TBL1.Rows.Count > 0 Then
                PrimaryMax = -1
                PrimaryMin = Double.MaxValue
                For Each row1 In TBL1.Rows
                    If row1.Primary_MO > PrimaryMax Then
                        PrimaryMax = row1.Primary_MO
                    End If
                    If row1.Primary_MO < PrimaryMin Then
                        PrimaryMin = row1.Primary_MO
                    End If
                Next

                'TR//Tambahan Filter
                i.TxtFilter_CodeNo.Text = i.TxtFilter_CodeNo.Text.Trim
                If i.TxtFilter_CodeNo.Text <> "" Then
                    If filter = "" Then
                        filter = "[Product Code] LIKE '" + ObjCD.FormatSpecialChar_DataFilter(i.TxtFilter_CodeNo.Text, True) + "*'"
                    Else
                        filter &= " AND [Product Code] LIKE '" + ObjCD.FormatSpecialChar_DataFilter(i.TxtFilter_CodeNo.Text, True) + "*'"
                    End If

                End If
                i.TxtFilter_Desc.Text = i.TxtFilter_Desc.Text.Trim
                If i.TxtFilter_Desc.Text <> "" Then
                    If filter = "" Then
                        filter = "[Description] like '" + ObjCD.FormatSpecialChar_DataFilter(i.TxtFilter_Desc.Text, True) + "*'"
                    Else
                        filter &= " AND [Description] like '" + ObjCD.FormatSpecialChar_DataFilter(i.TxtFilter_Desc.Text, True) + "*'"
                    End If
                End If

                If i.TxtFilter_PrimaryMO.Text <> "" And i.TxtFilter_PrimaryMO.Text <> -1 Then
                    If filter = "" Then
                        filter = "[Primary MO] = " & i.TxtFilter_PrimaryMO.Text
                    Else
                        filter &= " AND [Primary MO] = " & i.TxtFilter_PrimaryMO.Text
                    End If

                End If
                If filter <> "" Then
                    filter = " AND " & filter
                End If

                DTA.FillBy_PrimaryMOMaxMin(TBL, PrimaryMin, PrimaryMax)

                For Each row1 In TBL1.Rows
                    'TR//Tambahin filter di bawah ini
                    TBL.DefaultView.RowFilter = "[Identifier JO] = '" + row1.Identifier_JO + "' " & filter
                    If TBL.DefaultView.Count > 0 Then
                        row = TBL.DefaultView(0).Row
                        vDGV.Rows.Add()
                        rowDGV = vDGV.Rows(vDGV.Rows.Count - 1)
                        With rowDGV
                            .Cells(i.ColCodeNo.Index).Value = row.Product_Code
                            .Cells(i.ColIdentifierJO.Index).Value = row.Identifier_JO
                            .Cells(i.ColDesc.Index).Value = row.Description
                            .Cells(i.ColJO.Index).Value = row.JO_Number
                            .Cells(i.ColMO.Index).Value = row.MO_Number
                            .Cells(i.ColPrimaryMO.Index).Value = row.Primary_MO
                            .Cells(i.ColPriority.Index).Value = rowDGV.Index + 1
                            .Cells(i.ColQtyProduce.Index).Value = ObjCD.DigitGrouping(row.Produce_Qty, row._Decimal)
                            .Cells(i.ColQtyTarget.Index).Value = ObjCD.DigitGrouping(row.Target_Qty, row._Decimal)
                            .Cells(i.ColSchedule.Index).Value = Boolean.FalseString
                            .Cells(i.ColUnit.Index).Value = row.Unit_Code
                            If row.Target_Finish_Date_R2 <> #1/1/1950# Then
                                .Cells(i.ColTargetFinishDate.Index).Value = row.Target_Finish_Date_R2
                            ElseIf row.Target_Finish_Date_R1 <> #1/1/1950# Then
                                .Cells(i.ColTargetFinishDate.Index).Value = row.Target_Finish_Date_R1
                            Else
                                .Cells(i.ColTargetFinishDate.Index).Value = row.Target_Finish_Date
                            End If
                        End With
                    End If
                Next
            End If

            vDGV.Sort(i.ColTargetFinishDate, System.ComponentModel.ListSortDirection.Ascending)
            ObjCD.SetNumber(vDGV)
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
            TBL1.Dispose()
            DTA.Dispose()
            TBL.Dispose()
            ObjCD = Nothing
        End Try
    End Sub

    Public Function GetNextIdentifierJO(ByVal vDGVItem As DataGridView, ByVal vPrimary As String, _
      ByVal Frm As FrmTrans_ManufacturingOrder) As String
        Dim rowDGV As DataGridViewRow
        Dim MaxVal As Double = 0
        Dim Index As Integer = -1

        Try
            For Each rowDGV In vDGVItem.Rows
                If rowDGV.Cells(Frm.ColItem_Identifier.Index).Value <> "" Then
                    Index = rowDGV.Cells(Frm.ColItem_Identifier.Index).Value.ToString.LastIndexOf("-")
                    If Index <> -1 Then
                        If MaxVal < CDbl(rowDGV.Cells(Frm.ColItem_Identifier.Index).Value.ToString.Substring(Index + 1)) Then
                            MaxVal = rowDGV.Cells(Frm.ColItem_Identifier.Index).Value.ToString.Substring(Index + 1)
                        End If
                    End If
                End If
            Next
            Return vPrimary + "-" + (MaxVal + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-0"
        End Try
    End Function

    Public Function GetNextIdentifierWO(ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, ByVal vPrimary As String, _
      ByVal Frm As FrmTrans_ManufacturingOrder) As String
        Dim rowItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim MaxVal As Double = 0
        Dim Index As Integer = -1
        Try
            For Each rowItemWO In vTblItemWO.Rows
                If rowItemWO.Identifier <> "" Then
                    Index = rowItemWO.Identifier.ToString.LastIndexOf("-")
                    If Index <> -1 Then
                        If MaxVal < CDbl(rowItemWO.Identifier.ToString.Substring(Index + 1)) Then
                            MaxVal = rowItemWO.Identifier.ToString.Substring(Index + 1)
                        End If
                    End If
                End If
            Next
            Return vPrimary + "-" + (MaxVal + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-0"
        End Try
    End Function

    Public Function GetNextIdentifierBOM(ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, ByVal vPrimary As String) As String
        Dim rowItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim MaxVal As Double = 0
        Dim Index As Integer = -1
        Try
            For Each rowItemBOM In vTblItemBOM.Rows
                If rowItemBOM.Identifier_BOM <> "" Then
                    Index = rowItemBOM.Identifier_BOM.ToString.LastIndexOf("-")
                    If Index <> -1 Then
                        If MaxVal < CDbl(rowItemBOM.Identifier_BOM.ToString.Substring(Index + 1)) Then
                            MaxVal = rowItemBOM.Identifier_BOM.ToString.Substring(Index + 1)
                        End If
                    End If
                End If
            Next
            Return vPrimary + "-" + (MaxVal + 1).ToString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-1-0"
        End Try
    End Function

    Public Sub UpdateQtySpareMP(ByVal vPrimaryMO As Double, ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim TBL As New DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable
        Dim row, rowActual As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_PrimaryMO4Spare(TBL, vPrimaryMO)

            'untuk setiap baris, set mp spare = 0
            For Each rowActual In vTblItemRM.Rows
                rowActual.MP_Spare_Qty = 0
            Next

            'untuk setiap baris, cari padanannya, lalu tentukan qty spare nya
            If TBL.Rows.Count > 0 Then
                For Each rowActual In vTblItemRM.Rows
                    For Each row In TBL.Rows
                        If row.Identifier_JO = rowActual.Identifier_JO And row.Identifier_WO = rowActual.Identifier_WO _
                        And row.Primary_Product_Code_RM = rowActual.Primary_Product_Code_RM Then
                            rowActual.MP_Spare_Qty = row.MP_Spare_Qty
                            Exit For
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    'fungsi ini tidak meghapus item sch, karena asumsi item sch itu bukan urusan form mo
    Public Function HapusAllItem(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTAItemJO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemTableAdapter
        Dim DTAItemBom As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_bomTableAdapter
        Dim DTAProsesOrder As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_processorderTableAdapter
        Dim DTAItemWO As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_woTableAdapter
        Dim DTAItemRM As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter
        Dim DTAItemRMAdd As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rm_addTableAdapter
        Dim DTAWIP As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_wipTableAdapter
        Dim DTAItemDelivery As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemdeliveryTableAdapter
        Try
            DTAItemJO.Connection.ConnectionString = Core_Module.GetConstring
            DTAItemBom.Connection.ConnectionString = DTAItemJO.Connection.ConnectionString
            DTAProsesOrder.Connection.ConnectionString = DTAItemJO.Connection.ConnectionString
            DTAItemWO.Connection.ConnectionString = DTAItemJO.Connection.ConnectionString
            DTAItemRM.Connection.ConnectionString = DTAItemJO.Connection.ConnectionString
            DTAItemRMAdd.Connection.ConnectionString = DTAItemJO.Connection.ConnectionString
            DTAWIP.Connection.ConnectionString = DTAItemJO.Connection.ConnectionString
            DTAItemDelivery.Connection.ConnectionString = Core_Module.GetConstring

            'hapus item MO
            DTAItemJO.Delete(vPrimaryTransaksi)
            DTAItemBom.Delete(vPrimaryTransaksi)
            DTAProsesOrder.Delete(vPrimaryTransaksi)
            DTAItemWO.Delete(vPrimaryTransaksi)
            DTAItemRM.Delete(vPrimaryTransaksi)
            DTAItemRMAdd.Delete(vPrimaryTransaksi)
            DTAWIP.DeleteQuery(vPrimaryTransaksi)
            DTAItemDelivery.DeleteQuery(vPrimaryTransaksi)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTAItemJO.Dispose()
            DTAItemBom.Dispose()
            DTAProsesOrder.Dispose()
            DTAItemWO.Dispose()
            DTAItemRM.Dispose()
        End Try
    End Function

    Public Function SimpanAllItem(ByVal vPrimaryMO As Double, ByVal vDgvItemJO As DataGridView, _
    ByVal vTblItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vUserCode As String, ByVal vFrmMO As FrmTrans_ManufacturingOrder, _
    ByVal vTblItemRMAdd As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vTblItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable) As Boolean
        Try
            'update status RM untuk mengetahui product merupakan RM dan bukan WIP (tidak memiliki bom)
            Me.UpdateStatusRM(vDgvItemJO, vTblItemWO, vTblItemRM, vFrmMO)
            'simpan item JO
            Me.SimpanItemJO_UsingScript(vPrimaryMO, vDgvItemJO, vFrmMO)
            'simpan item BOM
            Me.SimpanItemBOM_UsingScript(vPrimaryMO, vTblItemBom)
            'simpan item Proses Order
            Me.SimpanProsesOrder_UsingScript(vPrimaryMO, vTblProsesOrder)
            'simpan item WO
            Me.SimpanItemWO_UsingScript(vPrimaryMO, vTblItemWO, vUserCode)
            'simpan item RM
            Me.SimpanItemRM_UsingScript(vPrimaryMO, vTblItemRM)
            Me.SimpanItemRM_Add_UsingScript(vPrimaryMO, vTblItemRMAdd)
            'simpan item wip
            Me.SimpanItemWIP_UsingScript(vPrimaryMO, vTblItemWIP)

            Dim dtaItemDelivery As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_itemdeliveryTableAdapter
            Dim rowItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow
            dtaItemDelivery.Connection.ConnectionString = Core_Module.GetConstring
            For Each rowItemDelivery In vTblItemDelivery.Rows
                With rowItemDelivery
                    If .identifier_jo.StartsWith("-1") = True Then
                        .identifier_jo = vPrimaryMO.ToString + .identifier_jo.Substring(2)
                    End If
                    dtaItemDelivery.InsertQuery(vPrimaryMO, .identifier_jo, .delivery_qty, .delivery_date, .remarks)
                End With
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanItemWIP_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTblItemWIP.Rows.Count
                sql = Me.GetInsertString_ItemWIP(StartRowIndex, BatchQty, vTblItemWIP, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemWIP(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblItemWIP As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_wipRow
        Dim IdentifierJO As String = ""
        Dim IdentifierWO As String = ""

        Try
            retval = "INSERT INTO `tbltrans_mo_item_wip` (`Primary MO`, `Identifier JO`, `Identifier WO`, `Primary Product Code WIP`, `Qty`, `Remarks`, `Qty Used`, `Qty Balance`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblItemWIP.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblItemWIP.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblItemWIP.Rows(vStartRowIndex + i)

                    'set identifier JO
                    IdentifierJO = row.Identifier_JO
                    If IdentifierJO.ToString.StartsWith("-1") Then
                        IdentifierJO = vPrimaryTransaksi.ToString + IdentifierJO.ToString.Substring(IdentifierJO.IndexOf("-", 1))
                        row.Identifier_JO = IdentifierJO
                    End If

                    'set identifier WO
                    IdentifierWO = row.Identifier_WO
                    If IdentifierWO.ToString.StartsWith("-1") Then
                        IdentifierWO = vPrimaryTransaksi.ToString + IdentifierWO.ToString.Substring(IdentifierWO.IndexOf("-", 1))
                        row.Identifier_WO = IdentifierWO
                    End If

                    'set balance qty
                    row.Qty_Balance = row.Qty - row.Qty_Used
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + "'" + IdentifierJO.ToString + "'," _
                    + "'" + IdentifierWO.ToString + "'," _
                    + row.Primary_Product_Code_WIP.ToString + "," _
                    + row.Qty.ToString + "," _
                    + "'" + row.Remarks + "'," _
                    + row.Qty_Used.ToString + "," _
                    + row.Qty_Balance.ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SimpanItemRM_Add_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTblItemRM.Rows.Count
                sql = Me.GetInsertString_ItemRMAdd(StartRowIndex, BatchQty, vTblItemRM, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemRMAdd(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_rm_addRow
        Dim IdentifierJO As String = ""
        Dim IdentifierWO As String = ""

        Try
            retval = "INSERT INTO `tbltrans_mo_item_rm_add` (`Primary MO`, `Identifier JO`, `Identifier WO`, `Primary Product Code RM`, `Qty`, `Remarks`, `Add Date`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblItemRM.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblItemRM.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblItemRM.Rows(vStartRowIndex + i)

                    'set identifier JO
                    IdentifierJO = row.Identifier_JO
                    If IdentifierJO.ToString.StartsWith("-1") Then
                        IdentifierJO = vPrimaryTransaksi.ToString + IdentifierJO.ToString.Substring(IdentifierJO.IndexOf("-", 1))
                        row.Identifier_JO = IdentifierJO
                    End If

                    'set identifier WO
                    IdentifierWO = row.Identifier_WO
                    If IdentifierWO.ToString.StartsWith("-1") Then
                        IdentifierWO = vPrimaryTransaksi.ToString + IdentifierWO.ToString.Substring(IdentifierWO.IndexOf("-", 1))
                        row.Identifier_WO = IdentifierWO
                    End If

                    'set balance qty
                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + "'" + IdentifierJO.ToString + "'," _
                    + "'" + IdentifierWO.ToString + "'," _
                    + row.Primary_Product_Code_RM.ToString + "," _
                    + row.Qty.ToString + "," _
                    + "'" + row.Remarks + "'," _
                    + "'" + Format(row.Add_Date, "yyyy-MM-dd") + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Private Sub UpdateStatusRM(ByVal vDgvItemJO As DataGridView, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, ByVal vFrm As FrmTrans_ManufacturingOrder)
        Dim rowRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim rowD As DataGridViewRow

        Try
            For i As Integer = 0 To vDgvItemJO.RowCount - 2
                rowD = vDgvItemJO.Rows(i)
                vTblItemRM.DefaultView.RowFilter = "[Identifier JO] = '" & rowD.Cells(vFrm.ColItem_Identifier.Index).Value & "'"
                If vTblItemRM.DefaultView.Count > 0 Then
                    For Each dv As Data.DataRowView In vTblItemRM.DefaultView
                        rowRM = dv.Row
                        vTblItemWO.DefaultView.RowFilter = "[Identifier JO] = '" & rowRM.Identifier_JO & "' AND [Primary Product] = " _
                        & rowRM.Primary_Product_Code_RM
                        If vTblItemWO.DefaultView.Count = 0 Then
                            rowRM.Status_RM = 1
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function SimpanItemJO_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, ByVal Frm As FrmTrans_ManufacturingOrder) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vDGVItem.Rows.Count
                sql = Me.GetInsertString_ItemJO(StartRowIndex, BatchQty, vDGVItem, vPrimaryTransaksi, Frm)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemJO(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, ByVal vdgvItem As DataGridView, _
    ByVal vPrimaryTransaksi As Double, ByVal ij As FrmTrans_ManufacturingOrder) As String
        Dim retval As String = ""
        Dim row As DataGridViewRow
        Dim Delimiter As String = ","
        Dim TargetFinishDate, TargetFinishDateR1, TargetFinishDateR2, ActualFinishDate As Date
        Dim IdentifierJO As String = ""

        Try
            vdgvItem.AllowUserToAddRows = False
            retval = "INSERT INTO `tbltrans_mo_item` (`Primary MO`, `JO Number`, `Urutan`, `Primary Product`, `Target Qty`, `Produce Qty`, `Remarks`, `Is Closed`, `Identifier`, `Decimal Product`, `Actual Finish Date`, `Target Finish Date`, `Target Finish Date R1`, `Target Finish Date R2`) VALUES "
            vdgvItem.AllowUserToAddRows = False
            For i As Double = 0 To BatchQty - 1 Step 1
                If vdgvItem.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vdgvItem.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vdgvItem.Rows(vStartRowIndex + i)

                    'set identifier
                    IdentifierJO = row.Cells(ij.ColItem_Identifier.Index).Value
                    If IdentifierJO.ToString.StartsWith("-1") Then
                        IdentifierJO = vPrimaryTransaksi.ToString + IdentifierJO.ToString.Substring(IdentifierJO.IndexOf("-", 1))
                        row.Cells(ij.ColItem_Identifier.Index).Value = IdentifierJO
                    End If

                    'set format tanggal
                    If IsDate(row.Cells(ij.ColItem_TargetFinishDate.Index).Value) = False Then
                        TargetFinishDate = #1/1/1950#
                    Else
                        TargetFinishDate = CDate(row.Cells(ij.ColItem_TargetFinishDate.Index).Value)
                    End If
                    If IsDate(row.Cells(ij.ColItem_TargetFinishDateR1.Index).Value) = True Then
                        TargetFinishDateR1 = CDate(row.Cells(ij.ColItem_TargetFinishDateR1.Index).Value)
                    Else
                        TargetFinishDateR1 = #1/1/1950#
                    End If
                    If IsDate(row.Cells(ij.ColItem_TargetFinishDateR2.Index).Value) = True Then
                        TargetFinishDateR2 = CDate(row.Cells(ij.ColItem_TargetFinishDateR2.Index).Value)
                    Else
                        TargetFinishDateR2 = #1/1/1950#
                    End If
                    If IsDate(row.Cells(ij.ColItem_ActualFinishDate.Index).Value) = True Then
                        ActualFinishDate = CDate(row.Cells(ij.ColItem_ActualFinishDate.Index).Value)
                    Else
                        ActualFinishDate = #1/1/1950#
                    End If


                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + "'" + row.Cells(ij.ColItem_JONumber.Index).Value.ToString + "'," _
                    + CDbl(row.Cells(ij.ColItem_Urutan.Index).Value).ToString + "," _
                    + CDbl(row.Cells(ij.ColItem_PrimaryProduct.Index).Value).ToString + "," _
                    + CDec(row.Cells(ij.ColItem_QtyTarget.Index).Value).ToString + "," _
                    + CDec(row.Cells(ij.ColItem_QtyProduce.Index).Value).ToString + "," _
                    + "'" + row.Cells(ij.ColItem_Remarks.Index).Value.ToString + "'," _
                    + CInt(row.Cells(ij.ColItem_IsClosed.Index).Value).ToString + "," _
                    + "'" + IdentifierJO.ToString + "'," _
                    + CDec(row.Cells(ij.ColItem_Decimal.Index).Value).ToString + "," _
                    + "'" + Format(ActualFinishDate, "yyyy-MM-dd") + "'," _
                    + "'" + Format(TargetFinishDate, "yyyy-MM-dd") + "'," _
                    + "'" + Format(TargetFinishDateR1, "yyyy-MM-dd") + "'," _
                    + "'" + Format(TargetFinishDateR2, "yyyy-MM-dd") + "'" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next

            vdgvItem.AllowUserToAddRows = True
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SimpanItemBOM_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal vTblItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTblItemBom.Rows.Count
                sql = Me.GetInsertString_ItemBOM(StartRowIndex, BatchQty, vTblItemBom, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemBOM(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblItemBOM As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim IdentifierJO As String = ""
        Dim IdentifierWO As String = ""
        Dim IdentifierBOM As String = ""
        Dim IdentifierBOM_Parent As String = ""

        Try
            retval = "INSERT INTO `tbltrans_mo_item_bom` (`Primary MO`, `Identifier JO`, `Primary Product`, `Primary Process`, `Qty Ref Process`, `Good Percentage Process`, `Primary Product RM`, `Qty RM`, `Good Percentage RM`, `Decimal RM`, `Identifier WO`, `Remarks`, `Primary Parent Product`, `Identifier BOM`, `Identifier BOM Parent`, `tolerance process`, `production tolerance`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblItemBOM.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblItemBOM.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblItemBOM.Rows(vStartRowIndex + i)

                    'set identifier JO
                    IdentifierJO = row.Identifier_JO
                    If IdentifierJO.ToString.StartsWith("-1") Then
                        IdentifierJO = vPrimaryTransaksi.ToString + IdentifierJO.ToString.Substring(IdentifierJO.IndexOf("-", 1))
                        row.Identifier_JO = IdentifierJO
                    End If

                    'set identifier WO
                    IdentifierWO = row.Identifier_WO
                    If IdentifierWO.ToString.StartsWith("-1") Then
                        IdentifierWO = vPrimaryTransaksi.ToString + IdentifierWO.ToString.Substring(IdentifierWO.IndexOf("-", 1))
                        row.Identifier_WO = IdentifierWO
                    End If

                    'set identifier BOM
                    IdentifierBOM = row.Identifier_BOM
                    If IdentifierBOM.ToString.StartsWith("-1") Then
                        IdentifierBOM = vPrimaryTransaksi.ToString + IdentifierBOM.ToString.Substring(IdentifierBOM.IndexOf("-", 1))
                        row.Identifier_BOM = IdentifierBOM
                    End If

                    'set identifier BOM Parent
                    IdentifierBOM_Parent = row.Identifier_BOM_Parent
                    If IdentifierBOM_Parent.ToString.StartsWith("-1") Then
                        IdentifierBOM_Parent = vPrimaryTransaksi.ToString + IdentifierBOM_Parent.ToString.Substring(IdentifierBOM_Parent.IndexOf("-", 1))
                        row.Identifier_BOM_Parent = IdentifierBOM_Parent
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + "'" + IdentifierJO.ToString + "'," _
                    + row.Primary_Product.ToString + "," _
                    + row.Primary_Process.ToString + "," _
                    + row.Qty_Ref_Process.ToString + "," _
                    + row.Good_Percentage_Process.ToString + "," _
                    + row.Primary_Product_RM.ToString + "," _
                    + row.Qty_RM.ToString + "," _
                    + row.Good_Percentage_RM.ToString + "," _
                    + row.Decimal_RM.ToString + "," _
                    + "'" + IdentifierWO.ToString + "'," _
                    + "'" + row.Remarks.ToString + "'," _
                    + row.Primary_Parent_Product.ToString + "," _
                    + "'" + IdentifierBOM.ToString + "'," _
                    + "'" + IdentifierBOM_Parent.ToString + "'," _
                    + row.tolerance_process.ToString + "," _
                    + row.production_tolerance.ToString _
                    + ")" + Delimiter

                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SimpanProsesOrder_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTblProsesOrder.Rows.Count
                sql = Me.GetInsertString_ProsesOrder(StartRowIndex, BatchQty, vTblProsesOrder, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ProsesOrder(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim IdentifierJO As String = ""

        Try
            retval = "INSERT INTO `tbltrans_mo_processorder` (`primary mo`, `primary process`, `primary product`, `identifier jo`, `urutan`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblProsesOrder.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblProsesOrder.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblProsesOrder.Rows(vStartRowIndex + i)

                    'set identifier JO
                    IdentifierJO = row.identifier_jo
                    If IdentifierJO.ToString.StartsWith("-1") Then
                        IdentifierJO = vPrimaryTransaksi.ToString + IdentifierJO.ToString.Substring(IdentifierJO.IndexOf("-", 1))
                        row.identifier_jo = IdentifierJO
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + row.primary_process.ToString + "," _
                    + row.primary_product.ToString + "," _
                    + "'" + IdentifierJO.ToString + "'," _
                    + row.urutan.ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SimpanItemWO_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vUserCode As String) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTblItemWO.Rows.Count
                sql = Me.GetInsertString_ItemWO(StartRowIndex, BatchQty, vTblItemWO, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemWO(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblItemWO As DseTrans_ManufacturingOrder.tbltrans_mo_item_woDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_woRow
        Dim IdentifierJO As String = ""
        Dim IdentifierWO As String = ""
        Dim IdentifierWO_Parent As String = ""

        Try
            retval = "INSERT INTO `tbltrans_mo_item_wo` (`Primary MO`, `Identifier JO`, `WO Number`, `Urutan`, `Date`, `Status`, `Primary Process`, `Remarks`, `Identifier`, `Primary Work Unit`, `Plan Start Time`, `Plan End Time`, `Plan Setup Time`, `Plan Work Hour`, `Plan Delay Time`, `Plan Total Time`, `Plan Prod Rate`, `Actual Start Time`, `Actual End Time`, `Actual Setup Time`, `Actual Work Hour`, `Actual Total Time`, `Actual Down Time`, `Actual Prod Rate`, `Actual Delay Time`, `Plan Good Perct`, `Actual Good Perct`, `Qty Target Result`, `Qty Target Prod`, `Qty Target Prod Add`, `Qty Actual Prod`, `Qty Actual Result`, `Qty Projected Result`, `Created By`, `Created Date`, `Modified By`, `Modified Date`, `Timestamp`, `Identifier WO Parent`, `Level`, `Primary Product`, `Balance Produce`, `Balance Result`, `Qty Balance MP`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblItemWO.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblItemWO.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblItemWO.Rows(vStartRowIndex + i)

                    'set identifier JO
                    IdentifierJO = row.Identifier_JO
                    If IdentifierJO.ToString.StartsWith("-1") Then
                        IdentifierJO = vPrimaryTransaksi.ToString + IdentifierJO.ToString.Substring(IdentifierJO.IndexOf("-", 1))
                        row.Identifier_JO = IdentifierJO
                    End If

                    'set identifier WO
                    IdentifierWO = row.Identifier
                    If IdentifierWO.ToString.StartsWith("-1") Then
                        IdentifierWO = vPrimaryTransaksi.ToString + IdentifierWO.ToString.Substring(IdentifierWO.IndexOf("-", 1))
                        row.Identifier = IdentifierWO
                    End If

                    'set identifier WO Parent
                    If row.Identifier_WO_Parent <> "" Then
                        IdentifierWO_Parent = row.Identifier_WO_Parent
                        If IdentifierWO_Parent.ToString.StartsWith("-1") Then
                            IdentifierWO_Parent = vPrimaryTransaksi.ToString + IdentifierWO_Parent.ToString.Substring(IdentifierWO_Parent.IndexOf("-", 1))
                            row.Identifier_WO_Parent = IdentifierWO_Parent
                        End If
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + "'" + IdentifierJO.ToString + "'," _
                    + "'" + row.WO_Number.ToString + "'," _
                    + row.Urutan.ToString + "," _
                    + "'" + Format(row._Date, "yyyy-MM-dd") + "'," _
                    + row.Status.ToString + "," _
                    + row.Primary_Process.ToString + "," _
                    + "'" + row.Remarks.ToString + "'," _
                    + "'" + IdentifierWO.ToString + "'," _
                    + row.Primary_Work_Unit.ToString + "," _
                    + "'" + Format(row.Plan_Start_Time, "yyyy-MM-dd hh:mm:ss") + "'," _
                    + "'" + Format(row.Plan_End_Time, "yyyy-MM-dd hh:mm:ss") + "'," _
                    + row.Plan_Setup_Time.ToString + "," _
                    + row.Plan_Work_Hour.ToString + "," _
                    + row.Plan_Delay_Time.ToString + "," _
                    + row.Plan_Total_Time.ToString + "," _
                    + row.Plan_Prod_Rate.ToString + "," _
                    + "'" + Format(row.Actual_Start_Time, "yyyy-MM-dd hh:mm:ss") + "'," _
                    + "'" + Format(row.Actual_End_Time, "yyyy-MM-dd hh:mm:ss") + "'," _
                    + row.Actual_Setup_Time.ToString + "," _
                    + row.Actual_Work_Hour.ToString + "," _
                    + row.Actual_Total_Time.ToString + "," _
                    + row.Actual_Down_Time.ToString + "," _
                    + row.Actual_Prod_Rate.ToString + "," _
                    + row.Actual_Delay_Time.ToString + "," _
                    + row.Plan_Good_Perct.ToString + "," _
                    + row.Actual_Good_Perct.ToString + "," _
                    + row.Qty_Target_Result.ToString + "," _
                    + row.Qty_Target_Prod.ToString + "," _
                    + row.Qty_Target_Prod_Add.ToString + "," _
                    + row.Qty_Actual_Prod.ToString + "," _
                    + row.Qty_Actual_Result.ToString + "," _
                    + row.Qty_Projected_Result.ToString + "," _
                    + "'" + row.Created_By.ToString + "'," _
                    + "'" + Format(row.Created_Date, "yyyy-MM-dd hh:mm:ss") + "'," _
                    + "'" + row.Modified_By.ToString + "'," _
                    + "'" + Format(row.Modified_Date, "yyyy-MM-dd hh:mm:ss") + "'," _
                    + "'" + Format(row.Timestamp, "yyyy-MM-dd hh:mm:ss") + "'," _
                    + "'" + IdentifierWO_Parent.ToString + "'," _
                    + row.Level.ToString + "," _
                    + row.Primary_Product.ToString + "," _
                    + row.Balance_Produce.ToString + "," _
                    + row.Balance_Result.ToString + "," _
                    + row.Qty_Balance_MP.ToString + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function SimpanItemRM_UsingScript(ByVal vPrimaryTransaksi As Double, ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable) As Boolean
        Dim conStr As String = My.Settings.core_dbConnectionString
        Dim Conn As New MySql.Data.MySqlClient.MySqlConnection(conStr)
        Dim sql As String = ""
        Dim script As New MySql.Data.MySqlClient.MySqlScript(Conn, sql)
        Dim BatchQty As Double = 500
        Dim StartRowIndex As Double = 0

        Try
            'open connection
            Conn.Open()

            While StartRowIndex < vTblItemRM.Rows.Count
                sql = Me.GetInsertString_ItemRM(StartRowIndex, BatchQty, vTblItemRM, vPrimaryTransaksi)
                script.Query = sql
                'execute script
                script.Execute()
                StartRowIndex += BatchQty
            End While

            Conn.Close()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Conn = Nothing
            script = Nothing
        End Try
    End Function

    Private Function GetInsertString_ItemRM(ByVal vStartRowIndex As Double, ByVal BatchQty As Double, _
    ByVal vTblItemRM As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmDataTable, _
    ByVal vPrimaryTransaksi As Double) As String
        Dim retval As String = ""
        Dim Delimiter As String = ","
        Dim row As DseTrans_ManufacturingOrder.tbltrans_mo_item_rmRow
        Dim IdentifierJO As String = ""
        Dim IdentifierWO As String = ""

        Try
            retval = "INSERT INTO `tbltrans_mo_item_rm` (`Primary MO`, `Identifier JO`, `Identifier WO`, `Primary Product Code RM`, `Target Qty`, `Required Qty`," + _
            " `Allocated Qty`, `Additional Qty`, `Used Qty`, `Balance Qty`, `Primary StandBy Location`, `Primary PickUp Location`, `Remarks`, `GP Target`, `GP Actual`," + _
            " `Status RM`, `is wip`,`WIP Qty`,`is replacement`,`primary code rm replaced`,`ratio replacement`,`ratio replaced`, `qty replaced`,`primary unit replaced`,`Replace Qty`, `use priority`, `MP Spare Qty`,`required date`) VALUES "

            For i As Double = 0 To BatchQty - 1 Step 1
                If vTblItemRM.Rows.Count > vStartRowIndex + i Then
                    'jika baris terakhir dgv atau baris terakhir untuk batch tsb
                    If vTblItemRM.Rows.Count - 1 = vStartRowIndex + i Or i = BatchQty - 1 Then
                        Delimiter = ";"
                    End If

                    row = vTblItemRM.Rows(vStartRowIndex + i)

                    'set identifier JO
                    IdentifierJO = row.Identifier_JO
                    If IdentifierJO.ToString.StartsWith("-1") Then
                        IdentifierJO = vPrimaryTransaksi.ToString + IdentifierJO.ToString.Substring(IdentifierJO.IndexOf("-", 1))
                        row.Identifier_JO = IdentifierJO
                    End If

                    'set identifier WO
                    IdentifierWO = row.Identifier_WO
                    If IdentifierWO.ToString.StartsWith("-1") Then
                        IdentifierWO = vPrimaryTransaksi.ToString + IdentifierWO.ToString.Substring(IdentifierWO.IndexOf("-", 1))
                        row.Identifier_WO = IdentifierWO
                    End If

                    'set balance qty
                    If row.is_replacement = 1 Then
                        row.Balance_Qty = row.Required_Qty - row.Used_Qty - row.WIP_Qty - row.Replace_Qty
                    Else
                        row.Balance_Qty = row.Allocated_Qty + row.Additional_Qty - row.Used_Qty - row.WIP_Qty - row.Replace_Qty
                    End If

                    retval = retval + vbCrLf + "(" + vPrimaryTransaksi.ToString + "," _
                    + "'" + IdentifierJO.ToString + "'," _
                    + "'" + IdentifierWO.ToString + "'," _
                    + row.Primary_Product_Code_RM.ToString + "," _
                    + row.Target_Qty.ToString + "," _
                    + row.Required_Qty.ToString + "," _
                    + row.Allocated_Qty.ToString + "," _
                    + row.Additional_Qty.ToString + "," _
                    + row.Used_Qty.ToString + "," _
                    + row.Balance_Qty.ToString + "," _
                    + row.Primary_StandBy_Location.ToString + "," _
                    + row.Primary_PickUp_Location.ToString + "," _
                    + "'" + row.Remarks.ToString + "'," _
                    + row.GP_Target.ToString + "," _
                    + row.GP_Actual.ToString + "," _
                    + row.Status_RM.ToString + "," _
                    + row.is_wip.ToString + "," _
                    + row.WIP_Qty.ToString + "," _
                    + row.is_replacement.ToString + "," _
                    + row.primary_code_rm_replaced.ToString + "," _
                    + row.ratio_replacement.ToString + "," _
                    + row.ratio_replaced.ToString + "," _
                    + row.qty_replaced.ToString + "," _
                    + row.primary_unit_replaced.ToString + "," _
                    + row.Replace_Qty.ToString + "," _
                    + row.use_priority.ToString + "," _
                    + row.MP_Spare_Qty.ToString + "," _
                    + "'" + Format(row.required_date, "yyyy-MM-dd").ToString + "'" + "" _
                    + ")" + Delimiter
                Else
                    Exit For
                End If
            Next
            retval = retval + vbCrLf + vbCrLf + "COMMIT;"
            Return retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function IsiDGVItemJO(ByVal vPrimaryTransaksi As Double, ByVal vDGVItem As DataGridView, _
    ByVal vTblJO_Temp As DseTrans_ManufacturingOrder.JO_TempDataTable, ByVal vFrm As FrmTrans_ManufacturingOrder, _
    ByVal vObjCD As Core_Module.ClsCekData, ByVal vTBlItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryDataTable) As Boolean
        Dim DTAItem As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item1TableAdapter
        Dim TBLItem As New DseTrans_ManufacturingOrder.tbltrans_mo_item1DataTable
        Dim rowItem As DseTrans_ManufacturingOrder.tbltrans_mo_item1Row
        Dim rowJO_Temp As DseTrans_ManufacturingOrder.JO_TempRow
        Dim rowDGV As DataGridViewRow
        Dim count As Integer = 0
        Dim rowItemDelivery As DseTrans_ManufacturingOrder.tbltrans_mo_itemdeliveryRow
        Try
            DTAItem.Connection.ConnectionString = Core_Module.GetConstring
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            vDGVItem.AllowUserToAddRows = False
            vDGVItem.Rows.Clear()

            If DTAItem.FillBy_PrimaryTransaksi(TBLItem, vPrimaryTransaksi) > 0 Then
                For Each rowItem In TBLItem
                    vDGVItem.Rows.Add()
                    count += 1
                    If count = 1 And vDGVItem.Rows.Count - 1 = 1 Then
                        vDGVItem.AllowUserToAddRows = False
                    End If
                    rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                    With rowDGV
                        .Cells(vFrm.ColItem_No.Index).Value = ""
                        .Cells(vFrm.ColItem_JONumber.Index).Value = rowItem.JO_Number
                        .Cells(vFrm.ColItem_PrimaryProduct.Index).Value = rowItem.Primary_Product
                        .Cells(vFrm.ColItem_CodeNo.Index).Value = rowItem.Product_Code
                        .Cells(vFrm.ColItem_Desc.Index).Value = rowItem.Description
                        .Cells(vFrm.ColItem_Unit.Index).Value = rowItem.Unit_Code
                        .Cells(vFrm.ColItem_QtyProduce.Index).Value = vObjCD.DigitGrouping(rowItem.Produce_Qty, rowItem.Decimal_Product)
                        .Cells(vFrm.ColItem_QtyTarget.Index).Value = vObjCD.DigitGrouping(rowItem.Target_Qty, rowItem.Decimal_Product)
                        If rowItem.Target_Finish_Date = CDate("1/1/1950") Then
                            .Cells(vFrm.ColItem_TargetFinishDate.Index).Value = "-"
                        Else
                            .Cells(vFrm.ColItem_TargetFinishDate.Index).Value = Format(rowItem.Target_Finish_Date.Date, "dd MMM yyyy")
                        End If
                        If rowItem.Target_Finish_Date_R1 = CDate("1/1/1950") Then
                            .Cells(vFrm.ColItem_TargetFinishDateR1.Index).Value = "-"
                        Else
                            .Cells(vFrm.ColItem_TargetFinishDateR1.Index).Value = Format(rowItem.Target_Finish_Date_R1, "dd MMM yyyy")
                        End If
                        If rowItem.Target_Finish_Date_R2 = CDate("1/1/1950") Then
                            .Cells(vFrm.ColItem_TargetFinishDateR2.Index).Value = "-"
                        Else
                            .Cells(vFrm.ColItem_TargetFinishDateR2.Index).Value = Format(rowItem.Target_Finish_Date_R2, "dd MMM yyyy")
                        End If
                        If rowItem.Actual_Finish_Date = CDate("1/1/1950") Then
                            .Cells(vFrm.ColItem_ActualFinishDate.Index).Value = "-"
                        Else
                            .Cells(vFrm.ColItem_ActualFinishDate.Index).Value = Format(rowItem.Actual_Finish_Date, "dd MMM yyyy")
                        End If
                        .Cells(vFrm.ColItem_Remarks.Index).Value = rowItem.Remarks
                        .Cells(vFrm.ColItem_Urutan.Index).Value = rowItem.Urutan
                        .Cells(vFrm.ColItem_Identifier.Index).Value = rowItem.Identifier
                        .Cells(vFrm.ColItem_Decimal.Index).Value = rowItem.Decimal_Product
                        vTBlItemDelivery.DefaultView.RowFilter = "[identifier jo] = '" & rowItem.Identifier & "'"
                        If vTBlItemDelivery.DefaultView.Count = 0 Then
                            .Cells(vFrm.ColItem_DeliveryDate.Index).Value = "-"
                        ElseIf vTBlItemDelivery.DefaultView.Count = 1 Then
                            rowItemDelivery = vTBlItemDelivery.DefaultView(0).Row
                            .Cells(vFrm.ColItem_DeliveryDate.Index).Value = Format(rowItemDelivery.delivery_date, "dd MMM yyyy")
                        ElseIf vTBlItemDelivery.DefaultView.Count > 1 Then
                            .Cells(vFrm.ColItem_DeliveryDate.Index).Value = "Multiple"
                        End If

                        If IsNothing(vTblJO_Temp) = False Then
                            If rowItem.Primary_Product <> -1 Then
                                rowJO_Temp = vTblJO_Temp.NewJO_TempRow
                                rowJO_Temp.Primary_Product = rowItem.Primary_Product
                                rowJO_Temp.Identifier_JO = rowItem.Identifier
                                vTblJO_Temp.Rows.Add(rowJO_Temp)
                            End If
                        End If
                    End With
                Next
            End If

            'vDGVItem.AllowUserToAddRows = True
            vObjCD.SetNumber(vDGVItem)
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            DTAItem.Dispose()
            TBLItem = Nothing
        End Try
    End Function

    Public Sub IsiDgvItemBOM(ByVal vIdentifierJO As String, ByVal vIdentifierBOM_Parent As String, ByVal vDgvItemBOM As DataGridView, ByVal vTblItemBOM As _
    DseTrans_ManufacturingOrder.tbltrans_mo_item_bomDataTable, ByVal vFrm As FrmTrans_ManufacturingOrder_BOM)
        Dim rowItemBom As DseTrans_ManufacturingOrder.tbltrans_mo_item_bomRow
        Dim rowDgv As DataGridViewRow
        Dim Dummy As Boolean = False
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim Count As Integer = 0

        Try
            'salin data bom
            vDgvItemBOM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvItemBOM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDgvItemBOM.Rows.Clear()
            Dummy = vDgvItemBOM.AllowUserToAddRows
            vDgvItemBOM.AllowUserToAddRows = False
            vTblItemBOM.DefaultView.RowFilter = "[Identifier JO] = '" & vIdentifierJO.ToString & "' AND [Identifier BOM Parent] = '" & vIdentifierBOM_Parent & "'"
            If vTblItemBOM.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblItemBOM.DefaultView
                    rowItemBom = dv.Row
                    Count += 1
                    If Count <= 1 Then
                        Dim objUnit As New Core_Module.ClsMasterUnit
                        vFrm.txtPrimaryProses.Text = rowItemBom.Primary_Process
                        vFrm.txtProcessCode.Text = ObjProses.GetCode(rowItemBom.Primary_Process)
                        vFrm.TxtProcessDesc.Text = ObjProses.GetDesc(rowItemBom.Primary_Process)
                        vFrm.TxtQtyRef.Text = ObjCD.DigitGrouping(rowItemBom.Qty_Ref_Process, ObjProd.GetAngkaDecimal(rowItemBom.Primary_Product))
                        vFrm.TxtGoodPercentage.Text = ObjCD.DigitGrouping(rowItemBom.Good_Percentage_Process, 2)
                        vFrm.txtTolerance.Text = ObjCD.DigitGrouping(rowItemBom.tolerance_process, 2)
                        vFrm.LblUnitBOM.Text = objUnit.GetUnitCode(ObjProd.GetPrimaryUnit(rowItemBom.Primary_Product))
                        vFrm.LblUnitBOM.Left = vFrm.Label1.Right - vFrm.LblUnitBOM.Width
                        vFrm.TxtQtyRef.Width = vFrm.LblUnitBOM.Left - 5 - vFrm.TxtQtyRef.Left
                    End If

                    vDgvItemBOM.Rows.Add()
                    rowDgv = vDgvItemBOM.Rows(vDgvItemBOM.Rows.Count - 1)

                    With rowItemBom
                        rowDgv.Cells(vFrm.ColBOM_PrimaryProduct.Index).Value = .Primary_Product_RM
                        rowDgv.Cells(vFrm.ColBOM_CodeNo.Index).Value = ObjProd.GetProductCode(.Primary_Product_RM)
                        rowDgv.Cells(vFrm.ColBOM_Desc.Index).Value = ObjProd.GetDescription(.Primary_Product_RM)
                        rowDgv.Cells(vFrm.ColBOM_DrillDown.Index).Value = "Edit"
                        rowDgv.Cells(vFrm.ColBOM_Qty.Index).Value = ObjCD.DigitGrouping(.Qty_RM, .Decimal_RM)
                        rowDgv.Cells(vFrm.ColBOM_GoodPercentage.Index).Value = ObjCD.DigitGrouping(.Good_Percentage_RM, 2)
                        rowDgv.Cells(vFrm.ColBOM_Unit.Index).Value = ObjProd.GetUnitCode(.Primary_Product_RM)
                        rowDgv.Cells(vFrm.ColBOM_Remarks.Index).Value = .Remarks
                        rowDgv.Cells(vFrm.ColBOM_PrimaryParentProduct.Index).Value = .Primary_Parent_Product
                        rowDgv.Cells(vFrm.ColBOM_IdentifierWO.Index).Value = .Identifier_WO
                        rowDgv.Cells(vFrm.ColBOM_IdentifierBOM.Index).Value = .Identifier_BOM
                        rowDgv.Cells(vFrm.ColBOM_IdentifierBOM_Parent.Index).Value = .Identifier_BOM_Parent

                        rowDgv.Cells(vFrm.ColBOM_IdentifierJO.Index).Value = .Identifier_JO
                        rowDgv.Cells(vFrm.ColBOM_Tolerance.Index).Value = ObjCD.DigitGrouping(.production_tolerance, 2)

                    End With
                Next
            End If

            vDgvItemBOM.AllowUserToAddRows = Dummy
            vDgvItemBOM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvItemBOM.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjProd = Nothing
            ObjProses = Nothing
        End Try
    End Sub

    Public Sub IsiDgvProsesOrder(ByVal vIdentifierJO As String, ByVal vDgvProsesOrder As DataGridView, ByVal vTblProsesOrder As _
    DseTrans_ManufacturingOrder.tbltrans_mo_processorderDataTable, ByVal vFrm As FrmTrans_ManufacturingOrder_ProcessOrder)
        Dim rowProsesOrder As DseTrans_ManufacturingOrder.tbltrans_mo_processorderRow
        Dim rowDgv As DataGridViewRow
        Dim Dummy As Boolean = False
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProd As New Core_Module.clsMasterProduct
        Dim ObjProses As New Core_Module.ClsMasterProses
        Dim Count As Integer = 0

        Try
            'salin data bom
            vDgvProsesOrder.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvProsesOrder.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            vDgvProsesOrder.Rows.Clear()
            Dummy = vDgvProsesOrder.AllowUserToAddRows
            vDgvProsesOrder.AllowUserToAddRows = False

            vTblProsesOrder.DefaultView.RowFilter = "[identifier jo] = '" & vIdentifierJO & "'"
            If vTblProsesOrder.DefaultView.Count > 0 Then
                For Each dv As Data.DataRowView In vTblProsesOrder.DefaultView
                    rowProsesOrder = dv.Row
                    Count += 1
                    vDgvProsesOrder.Rows.Add()
                    rowDgv = vDgvProsesOrder.Rows(vDgvProsesOrder.Rows.Count - 1)
                    With rowProsesOrder
                        rowDgv.Cells(vFrm.colNo.Index).Value = rowDgv.Index + 1
                        rowDgv.Cells(vFrm.colPrimary.Index).Value = .primary_process
                        rowDgv.Cells(vFrm.colPrimaryProduct.Index).Value = .primary_product
                        rowDgv.Cells(vFrm.colProcessCode.Index).Value = ObjProses.GetCode(.primary_process)
                        rowDgv.Cells(vFrm.colProcessName.Index).Value = ObjProses.GetDesc(.primary_process)
                        rowDgv.Cells(vFrm.colUrutan.Index).Value = ""
                    End With
                Next
            End If

            vDgvProsesOrder.AllowUserToAddRows = Dummy

            vDgvProsesOrder.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvProsesOrder.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
            ObjProd = Nothing
            ObjProses = Nothing
        End Try
    End Sub

    Public Sub IsiDgvItemBundling(ByVal vIdentifierWO As String, ByVal vDgvBundling As DataGridView, ByVal vFrm As FrmTrans_ManufacturingOrder_WO, ByVal vIdentifierJO As String)
        Dim Dta As New DseTrans_ManufacturingOrderTableAdapters.tbltransaksi_stockbundlingTableAdapter
        Dim Tbl As New DseTrans_ManufacturingOrder.tbltransaksi_stockbundlingDataTable
        Dim Row As DseTrans_ManufacturingOrder.tbltransaksi_stockbundlingRow
        Dim RowD As DataGridViewRow

        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            vDgvBundling.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDgvBundling.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            Dta.FillBy_IdentifierWO(Tbl, vIdentifierWO, vIdentifierJO)

            vDgvBundling.Rows.Clear()
            For Each Row In Tbl.Rows
                vDgvBundling.Rows.Add()
                RowD = vDgvBundling.Rows(vDgvBundling.Rows.Count - 1)
                With Row
                    RowD.Cells(vFrm.dgvBundling_No.Index).Value = RowD.Index + 1
                    RowD.Cells(vFrm.dgvBundling_DocNo.Index).Value = .No_Transaksi
                    RowD.Cells(vFrm.dgvBundling_WorkUnitCode.Index).Value = .work_unit_code
                    RowD.Cells(vFrm.dgvBundling_StartDate.Index).Value = Format(.start_time_actual, "dd MMM yyyy HH:mm")
                    RowD.Cells(vFrm.dgvBundling_EndDate.Index).Value = Format(.end_time_actual, "dd MMM yyyy HH:mm")
                    RowD.Cells(vFrm.dgvBundling_Remarks.Index).Value = .Remarks
                    RowD.Cells(vFrm.dgvBundling_PrimaryBundling.Index).Value = .Primary
                End With
            Next

            vDgvBundling.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDgvBundling.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Dta.Dispose()
            Tbl.Dispose()
        End Try
    End Sub

    Public Sub UpdateQtySpare(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryRM As Double, ByVal vQtySpare As Decimal)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetSpareQty(vQtySpare, vIdentifierJO, vIdentifierWO, vPrimaryRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateRequiredDate(ByVal vIdentifierJO As String, ByVal vIdentifierWO As String, ByVal vPrimaryRM As Double, ByVal vRequiredDate As Date)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetRequiredDate(vRequiredDate, vIdentifierJO, vIdentifierWO, vPrimaryRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub UpdateRequiredDate(ByVal vIdentifierJO As String, ByVal vPrimaryRM As Double, ByVal vRequiredDate As Date)
        Dim DTA As New DseTrans_ManufacturingOrderTableAdapters.tbltrans_mo_item_rmTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetRequiredDate2(vRequiredDate, vIdentifierJO, vPrimaryRM)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
