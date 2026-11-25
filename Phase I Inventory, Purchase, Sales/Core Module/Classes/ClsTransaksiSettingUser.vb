Option Explicit On
Option Strict Off

Public Class ClsTransaksiSettingUser
    Public Enum EnumTipeTransaksi As Long
        ShippingForecast
        CollectionLetter
        MaterialPlanningGrouping
        MonthlyInventoryClosing
        HR_Attendance
        PurchaseOrder
        POC_Quick
        MaterialPlanning
        SalesOrder
        MasterCustomer_Obs
        Transaksi_AOP_CreateFromStockIn
        MasterCustomer
        MasterProduct
        MasterSupplier
        FrmTrans_ProductionSchedule_NewSch
        FrmTrans_ProductionSchedule_NewSch2
        FrmTrans_ProductionSchedule
        ProductionSetting
        FrmTrans_ProductionSchedule_NewSch2_Detail
        FrmTrans_ManufacturingOrder
        FrmTrans_ProductionSchedule_ChangeWU
        FrmTrans_Journal_View
        FrmTrans_StockBundling_Quick
        FrmTrans_ProductionSchedule_DetailWO
        MasterKaryawan
        HR_OvertimeOrder 'trf source infiniti to finalize
        HR_Salary 'trf source infiniti to finalize
        FrmLaporanStockTransfer
        SOC_Quick
        BAGUS_PrinterName
        BAGUS_PrinterName_QuickStockOutSales
        PrinterName_QuickInvoiceByStockOutSales
        UseCustomPrint_QuickInvoiceByStockOutSales
        FrmTransQuickOutSales_Filteran
        HistoricalPrice
        SettingPrint_Invoice
        SettingPrint_PO
        SettingPrint_PR
        SettingPrint_CustomProdList
        SettingPrint_CustomInvoiceList
    End Enum

    Public Enum EnumNamaParameter As Long
        ShippingForecast_ShowSONumber
        ShippingForecast_ShowCustPO
        ShippingForecast_ShowProductDesc
        ShippingForecast_PrintOrderNo
        ShippingForecast_PrintGroupBy
        CollectionLetter_IsSendEmail
        CollectionLetter_LastEmailUsed
        MaterialPlanningGrouping_ShowDetail
        MonthlyInventoryClosing_ShowSupplierInvoiceNumber
        MonthlyInventoryClosing_ShowDetailBegining
        MonthlyInventoryClosing_ShowReferenceParty
        MonthlyInventoryClosing_ShowProductDetail
        MonthlyInventoryClosing_ShowFormatDetailAkun
        HR_Attendance_ShowSchedule
        HR_Attendance_ShowLate
        HR_Attendance_ShowEarlyOut
        HR_Attendance_ShowActualBreak
        HR_Attendance_ShowOT
        HR_Attendance_ShowEartlyOT
        PurchaseOrder_ShowQtyDefault
        PurchaseOrder_ShowUnitCodeDefault
        PurchaseOrder_ShowToleranceDown
        PurchaseOrder_ShowToleranceUp
        POC_Quick_SupplierFormat
        MaterialPlanning_ShowProcurementStrategy
        SalesOrder_ShowUnitsDefault
        SalesOrder_ShowTolerance
        HR_Attendance_ShowCalculateEarlyOT
        HR_Attendance_ShowCalculateLate
        HR_Attendance_ShowCalculateEarlyOut
        HR_Attendance_ShowCalculateOT
        MasterCustomer_DefaultFindBy_Obs
        MaterialPlanningGrouping_PrOBLargerThanPOB
        Transaksi_AOP_CreateFromStockIn_ShowSupplierMode
        MasterCustomer_DefaultFindBy
        MasterProduct_DefaultFindBy
        MasterSupplier_DefaultFindBy
        MaterialPlanning_UseWarehouseExclusion
        HR_Attendance_ShowOutstation
        FrmTrans_ProductionSchedule_NewSch2_ShowWO
        FrmTrans_ProductionSchedule_NewSch2_ShowCodeNo
        FrmTrans_ProductionSchedule_NewSch2_ShowDesc
        FrmTrans_ProductionSchedule_NewSch2_ShowDayStart
        FrmTrans_ProductionSchedule_NewSch2_ShowDayFinish
        FrmTrans_ProductionSchedule_ShowWO
        FrmTrans_ProductionSchedule_ShowCodeNo
        FrmTrans_ProductionSchedule_ShowDesc
        FrmTrans_ProductionSchedule_ShowDayStart
        FrmTrans_ProductionSchedule_ShowDayFinish
        FrmTrans_ProductionSchedule_NewSch_ShowJO
        FrmTrans_ProductionSchedule_NewSch_ShowMO
        FrmTrans_ProductionSchedule_NewSch_ShowCodeNo
        FrmTrans_ProductionSchedule_NewSch_ShowDesc
        FrmTrans_ProductionSchedule_NewSch_ShowQtyTarget
        ProductionSetting_ShowFreeTime
        ProductionSetting_ShowOffTime
        FrmTrans_ProductionSchedule_ShowProcess
        FrmTrans_ProductionSchedule_NewSch2_Detail_ShowJO
        FrmTrans_ProductionSchedule_NewSch2_Detail_ShowProcess
        FrmTrans_ProductionSchedule_NewSch2_Detail_ShowCodeNo
        FrmTrans_ProductionSchedule_NewSch2_Detail_ShowDesc
        FrmTrans_ProductionSchedule_ShowProgress
        FrmTrans_ProductionSchedule_ShowRate
        FrmTrans_ProductionSchedule_ShowQty
        FrmTrans_ProductionSchedule_ShowUnit
        FrmTrans_ProductionSchedule_ShowWH
        FrmTrans_ManufacturingOrder_ShowCodeNo
        FrmTrans_ManufacturingOrder_ShowDescription
        FrmTrans_ManufacturingOrder_ShowProcessCode
        FrmTrans_ManufacturingOrder_ShowProcessDesc
        FrmTrans_ProductionSchedule_NewSch2_ShowProcess
        FrmTrans_ProductionSchedule_NewSch2_ShowRate
        FrmTrans_ProductionSchedule_NewSch2_ShowQty
        FrmTrans_ProductionSchedule_NewSch2_ShowUnit
        FrmTrans_ProductionSchedule_NewSch2_ShowWH
        MaterialPlanning_ProcurementType
        MaterialPlanning_ShowWONo
        FrmTrans_ManufacturingOrder_RMShowDetail
        FrmTrans_ManufacturingOrder_RMIncludeWIP
        FrmTrans_ManufacturingOrder_WO_UseTreeView
        FrmTrans_ProductionSchedule_ShowInputLocation
        FrmTrans_ProductionSchedule_ShowOutputLocation
        FrmTrans_ProductionSchedule_NewSch2_ShowInputLocation
        FrmTrans_ProductionSchedule_NewSch2_ShowOutputLocation
        FrmTrans_ProductionSchedule_ShowActualGoodQty
        FrmTrans_ProductionSchedule_ShowActualRejectQty
        FrmTrans_ProductionSchedule_ShowActualProduceQty
        HR_Attendance_MergeAttendance
        FrmTrans_ProductionSchedule_ChangeWU_ShowMO
        FrmTrans_ProductionSchedule_ChangeWU_ShowJO
        FrmTrans_ProductionSchedule_ChangeWU_ShowWO
        FrmTrans_ProductionSchedule_ChangeWU_ShowProcess
        FrmTrans_ProductionSchedule_ChangeWU_ShowCodeNo
        FrmTrans_ProductionSchedule_ChangeWU_ShowDesc
        FrmTrans_ProductionSchedule_ChangeWU_ShowRate
        FrmTrans_ProductionSchedule_ChangeWU_ShowQty
        FrmTrans_ProductionSchedule_ChangeWU_ShowUnit
        FrmTrans_ProductionSchedule_ChangeWU_ShowWH
        FrmTrans_ProductionSchedule_ChangeWU_ShowStartDay
        FrmTrans_ProductionSchedule_ChangeWU_ShowEndDay
        FrmTrans_ProductionSchedule_ChangeWU_ShowProgress
        FrmTrans_ManufacturingOrder_WIP_ShowWO
        FrmTrans_ManufacturingOrder_WIP_ShowProcessCode
        FrmTrans_ManufacturingOrder_WIP_ShowCodeNo
        FrmTrans_ManufacturingOrder_WIP_ShowDesc
        FrmTrans_ProductionSchedule_Reorder_ShowMO
        FrmTrans_ProductionSchedule_Reorder_ShowJO
        FrmTrans_ProductionSchedule_Reorder_ShowWO
        FrmTrans_ProductionSchedule_Reorder_ShowProcess
        FrmTrans_ProductionSchedule_Reorder_ShowCodeNo
        FrmTrans_ProductionSchedule_Reorder_ShowDesc
        FrmTrans_ProductionSchedule_Reorder_ShowRate
        FrmTrans_ProductionSchedule_Reorder_ShowQty
        FrmTrans_ProductionSchedule_Reorder_ShowUnit
        FrmTrans_ProductionSchedule_Reorder_ShowWH
        FrmTrans_ProductionSchedule_Reorder_ShowDayStart
        FrmTrans_ProductionSchedule_Reorder_ShowDayFinish
        FrmTrans_ProductionSchedule_Reorder_ShowProgress
        FrmTrans_Journal_View_Format
        FrmTrans_Journal_View_ShowRemarks
        FrmTrans_ProductionSchedule_StartDateDiff
        FrmTrans_StockBundling_Quick_ShowCodeNo
        FrmTrans_StockBundling_Quick_ShowDesc
        FrmTrans_StockBundling_Quick_ShowProcessCode
        FrmTrans_StockBundling_Quick_ShowProcessDesc
        FrmTrans_StockBundling_Quick_ShowStartTime
        FrmTrans_StockBundling_Quick_ShowEndTime
        FrmTrans_ProductionSchedule_DetailWO_ShowCodeNo
        FrmTrans_ProductionSchedule_DetailWO_ShowDesc
        FrmTrans_ProductionSchedule_DetailWO_ShowProcess
        FrmTrans_ProductionSchedule_DetailWO_ShowProcessDesc
        MasterKaryawan_DefaultFindBy
        HR_OvertimeOrder_AppliedBy 'trf source infiniti to finalize
        HR_OvertimeOrder_AppliedPosition 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowEarlyOT 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowOT 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowHolidayOT 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowEarlyDelay 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowDelay 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowHolidayDelay 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowEarlyBreak 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowBreak 'trf source infiniti to finalize
        HR_OvertimeOrder_ShowHolidayBreak 'trf source infiniti to finalize
        HR_Attendance_Quick_Type_Filter 'trf source infiniti to finalize
        HR_Salary_Quick_Type_Filter 'trf source infiniti to finalize
        HR_ClockInOut_ShowSchedule 'trf source infiniti to finalize
        HR_ClockInOut_ShowScheduleBreak 'trf source infiniti to finalize
        HR_ClockInOut_ShowActual 'trf source infiniti to finalize
        HR_ClockInOut_ShowActualBreak 'trf source infiniti to finalize
        HR_ClockInOut_ShowShift 'trf source infiniti to finalize
        HR_ClockInOut_ShowDept 'trf source infiniti to finalize
        MaterialPlanning_SupplierSelectionMode
        FrmLaporanStockTransfer_FilterByWarehouse
        SOC_Quick_CustomerFormat
        PrinterName
        FrmTransQuickOutSales_FilterIsWarehouse
        FrmTransQuickOutSales_FilterPrimary
        FrmTransQuickOutSales_FilterUseCustomPrinter
        FrmTransQuickOutSales_FilterPrinterName
        FrmTransQuickOutSales_FilterIsDirectlySave
        FrmTransQuickOutSales_FilterIsDirectlyPrint
        SalesInvQuickByStockOut_FilterUseCustomPrinter
        SalesInvQuickByStockOut_FilterPrinterName
        SalesInvQuickByStockOut_FilterDirectlySave
        SalesInvQuickByStockOut_FilterDirectlyPrint
        SettingHistoricalPrice_QuickInvoice
        SettingHistoricalPrice_TglDari
        SettingHistoricalPrice_TglSampai
    End Enum

    Public Function GetSettingValue(ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vNamaParameter As EnumNamaParameter, ByVal vUserName As String) As String
        Dim DTA As New DseTransaksiSettingUserTableAdapters.tbltrans_setting_userTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetValue(vTipeTransaksi, vNamaParameter, vUserName)) = False Then
                Return DTA.ScalarQuery_GetValue(vTipeTransaksi, vNamaParameter, vUserName)
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

    Public Sub SetSettingValue(ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vNamaParameter As EnumNamaParameter, ByVal vNilaiParameter As String, ByVal vUserName As String)
        Dim DTA As New DseTransaksiSettingUserTableAdapters.tbltrans_setting_userTableAdapter
        Dim Insert As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vTipeTransaksi, vNamaParameter, vUserName).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vTipeTransaksi, vNamaParameter, vUserName).Value = 0 Then
                    Insert = True
                Else
                    Insert = False
                End If
            Else
                Insert = True
            End If

            If Insert = True Then
                DTA.Insert(vTipeTransaksi, vNamaParameter, vNilaiParameter, vUserName)
            Else
                DTA.Update(vNilaiParameter, vTipeTransaksi, vNamaParameter, vUserName)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
