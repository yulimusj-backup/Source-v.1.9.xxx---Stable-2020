Option Explicit On
Option Strict Off

Public Class ClsTransaksiSetting
    Public Enum EnumTipeTransaksi As Long
        SalesJurnal
        PurchaseJurnal
        Invoice
        AOP
        DeliveryOrder
        PO
        PurchaseReturn
        Reporting
        Jurnal
        HR_Attendance
        OT_Duration
        Early_OT_Duration
        Late_Duration
        SO
        SuratTagihan
        MaterialPlanning
        SalaryCalculation
        Stock_In_Purchase
        Stock_In_SalesReturn
        Stock_In_Adjust
        Stock_Out_Sales
        Stock_Out_Purchase_Return
        Stock_Out_Adjust
        Stock_Transfer
        Stock_Bundling
        Purchase_Request
        Purchase_Request_PickUp_Order
        SalesQuotation
        MasterCustomer
        MasterProduct
        UpdateFromLAN_RequireUAC
        MasterSupplier
        HR_Salary
        SalesReturn
        ProductionSetting
        Overtime_Order
        Location_Privileges
        ManufacturingOrder
        MasterKaryawan
        HR_OvertimeOrder 'trf source infiniti to finalize
        VAT
        NPWP
        PackingList
    End Enum

    Public Enum EnumNamaParameter As Long
        Akun_AR
        Akun_AP
        UseTax
        TaxPercentage
        PrintUseCustomDO_Bool
        CustomDO_Str
        PrintUseCustomInvoice_Bool
        CustomInvoiceAddress_Str
        PrintUseCustomTaxInvoice_Bool
        CustomTaxInvoice_Str
        Use_Custom_Print_Logic
        SignedByName
        SignedByJabatan
        ShowRemarks
        UseBrackets
        CustomPrint
        Akun_CurrenctYearEarning
        Akun_VAT_In
        Akun_VAT_Out
        Akun_Dutystamp_Expense
        Akun_PaymentDifference
        Akun_ShippingIncome
        CetakJurnalPortrait
        CostingMethod
        Akun_AOP_DP
        Akun_Invoice_DP
        ShipmentTax
        GlobalDiscAcc
        QtyReminderFakturPajak
        Reminder_AccountReceivable_Overdue
        Reminder_AccountReceivable_ExceedsCreditLimit
        Reminder_AccountReceivable_Overdue_Tolerance
        Reminder_AccountReceivable_ExceedsCreditLimit_Tolerance
        IgnoreBreakData
        IgnoreEarlyIn
        IgnoreEarlyOut
        OT_Duration_RoundingType
        OT_Duration_RoundingUnit
        Early_OT_Duration_RoundingType
        Early_OT_Duration_RoundingUnit
        Late_Duration_RoundingType
        Late_Duration_RoundingUnit
        MaxValue_Early_In
        MaxValue_Early_Break_Out
        MaxValue_Late_Break_In
        MaxValue_Late
        MaxValue_Early_Out
        MaxValue_OT_Daily
        MaxValue_OT_Period
        MaxValue_Early_OT_Daily
        MaxValue_Early_OT_Period
        Break_duration
        Default_Salary_PrimaryCashAccount
        Default_Salary_PrimaryTaxPayableAccount1
        Default_Loan_PrimaryCashInAccount
        Default_Loan_PrimaryGainAccount
        Default_Loan_PrimaryLossAccount
        InvoiceIncludeCalculation
        Akun_Retained_Earning
        SuratTagihan_BuatCustom
        SuratTagihan_SignedByName
        SuratTagihan_SignedByPosition
        SuratTagihan_Remarks
        SuratTagihan_Subject
        MaterialPlanning_JmlHariKerjaPerBulan
        PreparationDelay
        Default_Salary_PrimaryTaxExpenseAccount
        Default_LoanPayment_PrimaryCashInAccount
        Default_LoanPayment_PrimaryGainAccount
        Default_LoanPayment_PrimaryLossAccount
        Akun_Credit_Note_Customer
        Akun_Credit_Note_Supplier
        AttEmail_Subject
        AttEmail_SignBy
        AttEmail_SignPosition
        Default_TimeFrom_ActualClockInOutReport
        Default_TimeTo_ActualClockInOutReport
        Default_IfMultipleClockInOut_ActualClockInOutReport
        Akun_Debit_Note_Customer
        Akun_Debit_Note_Supplier
        Use_Custom_CostFaktorCalc
        Jurnal_IsLedgerAccessActive
        Purchase_Request_DefaultUseTracking
        MasterCustomer_DefaultFindBy
        MasterProduct_DefaultFindBy
        UpdateFromLAN_RequireUAC
        MasterSupplier_DefaultFindBy
        SalaryEmail_Subject
        SalaryEmail_SignBy
        SalaryEmail_SignPosition
        SalesQuotation_UseLvl2_Quotation
        SalesQuotation_UseLvl2_Limit
        ProductionSetting_SchHor_StartDate
        ProductionSetting_SchHor_EndDate
        ProductionSetting_AllocatedWS_a
        ProductionSetting_AllocatedWS_r
        ProductionSetting_AllocatedWS_g
        ProductionSetting_AllocatedWS_b
        ProductionSetting_Offtime_a
        ProductionSetting_Offtime_r
        ProductionSetting_Offtime_g
        ProductionSetting_Offtime_b
        ProductionSetting_SchDT_a
        ProductionSetting_SchDT_r
        ProductionSetting_SchDT_g
        ProductionSetting_SchDT_b
        ProductionSetting_SchOT_a
        ProductionSetting_SchOT_r
        ProductionSetting_SchOT_g
        ProductionSetting_SchOT_b
        ProductionSetting_SchHor_FirmPlanned1
        Activate_Access_Retriction
        Default_User_Not_IncludeInList
        ProductionLogWorkShiftRequired
        ManufacturingOrder_CustomPrintMO
        ManufacturingOrder_CustomPrintJO
        ShowDialogOfProductionLog
        MO_BoM_ToolsEnabled
        DeliveryOrder_AllowNoStock
        Stock_Out_Sales_AllowOverrideDO
        DeliveryOrder_ShowDeliveryDate
        MasterKaryawan_DefaultFindBy
        Jurnal_VerifyProductAccount
        OvertimeOrderEmail_SignBy 'trf source infiniti to finalize
        OvertimeOrderEmail_SignPosition 'trf source infiniti to finalize
        OvertimeOrderEmail_Subject 'trf source infiniti to finalize
        ShowDiskon
        NPWP_StartPeriode
        NPWP_AkhirTax04 '2025 coretax tahap 2.2
    End Enum

    Public Function GetSettingValue(ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vNamaParameter As EnumNamaParameter) As String
        Dim DTA As New DseTransaksiSettingTableAdapters.tbltrans_settingTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNilaiParameter(vTipeTransaksi, vNamaParameter)) = False Then
                Return DTA.ScalarQuery_GetNilaiParameter(vTipeTransaksi, vNamaParameter)
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

    Public Sub SetSettingValue(ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal vNamaParameter As EnumNamaParameter, ByVal vNilaiParameter As String)
        Dim DTA As New DseTransaksiSettingTableAdapters.tbltrans_settingTableAdapter
        Dim Insert As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery(vTipeTransaksi, vNamaParameter).HasValue = True Then
                If DTA.ScalarQuery(vTipeTransaksi, vNamaParameter).Value = 0 Then
                    Insert = True
                Else
                    Insert = False
                End If
            Else
                Insert = True
            End If

            If Insert = True Then
                DTA.Insert(vTipeTransaksi, vNamaParameter, vNilaiParameter)
            Else
                DTA.Update(vNilaiParameter, vTipeTransaksi, vNamaParameter)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
End Class
