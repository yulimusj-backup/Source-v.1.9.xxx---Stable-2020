Public Module ModGlobalVariable
    Public Function GetConstring_Purchase_Module() As String
        Try
            Return My.Settings.core_dbConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub SetConstring_Purchase_Module(ByVal vNewConnString As String)
        Try
            My.Settings.Item("core_dbConnectionString") = vNewConnString
            My.Settings.Save()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetCulture()
        My.Application.Update1()
    End Sub

    Public Sub UpdateSettingsPO()
        Inventory_Module.Set_Setting_PO_PreparedBy(My.Settings.PO_Prepared_By)
        Inventory_Module.Set_Setting_PO_PreparedPosition(My.Settings.PO_Prepared_Position)
        Inventory_Module.Set_Setting_PO_AcknowledgeBy(My.Settings.PO_Acknowled_By)
        Inventory_Module.Set_Setting_PO_AcknowledgePosition(My.Settings.PO_Acknowledge_Position)
        Inventory_Module.Set_Setting_PO_ApprovedBy(My.Settings.PO_Approved_By)
        Inventory_Module.Set_Setting_PO_ApprovedPosition(My.Settings.PO_Approved_Position)
    End Sub

    Public Sub SetLicense(ByVal vLicenseKey As String)
        Try
            My.Settings.b = vLicenseKey
            My.Settings.Save()

            Dim ij As New System.Globalization.DateTimeFormatInfo
            Dim jk As New System.Globalization.NumberFormatInfo

            'cek ui culture dan keyboard input 
            My.Application.ChangeCulture("en-US")
            My.Application.ChangeUICulture("en-US")

            My.Application.ChangeCulture("en-US")
            My.Application.ChangeUICulture("en-US")
            ij.LongDatePattern = "dddd, MMM dd yyyy"
            ij.ShortDatePattern = "MMM dd yyyy"
            jk.CurrencyDecimalSeparator = "."
            jk.CurrencyGroupSeparator = ","
            My.Application.Culture.NumberFormat = jk
            My.Application.Culture.DateTimeFormat = ij

            My.Application.Update1()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumSearchCriteria As Integer
        ByNoTransaksi
        ByTanggalTransaksi
        ByKodeSupplier
        ByKodeCustomer
        ByProductCodeNo
        ByProductDesc
        ByProductSpec
        BySupplierInvoiceNo
        BySupplierName
        BySupplierDONo
        ByPONo
        BySalesOrderNo
        ByCustomerPO_NO
        ByNamaCustomer
        GIT_Number
        SupplierPackingListNo
        ByNoGRN
        ByNoStockInPurchase
        ByStockInNo
        By_IPO_Number
        ByAccountNo
        ByDescription
        ByPRNumber
        ByReason
        ByRemarks
        ByRequestBy
        ByNoStockOutAdjust
        ByStockInDate
    End Enum
End Module
