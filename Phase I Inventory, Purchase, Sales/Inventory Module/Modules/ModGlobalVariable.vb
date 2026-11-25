Public Module ModGlobalVariable
    Public Function GetConstring_Inventory_Module() As String
        Try
            Return My.Settings.core_dbConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub Set_Setting_PO_PreparedBy(ByVal vValue As String)
        My.Settings.PO_Prepared_By = vValue
        My.Settings.Save()
    End Sub
    Public Sub Set_Setting_PO_PreparedPosition(ByVal vValue As String)
        My.Settings.PO_Prepared_Position = vValue
        My.Settings.Save()
    End Sub
    Public Sub Set_Setting_PO_AcknowledgeBy(ByVal vValue As String)
        My.Settings.PO_Acknowled_By = vValue
        My.Settings.Save()
    End Sub
    Public Sub Set_Setting_PO_AcknowledgePosition(ByVal vValue As String)
        My.Settings.PO_Acknowledge_Position = vValue
        My.Settings.Save()
    End Sub
    Public Sub Set_Setting_PO_ApprovedBy(ByVal vValue As String)
        My.Settings.PO_Approved_By = vValue
        My.Settings.Save()
    End Sub
    Public Sub Set_Setting_PO_ApprovedPosition(ByVal vValue As String)
        My.Settings.PO_Approved_Position = vValue
        My.Settings.Save()
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
        ByNoWorkOrder
        ByTakingTeamNo
        ByProductSpec
        BySalesOrderNo
        ByNamaCustomer
        ByCustomerPO_NO
        ByNamaSupplier
        ByNoGRN
        ByNoStockInPurchase
        ByDeliveryOrderNo
        ByPONumber
        BySupplier_s_DONumber
        ByRemarks
        BySupplierName
        By_IPO_Number
        ByDescription
        ByProductGroupCode
        ByProductGroupDesc
        ByBrandCode
        ByBrandDesc
        BySubBrandCode
        BySubBrandDesc
        ByCategoryCode
        ByCategoryDesc
        ByWorkUnitCode
        ByProsesCode
        ByWorkShift
        ByRequestBy
        ByNoStockOutAdjust
        ByNoDokumen
        ByJobdesc
        ByPurchaseReturnNo
        ByJONumber
        ByCustomerShiptoCode
        ByCustomerShiptoName
    End Enum

    Public Enum EnumSreachCriteria_Material_Usage As Integer
        NoTransaksi
        _Date
        AssetCode
        AssetDescrition
        Remarks
    End Enum

    Public Sub UpdateSettingPO(ByVal PO_Prepared_By As String, ByVal PO_Prepared_Position As String, ByVal PO_Acknowled_By As String, _
    ByVal PO_Acknowledge_Position As String, ByVal PO_Approved_By As String, ByVal PO_Approved_Position As String)
        Try
            My.Settings.PO_Prepared_By = PO_Prepared_By
            My.Settings.PO_Prepared_Position = PO_Prepared_Position
            My.Settings.PO_Acknowled_By = PO_Acknowled_By
            My.Settings.PO_Acknowledge_Position = PO_Acknowledge_Position
            My.Settings.PO_Approved_By = PO_Approved_By
            My.Settings.PO_Approved_Position = PO_Approved_Position

            My.Settings.Save()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub SetConnectionString(ByVal vDT As dseTransaksiStockInOutTableAdapters.transaksi_stockinout_item_viewTableAdapter)
        Try
            vDT.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub SetConnectionString(ByVal vDT As dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_after1TableAdapter)
        Try
            vDT.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
    Public Sub SetConnectionString(ByVal vDT As dseTransaksiStockBundlingTableAdapters.tbltransaksi_stockbundling_item_before4TableAdapter)
        Try
            vDT.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_InventoryModule(ByVal vNewConnString As String)
        Try
            My.Settings.Item("core_dbConnectionString") = vNewConnString
            My.Settings.Save()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Module
