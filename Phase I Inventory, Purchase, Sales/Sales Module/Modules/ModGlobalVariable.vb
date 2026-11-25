Public Module ModGlobalVariable
    Public Function GetConstring_Sales_Module() As String
        Try
            Return My.Settings.core_dbConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

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
        ByNamaSupplier
        ByKodeCustomer
        ByNamaCustomer
        ByProductCodeNo
        ByProductDesc
        ByProductSpec
        ByCustomerInquiryNo
        ByCustomerPO_NO
        BySalesOrderNo
        TaxInvoiceNo
        ByDeliveryOrderNo
        ByNoGRN
        ByNoStockInPurchase
        ByNoStockInSalesReturn
        BySalesPersonCode
        BySalesPersonName
        ByRemarks
        ByCustomerShiptoCode
        ByCustomerShiptoName
    End Enum

    Public Sub SetConnectionString(ByVal vDT As DseLaporanSalesAnalysisTableAdapters.TblReturnInvoiceTableAdapter)
        Try
            vDT.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_SalesModule(ByVal vNewConnString As String)
        Try
            My.Settings.Item("core_dbConnectionString") = vNewConnString
            My.Settings.Save()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Module
