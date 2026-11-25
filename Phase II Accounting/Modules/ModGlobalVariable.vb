Public Module ModGlobalVariable
    Public Function GetConstring_AccModule() As String
        Try
            Return My.Settings.core_dbConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub SetConstring_AccModule(ByVal vNewConnString As String)
        Try
            My.Settings.Item("core_dbConnectionString") = vNewConnString
            My.Settings.Save()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
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
        ByCode
        ByDescription
        ByAccountNo
        ByProductSpec
        ByAOPNumber
        BySupplierCode
        BySupplierName
        ByInvoiceNumber
        ByCustomerName
        ByCustomerCode
        ByAOSRNumber
        ByAOPRNumber
        ByReferenceCode
        ByDate
        ByNoTransaksi
        TaxInvoiceNo
        ByTanggalTransaksi
        ByProductCodeNo
        ByProductDesc
        ByDeliveryOrderNo
        ByName
        ByRemarks
        ByKodeSupplier
        ByNoGRN
        BySalesOrderNo
        ByNoStockInSalesReturn
        ByNoStockInPurchase
        BySupplierInvoiceNo
        ByPurchaseReturnNo
        BySalesReturnNo
        ByKodeCustomer
        ByNamaCustomer
        ByCustomerPO_NO
        BySalesPersonCode
        BySalesPersonName
        By_IPO_Number
        ByReferenceName
        ByDescJurnal
    End Enum

    'Public Server As String = "localhost"
    'Public UserId As String = "root"
    'Public Password As String = "123"
    'Public DatabaseName As String = "core_db"
    'Public ConnectTimeOut As Integer = 300
    'Public CommandTimeOut As Integer = 300
    'Public Port As String = "3307"

    'Public Function GetConstring() As String
    '    Try
    '        Return ""
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    End Try
    'End Function
End Module
