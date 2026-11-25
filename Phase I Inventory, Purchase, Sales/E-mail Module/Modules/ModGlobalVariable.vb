Public Module ModGlobalVariable
    Public Sub SetLicense(ByVal vLicenseKey As String)
        Try
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

    Public Function GetConstring_EmailModule() As String
        Try
            Return My.Settings.core_dbConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub SetConstring_DTAEmail(ByVal DTAEmail As DsTransaksiEmailTableAdapters.tbltrans_email_automated_sendtoTableAdapter)
        Try
            DTAEmail.Connection.ConnectionString = Core_Module.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_EmailModule(ByVal vNewConnString As String)
        Try
            My.Settings.Item("core_dbConnectionString") = vNewConnString
            My.Settings.Save()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Module
