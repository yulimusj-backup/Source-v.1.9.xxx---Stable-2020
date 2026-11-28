Public Module ModGlobalVariable
    Public DefPrecExchangeRate As Integer = 4
    Public PrimaryBaseCurrency As Double = 1
    Dim vEncryptPass As String = "*#%SiMelekete83-wkwkwkwk"
    Dim vEncryptPassInstaller As String = "*#%Simple83_999%$***##@@"
    Dim vEncryptPassInstaller2 As String = "%*_C1tySquareB-12%#"

    Public Function GetEncryptPass_Installer() As String
        Return vEncryptPassInstaller
    End Function
    Public Function GetEncryptPass_Installer2() As String
        Return vEncryptPassInstaller2
    End Function
    Public Function GetEncryptPass() As String
        Return vEncryptPass
    End Function
    Public Function GetDBServer() As String
        Return My.Settings.server
    End Function
    Public Function GetDBUserID() As String
        Return My.Settings.user_id
    End Function
    Public Function GetDBPassword() As String
        Return My.Settings.password
    End Function
    Public Function GetDBName() As String
        Return My.Settings.database_name
    End Function
    Public Function GetDBPort() As String
        Return My.Settings.port_number
    End Function

    Public Function GetSMESSBackColorAll() As Color
        Return Color.FromArgb(0, 231, 249)
    End Function

    Public Function GetSMESSColor() As Color
        Return Color.FromArgb(0, 193, 208)
    End Function

    Public Function GetSMESSColor_Border() As Color
        Return Color.FromArgb(0, 86, 94)
    End Function

    Public Sub SetConstring_DTAMasterProduct_ItemMU(ByVal DTAMasterProduct_ItemMU As Core_Module.dseMasterProductTableAdapters.tblmaster_product_item_multiunitTableAdapter)
        Try
            DTAMasterProduct_ItemMU.Connection.ConnectionString = GetConstring()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_DTAMasterProduct(ByVal DTAMasterProduct As Core_Module.dseMasterProductTableAdapters.tblmaster_product1TableAdapter)
        Try
            DTAMasterProduct.Connection.ConnectionString = GetConstring()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_DTALocationPriority(ByVal DTALocationPriority As Core_Module.DseMasterLocationPriorityTableAdapters.tblmaster_location_priority1_TableAdapter)
        Try
            DTALocationPriority.Connection.ConnectionString = GetConstring()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_DTAAntrian(ByVal DTAAntri As Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter)
        Try
            DTAAntri.Connection.ConnectionString = GetConstring()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_DTAKary(ByVal DTAKary As Core_Module.DseMasterKaryawanTableAdapters.tblmaster_karyawan3TableAdapter)
        Try
            DTAKary.Connection.ConnectionString = GetConstring()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub SetConstring_DTAProductQty(ByVal DTAProductQty As dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter)
        Try
            DTAProductQty.Connection.ConnectionString = GetConstring()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetConstring_CoreModule1() As String
        Try
            Return My.Settings.Core_DbConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetConstring_CoreModule2() As String
        Try
            Return My.Settings.core_dbConnectionString1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ";default command timeout=19000000"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=fitri-pc;user id=root;Password=123;persist security info=True;database=payroll_db;port=3307;default command timeout=19000000"
        End Try
    End Function

    Public Function GetConstring_withDB(ByVal vDatabase As String) As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + vDatabase + ";port=" + My.Settings.port_number + ";default command timeout=19000000"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=fitri-pc;user id=root;Password=123;persist security info=True;database=payroll_db;port=3307;default command timeout=19000000"
        End Try
    End Function

    'berkaitan cut off
    Public Sub SetConstring(ByVal vNewConnString As String)
        Dim connection_ As New MySql.Data.MySqlClient.MySqlConnection
        Try
            My.Settings.Item("Core_DbConnectionString") = vNewConnString
            My.Settings.Item("core_dbConnectionString1") = vNewConnString
            connection_.ConnectionString = My.Settings.Core_DbConnectionString
            My.Settings.Item("database_name") = connection_.Database
            'My.Settings.database_name = connection_.Database
            My.Settings.Save()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Enum EnumBahasa As Integer
        Inggris
        Indonesia
    End Enum

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

    Public Enum EnumSearchCriteriaAsset As Integer
        ByAssetCode
        ByDescription
        ByRemark
        ByGroupCode
    End Enum

    Public Enum EnumSearchCriteria As Integer
        ByCode
        ByDescription
        ByAccountNo
        ByProductSpec
        ByNoTransaksi
        ByBarcode
        ByRemarks
        ByCompetenceCode
        ByCompetenceDesc
        ByEmployeeCode
        ByEmployeeName
        ByDate
        BySalaryNo
        ByFullName
        ByLoanNumber
        ByGroupCode
        ByGroupDesc
        BySalesPCode
        BySalesPName
        ByComponentName
        BySalaryFormula
        ByGroup
        ByGroupProductGroup
        ByGroupBrand
        ByGroupSubBrand
        ByGroupCategory
        ByMachineGroup
        ByEmployee
        ByWorkShift
        ByMachine
        ByEmail
        ByTanggalTransaksi
        ByRequestBy
        ByReason
        ByWONumber
        ByCodeLocationIn
        ByCodeLocationOut
        ByDate2
        ByRefDoc
        ByCodeCustom
        ByDescriptionCustom
        ByTaxUnitCode
        ByTaxUnitDesc
        ByTaxProductCode
        ByTaxProductDesc
    End Enum

    Public Sub SetUser(ByVal vUserCode As String)
        My.Settings.user_code = vUserCode
        My.Settings.Save()
    End Sub

    Public Function GetUser() As String
        Return My.Settings.user_code 
    End Function

    Public Function GetDownloadLink_ClientUpdateModule_Version() As String
        Return "https://www.dropbox.com/s/6v0hshaa87wwzkm/Client_Update_Version.smess?dl=1"
    End Function

    Public Function GetDownloadLink_ClientUpdateModule() As String
        Return "https://www.dropbox.com/s/5a4igegrmqlrnvv/Client_Update.smess?dl=1"
    End Function
End Module
