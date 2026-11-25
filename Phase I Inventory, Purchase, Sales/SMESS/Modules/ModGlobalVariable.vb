Public Module ModGlobalVariable
    Dim vEncryptPassInstaller As String = "*#%Simple83_999%$***##@@"
    Dim vEncryptPassInstaller2 As String = "%*_C1tySquareB-12%#"

    Public Function GetEncryptPass_Installer() As String
        Return vEncryptPassInstaller
    End Function
    Public Function GetEncryptPass_Installer2() As String
        Return vEncryptPassInstaller2
    End Function

    Public Function GetConstring_SMESS_Module() As String
        Try
            Return My.Settings.core_dbConnectionString
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function
End Module
