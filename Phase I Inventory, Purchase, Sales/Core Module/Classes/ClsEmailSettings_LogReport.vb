Public Class ClsEmailSettings_LogReport
    Public Function LoadEmailAccounts(ByVal vTbl As DseEmailSettings.tbltrans_emailsettingsDataTable) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.tbltrans_emailsettingsTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.Fill(vTbl) > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ShowLog(ByVal vTbl As DseEmailSettings.DataTable3DataTable, ByVal vDateFrom As Date, ByVal vDateTo As Date) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.DataTable3TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.FillBy_Date(vTbl, vDateFrom, vDateTo) > 0 Then
                Return True
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ShowLog(ByVal vTbl As DseEmailSettings.DataTable3DataTable, ByVal vDateFrom As Date, ByVal vDateTo As Date, ByVal vSent As Boolean) As Boolean
        Dim dta As New DseEmailSettingsTableAdapters.DataTable3TableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If vSent = True Then
                If dta.FillBy_DateAndStatus(vTbl, vDateFrom, vDateTo, 1) > 0 Then
                    Return True
                End If
            Else
                If dta.FillBy_DateAndStatus(vTbl, vDateFrom, vDateTo, 0) > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function
End Class
