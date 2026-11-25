Option Explicit On
Option Strict Off

Public Class Clstrans_VersionControl
    Dim ObjEncrypt As New Core_Module.ClsEncrypt("?qt$UG2#@&=eBHRF%Aw6VL$2L7^GRQ@D7%@EsRz=GM@yvyUM%@6=NhdX?@-#jVRAD5QFw8Va6g_WuS-wWX8eq7Wj$?hWhEFVfX5qe-d*Ud$p?DydxKY--VnrQAmb_X5PtwnJ@gHd6rj#m&!*tc-5SUmgsW7GtrqaD_Q$v%e*#q2uG%xzwe^Z#Zu-%h$jqx=H+by=K8?#VqH@T9FTFWze#5VHjkuX#F_ht@VA%!6nBn?^xcEJ@_2D@wZkMpvD*$du")

    Public Function GetExpiry(ByVal vExpiry_string As String, ByVal vProductKey As String) As Date
        Dim RetVal As String

        Try
            If vExpiry_string = "-" Then
                Return #1/1/1950#
            End If
            RetVal = Me.ObjEncrypt.Decrypt(vExpiry_string).Replace("-" + vProductKey, "")
            If IsDate(RetVal) = False Then
                Return #1/1/1950#
            Else
                Return CDate(RetVal)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        End Try
    End Function

    Private Function GetExpiry_String(ByVal vExpiry As Date, ByVal vProductKey As String) As String
        Try
            Return Me.ObjEncrypt.Encrypt(Format(vExpiry, "dd MMM yyyy") + "-" + vProductKey)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "-"
        End Try
    End Function

    Public Sub SetLoginTimeStamp(ByVal vProductKey As String, ByVal vUserCode As String, ByVal vCompName As String, ByVal vLicenseExpiryDate As Date)
        Dim DTA As New DseTrans_Version_ControlTableAdapters.tbltrans_version_controlTableAdapter
        Dim ObjMySQL As New Core_Module.ClsMySQLReference

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_SetLoginTimestamp(ObjMySQL.GetServerTime, vCompName, vUserCode, Me.GetExpiry_String(vLicenseExpiryDate, vProductKey), vProductKey)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjMySQL = Nothing
        End Try
    End Sub

    Public Function GetVersion(ByVal vProductKey As String) As String
        Dim DTA As New DseTrans_Version_ControlTableAdapters.tbltrans_version_controlTableAdapter
        Dim TBL As New DseTrans_Version_Control.tbltrans_version_controlDataTable
        Dim row As DseTrans_Version_Control.tbltrans_version_controlRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_product_key(TBL, vProductKey)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                Return row.version
            Else
                Return "1.6.0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1.6.0"
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub SetVersion(ByVal vProductKey As String, ByVal vVersion As String)
        Dim DTA As New DseTrans_Version_ControlTableAdapters.tbltrans_version_controlTableAdapter
        Dim TBL As New DseTrans_Version_Control.tbltrans_version_controlDataTable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_product_key(TBL, vProductKey)
            If TBL.Rows.Count > 0 Then
                DTA.Update(vVersion, vProductKey)
            Else
                DTA.Insert(vProductKey, vVersion)
            End If
            'MsgBox("Update version di my sql: " + vVersion, MsgBoxStyle.Information, "Information")

            While Me.ApakahNaikVersi(Me.GetOperatingVersion, vVersion) = True
                Me.SetOperatingVersion(vVersion)
            End While
            'MsgBox("Update operating version di my sql: " + vVersion, MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function GetOperatingVersion() As String
        Dim DTA As New DseTrans_Version_ControlTableAdapters.tbltrans_version_operatingTableAdapter
        Dim TBL As New DseTrans_Version_Control.tbltrans_version_operatingDataTable
        Dim row As DseTrans_Version_Control.tbltrans_version_operatingRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            If TBL.Rows.Count = 0 Then
                Return "1.6.40"
            Else
                row = TBL.Rows(0)
                Return row.version
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "1.6.40"
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Sub SetOperatingVersion(ByVal vVersion As String)
        Dim DTA As New DseTrans_Version_ControlTableAdapters.tbltrans_version_operatingTableAdapter
        Dim TBL As New DseTrans_Version_Control.tbltrans_version_operatingDataTable
        Dim row As DseTrans_Version_Control.tbltrans_version_operatingRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL)

            If TBL.Rows.Count = 0 Then
                DTA.Insert(vVersion)
            Else
                row = TBL.Rows(0)
                DTA.Update(vVersion, row.primary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Private Function ApakahNaikVersi(ByVal vVersiLama As String, ByVal vVersiBaru As String) As Boolean
        Dim ParamLama As String() = vVersiLama.Split(".")
        Dim ParamNew As String() = vVersiBaru.Split(".")
        Dim LengthVersion As Integer = vVersiLama.Split(".").Length

        Try
            For k As Integer = 0 To LengthVersion - 1
                If CInt(ParamNew(k)) > CInt(ParamLama(k)) Then
                    Return True
                End If
            Next

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'byahdar utk keperluan update license masal
    Public Function GetLicenseReplace(ByVal vProductKey As String) As String
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetLicenseReplaceFrom(vProductKey)
            If String.IsNullOrEmpty(Retval) = False Then
                Return Retval
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function ApakahKompIsBlacklist(ByVal vProductKey As String) As Boolean
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Dim Retval As New Nullable(Of Integer)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetIsBlacklist(vProductKey)

            If Retval.HasValue = True Then
                If Retval.Value = 1 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetNewLicensiExpiry(ByVal vProductKey As String) As String
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetNewLicenseExpiry(vProductKey)

            If String.IsNullOrEmpty(Retval) = False Then
                Return Me.GetDecrypt(Retval)
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function UpdateLicenseReplace(ByVal LicenseReplace As String, ByVal vProductKey As String) As Boolean
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Dim NewLicenseKey As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_LicenseReplace(LicenseReplace, vProductKey)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function HapusLicenseReplace(ByVal LicenseReplace As String) As Boolean
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Dim NewLicenseKey As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_ReplaceTo(NewLicenseKey, LicenseReplace)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function UpdateOldLicenseKey(ByVal vLicenseExpiryDate As Date, ByVal vProductKey As String) As Boolean
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Dim NewLicenseKey As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.UpdateQuery_OldLicenseKey(Me.GetExpiry_String(vLicenseExpiryDate, vProductKey), vProductKey)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function UpdateNewLicenseKey(ByVal vNewLicenseKey As String, ByVal vProductKey As String) As Boolean
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Dim NewLicenseKey As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            NewLicenseKey = Me.GetEncrypt(vNewLicenseKey)

            DTA.UpdateQuery_NewLicenseKey(NewLicenseKey, vProductKey)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function UpdateBlacklist(ByVal vBlacklistNew As String, ByVal vProductKey As String) As Boolean
        Dim DTA As New DseLicenseManager_UpdateTableAdapters.tbltrans_version_controlTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vBlacklistNew = "1" Then
                DTA.UpdateQuery_NewBlacklist(vBlacklistNew, vProductKey)
            Else
                DTA.UpdateQuery_NewBlacklist(vBlacklistNew, vProductKey)
                HapusLicenseReplace(vProductKey)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetEncrypt(ByVal vLicensiKey As String) As String
        Try
            Return Me.ObjEncrypt.Encrypt(vLicensiKey)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    Public Function GetDecrypt(ByVal vLicensiKeyEncrypt As String) As String
        Try
            Return Me.ObjEncrypt.Decrypt(vLicensiKeyEncrypt)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function
    '''
End Class
