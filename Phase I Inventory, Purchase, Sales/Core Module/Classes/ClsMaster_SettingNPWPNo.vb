Option Explicit On
Option Strict Off

Public Class ClsMaster_SettingNPWPNo

    Public Enum TipeNPWP
        Company
        Individual
    End Enum

    Public Function ApakahDataSudahBenar_MasterCustomer(ByVal vNPWPNo As String, ByVal vNIKOrPassword As String, ByVal vNPWPLama As String, _
    ByVal vCboNPWPType As ComboBox, ByVal vTampilkanNotif As Boolean) As Boolean

        Try

            If Me.ApakahSudah16Digit(vNPWPNo, True) = False Then
                Return False
            End If

            If Me.ApakahNPWPNoDanNIKSama(vNPWPNo, vNIKOrPassword, True) = False Then
                Return False
            End If

            If String.IsNullOrEmpty(vNPWPLama) = True Then
                MsgBox("Old NPWP No cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            If vCboNPWPType.SelectedIndex = -1 Then
                MsgBox("NPWP Type cannot empty string.", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudah16Digit(ByVal vNPWPNo As String, ByVal vTampilkanNotif As Boolean) As Boolean
        Dim Retval As String = ""
        Try
            If String.IsNullOrEmpty(vNPWPNo) = False Then
                Retval = vNPWPNo.ToString.Replace(".", "").Replace("-", "")
                If Retval.Length = 16 Then
                    Return True
                End If
            End If
            If vTampilkanNotif = True Then
                MsgBox("Used 16 digit for NPWP No.", MsgBoxStyle.Information, "Information")
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahNPWPNoDanNIKSama(ByVal vNPWPNo As String, ByVal vNIKOrPassword As String, ByVal vTampilkanNotif As Boolean) As Boolean
        Try
            If String.IsNullOrEmpty(vNPWPNo) = False And String.IsNullOrEmpty(vNIKOrPassword) = False Then
                If vNPWPNo = vNIKOrPassword Then
                    Return True
                End If
            End If
            If vTampilkanNotif = True Then
                MsgBox("NPWP No and NIK/PASSPORT must be identical", MsgBoxStyle.Information, "Information")
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function UpdateNPWP_MasterCustomer(ByVal vNPWPNo As String, ByVal vNPWPNoOld As String, _
    ByVal vNPWPType As ClsMaster_SettingNPWPNo.TipeNPWP, ByVal vPrimaryCustomer As Double) As Boolean
        'fungis ini mengupdate NPWP, NPWP Lama dan NPWP Type pd Master Customer.
        Dim DTA As New DseTrans_SettingNPWPTableAdapters.tblmaster_customerTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_CustomerData(vNPWPNo, vNPWPNoOld, vNPWPType, vPrimaryCustomer)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNPWPType_MasterCustomer(ByVal vPrimaryCustomer As Double) As ClsMaster_SettingNPWPNo.TipeNPWP

        Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try


    End Function
End Class
