Option Explicit On
Option Strict Off

Imports System.net
Imports System.Environment
Imports System.IO

Public Class ClsAutoUpdate
    Public ListUpdate As New List(Of String)

    Public Function checkURL(ByVal vUrl As String) As String
        Dim s As String = vUrl.ToLower()
        If (s.StartsWith("http://www.") Or s.StartsWith("https://www.")) Then
            Return s
        ElseIf (s.StartsWith("http://")) Then
            If (s.Contains("www.")) Then
                s = "http://www." & s.Substring(1, s.Length - 1)
                vUrl = s
                Return s
            Else
                s = "http://www." & s.Substring(7, s.Length - 7)
                vUrl = s
                Return s
            End If
        ElseIf (s.StartsWith("https://")) Then
            If (s.Contains("www.")) Then
                s = "https://" & s.Substring(12, s.Length - 12)
                vUrl = s
                Return s
            Else
                s = "https://www." & s.Substring(8, s.Length - 8)
                vUrl = s
                Return s
            End If
        Else
            If (s.StartsWith("www.")) Then
                s = "http://www." & s.Substring(4, s.Length - 4)
                vUrl = s
                Return s
            Else
                s = "http://www." & s
                vUrl = s
                Return s
            End If
        End If
    End Function

    Public Function GetAppPath() As String
        Return My.Application.Info.DirectoryPath
    End Function

    Public Function GetServerPath() As String
        Dim dta As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        Try
            Return dta.ScalarQuery_GetServerPath
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetProductName1() As String
        Return My.Application.Info.ProductName
    End Function

    Public Function GetJumlah() As Decimal
        Dim DTA As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        Dim Dummy As System.Nullable(Of Long)
        Try
            Dummy = DTA.ScalarQuery_GetJumlah()
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
