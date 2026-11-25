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

    'Public Sub CekForUpdate(ByVal vFrm As FrmListUpdate, ByVal vAppPath As String, ByVal vServerPath As String)
    '    Dim er As String = ""
    '    Dim extension As String = ""
    '    Dim WriteTimeLocal As String = ""
    '    Dim WriteTimeServer As String = ""

    '    Try
    '        ListUpdate.Clear()
    '        'ambil semua list file yang ada di file instalasi
    '        Dim ListFile As String() = Directory.GetFiles(vAppPath)
    '        'ambil list file update dari server
    '        Dim ListFileServer As String() = Directory.GetFiles(vServerPath)
    '        'karena format ListFile = direktori\namafile, maka setiap item di split berdasarkan "\" lalu ambil index terakhirnya
    '        'vFrm.MainProgressBar.Maximum = ListFile.Length
    '        For i As Integer = 1 To ListFile.Length
    '            Dim NamaFile As String() = ListFile(i - 1).Split("\")
    '            'vFrm.Label2.Text = NamaFile(NamaFile.Length - 1)
    '            'nama file nya = NamaFile(NamaFile.Length - 1)
    '            'setelah dapet nama filenya cek apakah di folder update ada yang sama
    '            For j As Integer = 1 To ListFileServer.Length
    '                er = "looping list file di tempat server"
    '                Dim NamaFileServer As String() = ListFileServer(j - 1).Split("\")
    '                If NamaFile(NamaFile.Length - 1) = NamaFileServer(NamaFileServer.Length - 1) Then
    '                    Me.CekLastWriteTime(j, ListFileServer, vAppPath, NamaFile, vServerPath, NamaFileServer)
    '                End If
    '            Next
    '        Next i

    '        'vFrm.MainProgressBar.Value = ListFile.Length
    '        If Me.ListUpdate.Count > 0 Then
    '            vFrm.DgvUpdate.Rows.Clear()
    '            'vFrm.Cursor = Cursors.WaitCursor
    '            For i As Integer = 1 To Me.ListUpdate.Count
    '                Dim NamaFile As String() = Me.ListUpdate(i - 1).Split("\")
    '                vFrm.DgvUpdate.Rows.Add()
    '                Dim rowdgv As DataGridViewRow
    '                rowdgv = vFrm.DgvUpdate.Rows(vFrm.DgvUpdate.Rows.Count - 1)
    '                rowdgv.Cells(vFrm.ColNo.Index).Value = vFrm.DgvUpdate.Rows.Count
    '                rowdgv.Cells(vFrm.ColFile.Index).Value = NamaFile(NamaFile.Length - 1)
    '                rowdgv.Cells(vFrm.ColVer.Index).Value = FileVersionInfo.GetVersionInfo(vServerPath & "\" & NamaFile(NamaFile.Length - 1)).FileVersion
    '            Next i
    '            vFrm.Cursor = Cursors.Default
    '        Else
    '            MsgBox("Application is up to date", MsgBoxStyle.Information, "Information")
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.InnerException, MsgBoxStyle.Critical, "Error " & er.ToString)
    '    End Try
    'End Sub

    Private Sub CekLastWriteTime(ByVal j As Integer, ByVal ListFileServer As String(), ByVal vAppPath As String, ByVal NamaFile As String(), _
    ByVal vServerPath As String, ByVal NamaFileServer As String())
        Dim extension As String = ""
        Dim WriteTimeLocal As String = ""
        Dim WriteTimeServer As String = ""
        Try
            extension = System.IO.Path.GetExtension(vAppPath & "\" & NamaFile(NamaFile.Length - 1))
            If extension = ".dll" Or extension = ".exe" Or extension = ".config" Then
                WriteTimeLocal = System.IO.File.GetLastWriteTime(vAppPath & "\" & NamaFile(NamaFile.Length - 1))
                WriteTimeServer = System.IO.File.GetLastWriteTime(vServerPath & "\" & NamaFileServer(NamaFileServer.Length - 1))
                If WriteTimeServer > WriteTimeLocal Then
                    ListUpdate.Add(ListFileServer(j - 1))
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Public Function GetProductName() As String
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
