Option Explicit On
Option Strict Off

Imports System.net
Imports System.Environment
Imports System.Security.Principal

Public Class FrmSettingUpdate
    Dim DariFromClosing As Boolean = False

    Private Sub simpan()
        'Dim dta As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        'Dim Hasil As MsgBoxResult
        'Try
        '    Hasil = MsgBox("Save ?", MsgBoxStyle.YesNo, "Confirm")
        '    If Hasil = MsgBoxResult.Yes Then
        '        dta.Delete()
        '        dta.Insert(Me.TxtDef.Text, Me.TxtPublic.Text, Me.TxtProductName.Text)
        '        'cek folder public
        '        If (Not System.IO.Directory.Exists(Me.TxtPublic.Text)) Then
        '            System.IO.Directory.CreateDirectory(Me.TxtPublic.Text)
        '        End If
        '    Else
        '        If Me.DariFromClosing = False Then
        '            Exit Sub
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    dta.Dispose()
        'End Try
    End Sub

    Private Sub FrmSettingUpdate_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.DariFromClosing = True
            Me.simpan()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmSettingUpdate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Me.GetJumlah() > 0 Then
                Me.TxtDef.Text = Me.GetAppPath()
                Me.TxtPublic.Text = Me.GetServerPath()
                Me.TxtProductName.Text = Me.GetProductName()
            Else
                Me.TxtDef.Text = CurrentDirectory
                Me.TxtPublic.Text = "\\" & Dns.GetHostName & "\Users\Folder Update\" & My.Application.Info.ProductName
                Me.TxtProductName.Text = My.Application.Info.ProductName
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Function GetJumlah() As Decimal
        'Dim DTA As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        'Dim Dummy As System.Nullable(Of Long)
        'Try
        '    Dummy = DTA.ScalarQuery_GetJumlah()
        '    If Dummy.HasValue Then
        '        Return Dummy.Value
        '    Else
        '        Return 0
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        'Finally
        '    DTA.Dispose()
        'End Try
    End Function

    Private Function GetProductName() As String
        'Dim DTA As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        'Dim Dummy As String = ""
        'Try
        '    Dummy = DTA.ScalarQuery_GetProductName()
        '    If String.IsNullOrEmpty(Dummy) = False Then
        '        Return Dummy
        '    Else
        '        Return ""
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        '    Return ""
        'Finally
        '    DTA.Dispose()
        'End Try
        Return ""
    End Function

    Private Function GetAppPath() As String
        'Dim DTA As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        'Dim Dummy As String = ""
        'Try
        '    Dummy = DTA.ScalarQuery_GetAppPath()
        '    If String.IsNullOrEmpty(Dummy) = False Then
        '        Return Dummy
        '    Else
        '        Return ""
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        '    Return ""
        'Finally
        '    DTA.Dispose()
        'End Try
        Return ""
    End Function

    Private Function GetServerPath() As String
        'Dim DTA As New DseAutoUpdateTableAdapters.tbl_setting_updateTableAdapter
        'Dim Dummy As String = ""
        'Try
        '    Dummy = DTA.ScalarQuery_GetServerPath()
        '    If String.IsNullOrEmpty(Dummy) = False Then
        '        Return Dummy
        '    Else
        '        Return ""
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        '    Return ""
        'Finally
        '    DTA.Dispose()
        'End Try
        Return ""
    End Function

    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
        Me.Close()
    End Sub
End Class