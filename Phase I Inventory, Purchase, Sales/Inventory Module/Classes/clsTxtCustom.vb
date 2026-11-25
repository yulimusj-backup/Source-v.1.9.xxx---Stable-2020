Option Explicit On
Option Strict Off

Public Class clsTxtCustom
    Inherits TextBox
    Dim BackColorDummy As Color = Me.BackColor

    Private Sub clsTxtCustom_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        Try
            If Me.ReadOnly = False Then
                Me.BackColor = Core_Module.GetSMESSBackColorAll
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub clsTxtCustom_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
        Try
            If Me.ReadOnly = False Then
                Me.BackColor = BackColorDummy
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class
