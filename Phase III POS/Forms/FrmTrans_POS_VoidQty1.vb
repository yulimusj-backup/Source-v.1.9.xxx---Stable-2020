Option Explicit On
Option Strict Off
Imports Core_Module

Public Class FrmTrans_POS_VoidQty1
    Public BackColorAll As Color = Color.LightGoldenrodYellow
    Public Bahasa As Core_Module.ModGlobalVariable.EnumBahasa = Core_Module.EnumBahasa.Inggris
    Public Parent_Form As FrmTrans_POS
    Public mrClose As Boolean = False
    Public mrOk As Boolean = False

    Dim ObjCD As New ClsCekData

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim Hasil As MsgBoxResult
        Try
            'If String.IsNullOrEmpty(Me.TxtReason.Text) = True Then
            'MsgBox("Please input reason", MsgBoxStyle.Information, "Information")
            'Me.TxtReason.Focus()
            'Exit Sub
            'End If

            Hasil = MsgBox("Void ?", MsgBoxStyle.YesNo, "Confirm")
            If Hasil = MsgBoxResult.Yes Then
                Me.mrClose = False
                Me.mrOk = True
                Me.Close()
            Else
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.mrClose = True
            Me.mrOk = False
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReason_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReason.GotFocus, TxtReason.GotFocus
        Try
            If Me.TxtReason.Enabled = True Then
                If Me.TxtReason.ReadOnly = False Then
                    Me.TxtReason.BackColor = Me.BackColorAll
                    Me.TxtReason.SelectAll()
                Else
                    Me.TxtReason.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TxtReason_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtReason.LostFocus
        Try
            If Me.TxtReason.Enabled = True Then
                If Me.TxtReason.ReadOnly = False Then
                    Me.TxtReason.BackColor = Color.White
                Else
                    Me.TxtReason.BackColor = Me.TxtReadOnly.BackColor
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_POS_VoidQty_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.mrClose = False
            Me.mrOk = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class