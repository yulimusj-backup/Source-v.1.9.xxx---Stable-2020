Option Explicit On
Option Strict Off

Public Class FrmTrans_SPL_Cetak
    Dim Obj As New ClsTrans_SPL
    Public PrimarySpl As Double = -1

    Private Sub FrmTrans_SPLCetak_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub CreatePDF()
        Try
            Me.Obj.Cetak(Me.PrimarySpl, Me.crv, Me.crvDetail, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            If Me.TabControl1.SelectedIndex = 0 Then
                Me.crv.DisplayStatusBar = True
                Me.crv.DisplayToolbar = True
            Else
                Me.crvDetail.DisplayToolbar = True
                Me.crvDetail.DisplayStatusBar = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub
End Class