Option Explicit On
Option Strict Off

Public Class FrmTransaksiTagihan_Cetak

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub


    Private Sub FrmTransaksiTagihan_Cetak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error")
        End Try
    End Sub
End Class