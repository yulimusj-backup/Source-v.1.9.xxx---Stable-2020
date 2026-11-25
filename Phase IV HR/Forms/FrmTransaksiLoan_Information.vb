Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class FrmTransaksiLoan_Information

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

End Class