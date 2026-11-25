Option Explicit On
Option Strict Off

Public Class frmCetakSO
    Public Primary As Double = -1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frmCetakSO_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.ShowIcon = False
    End Sub
End Class