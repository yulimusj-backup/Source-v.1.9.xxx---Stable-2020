Option Explicit On
Option Strict Off

Public Class FrmMasterKaryawan_Cetak1
    Public mPrimaryKaryawan As Double = -1
    Dim ObjKaryawan As New Core_Module.ClsMasterKaryawan

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub FrmTransaksiJournal_Cetak1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.ObjKaryawan.CetakanDataKaryawan(Me.crvReport, Me.mPrimaryKaryawan)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class