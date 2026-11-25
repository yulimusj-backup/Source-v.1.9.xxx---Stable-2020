Option Explicit On
Option Strict Off

Imports System.Windows.Forms, System.Drawing
Public Class FrmTrans_Gaji_MasaSblm_Cetak
    Dim ObjCD As New Core_Module.ClsCekData
    Dim Obj As New ClsTrans_Gaji_MasaSblm
    Public PrimaryTrans As Double = -1

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click

        Try
            If Me.dtpFrom.Value.Date > Me.dtpTo.Value.Date Then
                MsgBox("Date from cannot be smaller than date to.", MsgBoxStyle.Information, "Information")
                Me.dtpTo.Focus()

                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            Me.Obj.Cetak(Me.PrimaryTrans, Me.dtpFrom.Value.Date, Me.dtpTo.Value.Date, Me.crvLaporanAnalysisOvertime)
            Me.Cursor = Cursors.Default
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub FrmTrans_Gaji_MasaSblm_Cetak_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub
End Class