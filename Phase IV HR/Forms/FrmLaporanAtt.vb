''by    : rosadin
''date  : 2014-08-08

Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmLaporanAtt

    Dim obj As New ClsTrans_Attendance_laporan

    Private Sub FrmLaporanAtt_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.CboFomat.SelectedIndex = 0
            Me.ShowIcon = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub LblSetting_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LblSetting.LinkClicked
        Dim Frm As New FrmSettingLeave
        Try
            Frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub BtnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShow.Click
        Try
            If dtpFrom.Value > dtpTo.Value Then
                MsgBox("Start date value can't greater than end date value", MsgBoxStyle.Exclamation, "Exclamation")
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            If Me.CboFomat.SelectedIndex = 0 Then
                Me.obj.cetak(Me.dtpFrom.Value, Me.dtpTo.Value, Me.CrvReport)
            Else
                Me.obj.cetak_detail(Me.dtpFrom.Value, Me.dtpTo.Value, Me.CrvReport)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub dtpTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpTo.KeyDown
        Try
            If e.KeyValue = 13 Then
                BtnShow.PerformClick()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub
End Class