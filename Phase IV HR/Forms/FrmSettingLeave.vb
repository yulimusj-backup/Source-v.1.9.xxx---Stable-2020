''by    : rosadin
''date  : 2014-08-07
Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Public Class FrmSettingLeave

    Dim obj As New ClsSettingLeave

    Private Sub FrmSettingLeave_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.obj.isiDgv(Me.dgvLeaveCode, Me, Me.dtpPeriode.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        Finally

        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    Private Sub dtpPeriode_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPeriode.ValueChanged
        Try
            Me.Cursor = Cursors.WaitCursor
            Me.obj.hapusSettingLeave(Me.dgvLeaveCode, Me)
            Me.obj.simpanSettingLeave(Me.dgvLeaveCode, Me)
            Me.obj.isiDgv(Me.dgvLeaveCode, Me, Me.dtpPeriode.Value)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOk.Click
        Try
            If MsgBox("Save data and close?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Save") = MsgBoxResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                Me.obj.hapusSettingLeave(Me.dgvLeaveCode, Me)
                Me.obj.simpanSettingLeave(Me.dgvLeaveCode, Me)
                Me.obj.isiDgv(Me.dgvLeaveCode, Me, Me.dtpPeriode.Value)
                Me.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Information")
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
End Class